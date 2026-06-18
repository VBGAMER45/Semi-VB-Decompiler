Attribute VB_Name = "modCOM"
'*****************************
'modCom.bas
'Purpose to Retrive the members and variable types of a control
'*****************************
Option Explicit
Global tliTypeLibInfo As TypeLibInfo
'--- Late-bound (__vbaLateIdCall) member resolution: DISPID -> name via the OCX typelib ---
Private mLateTlbCache As Collection   'OCX file path -> clsTypeLibInfo (each loaded once)
Private mOcxEvtTlbCache As Collection 'OCX file path -> TypeLibInfo (event sig resolution)

Private Function OpenOcxTlb(ByVal ocxFile As String) As clsTypeLibInfo
    'Load ocxFile's typelib once (cached) - shared by the late-bound (dispid) and
    'early-bound (vtable offset) resolvers.
    On Error GoTo fail
    If Len(ocxFile) = 0 Then Exit Function
    If mLateTlbCache Is Nothing Then Set mLateTlbCache = New Collection
    Dim tlb As clsTypeLibInfo
    On Error Resume Next
    Set tlb = mLateTlbCache(ocxFile)
    On Error GoTo fail
    If tlb Is Nothing Then
        Set tlb = New clsTypeLibInfo
        If Not tlb.OpenTypeLib(ocxFile) Then Exit Function
        mLateTlbCache.Add tlb, ocxFile
    End If
    Set OpenOcxTlb = tlb
fail:
End Function

Private Function CoclassIfaceGuids(tlb As clsTypeLibInfo, ByVal ctlClass As String) As String
    'The "|guid|guid|" set of interfaces implemented by the coclass named ctlClass
    '(e.g. "MSComctlLib.Slider" -> coclass "Slider" -> ISlider/_Slider...).  Used to
    'scope member resolution to ONE control's own interfaces: an OCX typelib reuses the
    'same dispid/vtable offset across unrelated controls (MSCOMCTL dispid 0xB is
    'ITabStrip.ClientHeight AND ISlider.Value), so an un-scoped scan mis-resolves.  ""
    'when the class is unknown / not a coclass here (caller then falls back to all).
    On Error GoTo done
    If Len(ctlClass) = 0 Then Exit Function
    Dim shortClass As String
    shortClass = ctlClass
    If InStr(shortClass, ".") > 0 Then shortClass = Mid$(shortClass, InStrRev(shortClass, ".") + 1)
    Dim i As Long, j As Long, res As String
    For i = 0 To tlb.TypeInfoCount - 1
        tlb.SelectTypeInfo i
        If tlb.TypeInfoKind = TKIND_COCLASS And StrComp(tlb.TypeInfoName, shortClass, vbTextCompare) = 0 Then
            For j = 0 To tlb.TypeInfoImplements - 1
                tlb.SelectImplement j
                res = res & "|" & tlb.ImplementGUID
            Next j
            Exit For
        End If
    Next i
    If Len(res) > 0 Then CoclassIfaceGuids = res & "|"
done:
End Function

Public Function ResolveLateMember(ByVal ocxFile As String, ByVal dispid As Long, ByVal wantKind As Long, ByRef invKind As Long, Optional ByVal ctlClass As String) As String
    'Load ocxFile's typelib (cached) and return the member NAME whose memid = dispid.
    'wantKind (1=Func, 2=Get, 4=Put, 0=any) breaks ties: a control typelib can carry
    'the SAME memid in two dispinterfaces (e.g. TABCTL32 has 0x4 = IVBDataObject.SetData
    '[Func] AND ISSTabCtl.Tab [Put]); a property STORE must pick the Put. Returns the
    'best (wantKind) match, else the first match.  "" if none.  When ctlClass is given,
    'the scan is SCOPED to that coclass's interfaces - vital because a dispid collides
    'across MSCOMCTL controls (0xB = ITabStrip.ClientHeight AND ISlider.Value), so an
    'un-scoped scan mis-resolved Slider .Value to ClientHeight.
    On Error GoTo done
    Dim tlb As clsTypeLibInfo
    Set tlb = OpenOcxTlb(ocxFile)
    If tlb Is Nothing Then Exit Function
    Dim scope As String
    scope = CoclassIfaceGuids(tlb, ctlClass)        '"" => unscoped (legacy behaviour)
    Dim i As Long, g As Long, anyName As String, anyKind As Long
    For i = 0 To tlb.TypeInfoCount - 1
        tlb.SelectTypeInfo i
        If (tlb.TypeInfoKind = TKIND_DISPATCH Or tlb.TypeInfoKind = TKIND_INTERFACE) _
           And (Len(scope) = 0 Or InStr(scope, "|" & tlb.TypeInfoGUID & "|") > 0) Then
            For g = 0 To tlb.TypeInfoFunctions - 1
                If tlb.SelectFunction(g) Then
                    If tlb.FunctionMemberId = dispid Then
                        If wantKind <> 0 And (tlb.FunctionInvKind And wantKind) <> 0 Then
                            ResolveLateMember = tlb.FunctionName       'exact invoke-kind match
                            invKind = tlb.FunctionInvKind
                            Exit Function
                        ElseIf Len(anyName) = 0 Then
                            anyName = tlb.FunctionName: anyKind = tlb.FunctionInvKind  'remember first
                        End If
                    End If
                End If
            Next g
        End If
    Next i
    ResolveLateMember = anyName: invKind = anyKind
done:
End Function

Public Function OcxFileFromClsid(ByVal clsid As String) As String
    'On-disk file (typelib container) for a coclass CLSID via the registry: prefer the
    'InprocServer32 .ocx; else the TypeLib GUID + Version -> the registered win32 path.
    On Error Resume Next
    If Len(clsid) = 0 Then Exit Function
    Dim f As String
    f = modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & clsid & "\InprocServer32", "")
    If Len(f) > 0 And InStr(f, "\") > 0 Then OcxFileFromClsid = f: Exit Function
    Dim tlbGuid As String, ver As String
    tlbGuid = modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & clsid & "\TypeLib", "")
    ver = modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & clsid & "\Version", "")
    If Len(tlbGuid) > 0 And Len(ver) > 0 Then
        OcxFileFromClsid = modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\TypeLib\" & tlbGuid & "\" & ver & "\0\win32", "")
    End If
End Function

Public Function LateMemberName(ByVal libClass As String, ByVal ctlBase As String, ByVal dispid As Long, ByVal wantKind As Long, ByRef invKind As Long) As String
    'Resolve a late-bound member (dispid) on an OCX control to its name.  Prefer the
    'control's EXACT external class (libClass, e.g. "TabDlg.SSTab" from gControlClass);
    'else its base name (Winsock1 -> "Winsock" -> MSWinsockLib.Winsock).  Deliberately
    'does NOT blind-scan every OCX - that mis-resolved e.g. MFD.Tab to a same-dispid
    'Winsock property when the control name gave no hint.
    On Error Resume Next
    Dim i As Long, f As String, nm As String
    If UBound(gOcxList) < 1 Then Exit Function
    If Len(libClass) > 0 Then                         'pass 1: exact external class
        For i = 0 To UBound(gOcxList) - 1
            If Len(gOcxList(i).strGuid) > 0 And StrComp(gOcxList(i).strLibname, libClass, vbTextCompare) = 0 Then
                f = OcxFileFromClsid(gOcxList(i).strGuid)
                nm = ResolveLateMember(f, dispid, wantKind, invKind, libClass)
                If Len(nm) > 0 Then LateMemberName = nm: Exit Function
            End If
        Next i
    End If
    If Len(ctlBase) > 0 Then                          'pass 2: OCX whose name matches the control base
        For i = 0 To UBound(gOcxList) - 1
            If Len(gOcxList(i).strGuid) > 0 Then
                If InStr(1, gOcxList(i).strLibname, ctlBase, vbTextCompare) > 0 _
                   Or InStr(1, gOcxList(i).strocxName, ctlBase, vbTextCompare) > 0 Then
                    f = OcxFileFromClsid(gOcxList(i).strGuid)
                    nm = ResolveLateMember(f, dispid, wantKind, invKind, gOcxList(i).strLibname)
                    If Len(nm) > 0 Then LateMemberName = nm: Exit Function
                End If
            End If
        Next i
    End If
End Function

'--- Early-bound (vtable) member resolution: vtable offset -> name via the OCX typelib ---
Public Function ResolveVtableMember(ByVal ocxFile As String, ByVal ctlClass As String, ByVal vtOffset As Long, ByVal wantKind As Long, ByRef invKind As Long) As String
    'Load ocxFile's typelib (cached) and return the member NAME whose VTABLE OFFSET
    '(oVft) = vtOffset, on the interface(s) implemented by the control's coclass.  The
    'twin of ResolveLateMember for EARLY-bound calls (`call [vt+off]`): VB6 OCX controls
    '(MSCOMCTL Slider/StatusBar, etc.) are accessed through their dual interface's vtable,
    'not late-bound dispatch.  Scoped to the coclass's own interfaces (ctlClass, e.g.
    '"MSComctlLib.Slider" -> coclass "Slider") so a vtable offset shared across unrelated
    'interfaces can't mis-resolve.  wantKind (1=Func,2=Get,4=Put,0=any) breaks ties.
    On Error GoTo done
    Dim tlb As clsTypeLibInfo
    Set tlb = OpenOcxTlb(ocxFile)
    If tlb Is Nothing Then Exit Function
    Dim implGuids As String
    implGuids = CoclassIfaceGuids(tlb, ctlClass)
    If Len(implGuids) = 0 Then GoTo done       'coclass not found in this typelib
    'Scan the coclass's interfaces for a function at the wanted vtable offset.
    Dim i As Long, j As Long, anyName As String, anyKind As Long
    For i = 0 To tlb.TypeInfoCount - 1
        tlb.SelectTypeInfo i
        If (tlb.TypeInfoKind = TKIND_DISPATCH Or tlb.TypeInfoKind = TKIND_INTERFACE) _
           And InStr(implGuids, "|" & tlb.TypeInfoGUID & "|") > 0 Then
            For j = 0 To tlb.TypeInfoFunctions - 1
                If tlb.SelectFunction(j) Then
                    If tlb.FunctionVTOffset = vtOffset Then
                        If wantKind <> 0 And (tlb.FunctionInvKind And wantKind) <> 0 Then
                            ResolveVtableMember = tlb.FunctionName: invKind = tlb.FunctionInvKind: Exit Function
                        ElseIf Len(anyName) = 0 Then
                            anyName = tlb.FunctionName: anyKind = tlb.FunctionInvKind
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    ResolveVtableMember = anyName: invKind = anyKind
done:
End Function

Public Function VtableMemberName(ByVal libClass As String, ByVal ctlBase As String, ByVal vtOffset As Long, ByVal wantKind As Long, ByRef invKind As Long) As String
    'Resolve an early-bound vtable member on an OCX control to its name, by the control's
    'external class (libClass, e.g. "MSComctlLib.Slider" from gControlClass).  Mirrors
    'LateMemberName: find the gOcxList entry for the class, get its OCX file, resolve the
    'offset on that coclass's interfaces.  ctlBase (the control's base name) is a weak
    'fallback when the external class gave no hit.
    On Error Resume Next
    Dim i As Long, f As String, nm As String
    If UBound(gOcxList) < 1 Then Exit Function
    If Len(libClass) > 0 Then                          'pass 1: exact external class
        For i = 0 To UBound(gOcxList) - 1
            If Len(gOcxList(i).strGuid) > 0 And StrComp(gOcxList(i).strLibname, libClass, vbTextCompare) = 0 Then
                f = OcxFileFromClsid(gOcxList(i).strGuid)
                nm = ResolveVtableMember(f, libClass, vtOffset, wantKind, invKind)
                If Len(nm) > 0 Then VtableMemberName = nm: Exit Function
            End If
        Next i
    End If
    If Len(ctlBase) > 0 Then                            'pass 2: OCX whose name matches the control base
        For i = 0 To UBound(gOcxList) - 1
            If Len(gOcxList(i).strGuid) > 0 Then
                If InStr(1, gOcxList(i).strLibname, ctlBase, vbTextCompare) > 0 _
                   Or InStr(1, gOcxList(i).strocxName, ctlBase, vbTextCompare) > 0 Then
                    f = OcxFileFromClsid(gOcxList(i).strGuid)
                    nm = ResolveVtableMember(f, gOcxList(i).strLibname, vtOffset, wantKind, invKind)
                    If Len(nm) > 0 Then VtableMemberName = nm: Exit Function
                End If
            End If
        Next i
    End If
End Function

Public Function GetSearchType(ByVal SearchData As Long) As TliSearchTypes
    'This helper function adapted from Microsoft documentation
    If SearchData And &H80000000 Then
        GetSearchType = ((SearchData And &H7FFFFFFF) \ &H1000000 And &H7F&) Or &H80
    Else
        GetSearchType = SearchData \ &H1000000 And &HFF&
    End If
End Function


Public Function ProduceDefaultValue(DefVal As Variant, ByVal tliTypeInfo As TypeInfo) As String
'This helper function adapted from Microsoft documentation
Dim lngTrackVal As Long
Dim MI As MemberInfo
Dim tliTypeKinds As TypeKinds
    
If tliTypeInfo Is Nothing Then
    Select Case VarType(DefVal)
        Case vbString
            If Len(DefVal) Then
                ProduceDefaultValue = """" & DefVal & """"
            End If
        Case vbBoolean 'Always show for Boolean
            ProduceDefaultValue = DefVal
        Case vbDate
            If DefVal Then
                ProduceDefaultValue = "#" & DefVal & "#"
            End If
        Case Else 'Numeric Values
            If DefVal <> 0 Then
                ProduceDefaultValue = DefVal
            End If
    End Select
Else
    'Resolve constants to their enums
    tliTypeKinds = tliTypeInfo.TypeKind
    Do While tliTypeKinds = TKIND_ALIAS
        tliTypeKinds = TKIND_MAX
        On Error Resume Next
        Set tliTypeInfo = tliTypeInfo.ResolvedType
        If err = 0 Then
            tliTypeKinds = tliTypeInfo.TypeKind
        End If
        On Error GoTo 0
    Loop
    If tliTypeInfo.TypeKind = TKIND_ENUM Then
        lngTrackVal = DefVal
        For Each MI In tliTypeInfo.Members
            If MI.value = lngTrackVal Then
                ProduceDefaultValue = " = " & MI.name
                Exit For
            End If
        Next
    End If
End If
End Function

Public Function ReturnGuiOpcode(ByVal SearchData As Long, _
    ByVal InvokeKinds As InvokeKinds, _
    Optional ByVal MemberName As String) As Integer
'*****************************
'Purpose: To return the opcode of a property used in form decompiling
'*****************************
On Error GoTo exitFunction
    Dim Num As Integer
    With tliTypeLibInfo
        
        With .GetMemberInfo(SearchData, InvokeKinds, , MemberName)
            'Debug.Print "MemberID: 0x" & Hex(.MemberId - &H10000)
        
            Num = (.MemberId - 65536)
        End With
     End With
     
     If Num > 255 Then Num = -1
        
     ReturnGuiOpcode = Num
     Exit Function
exitFunction:
    ReturnGuiOpcode = -1
Exit Function
End Function
Public Function ReturnDataType(ByVal SearchData As Long, _
    ByVal InvokeKinds As InvokeKinds, _
    Optional ByVal MemberName As String) As String
'*****************************
'Purpose: To return the data type of a property
'*****************************
    On Error GoTo exitFunction

    Dim bIsConstant As Boolean
    Dim strReturn As String
    Dim ConstVal As Variant
    Dim strTypeName As String
    Dim intVarTypeCur As Integer
    

  
    With tliTypeLibInfo
        
        'First, determine the type of member we're dealing with
        bIsConstant = GetSearchType(SearchData) And tliStConstants
        With .GetMemberInfo(SearchData, InvokeKinds, , MemberName)

        
            If bIsConstant Then
                ConstVal = .value
                strReturn = strReturn & " = " & ConstVal
                Select Case VarType(ConstVal)
                    Case vbInteger, vbLong
                        If ConstVal < 0 Or ConstVal > 15 Then
                            strReturn = strReturn & " (&H" & Hex$(ConstVal) & ")"
                        End If
                End Select
            Else
                With .ReturnType
                    intVarTypeCur = .VarType
                    If intVarTypeCur = 0 Or (intVarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                        On Error Resume Next
                        If Not .TypeInfo Is Nothing Then
                            If err Then 'Information not available
                                strReturn = strReturn & " As ?"
                            Else
                                If .IsExternalType Then
                                    strReturn = strReturn & .TypeLibInfoExternal.name & "." & .TypeInfo.name
                                Else
                                    strReturn = strReturn & .TypeInfo.name
                                End If
                            End If
                        End If
                        
                        If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                            strReturn = strReturn & "()"
                        End If
                        On Error GoTo 0
                    Else
                        Select Case intVarTypeCur
                            Case VT_VARIANT, VT_VOID, VT_HRESULT
                            Case Else
                                strTypeName = TypeName(.TypedVariant)
                                If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                    strReturn = strReturn & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                                Else
                                    strReturn = strReturn & strTypeName
                                End If
                        End Select
                    End If
                End With
            End If
            
            ReturnDataType = strReturn & vbCrLf

        End With
    End With
exitFunction:
    
End Function

Public Sub ProcessTypeLibrary()
'*****************************
'Purpose: Procces the Type Libary
'*****************************
    'Clear lists
    frmMain.lstTypeInfos.Clear
    frmMain.lstMembers.Clear
    
    'Display members for type library
    tliTypeLibInfo.GetTypesDirect frmMain.lstTypeInfos.hwnd, , tliStAll
End Sub

Public Function getEventInfo(MI As MemberInfo, ObjectName As String, ShowOpcode As Boolean) As String
'*****************************
'Purpose: Get a specific event information
'*****************************
Dim sOutput As String, strTypeName As String, ConstVal As String
Dim lSearchData As Long
Dim bIsConstant As Boolean, bDefault As Boolean, bFirstParameter As Boolean
Dim bParamArray As Boolean, bOptional As Boolean, bByVal As Boolean
Dim tliParameterInfo As ParameterInfo
Dim tliTypeInfo As TypeInfo, tliResolvedTypeInfo As TypeInfo
Dim tliTypeKinds As TypeKinds
Dim intVarTypeCur As Integer
            With MI
                If ShowOpcode = True Then
                sOutput = sOutput & .VTableOffset
                
                End If
                bIsConstant = GetSearchType(lSearchData) And tliStConstants
                
                sOutput = sOutput & .name
                With .Parameters
                    If .count Then
                        sOutput = sOutput & " ("
                        bFirstParameter = True
                        bParamArray = .OptionalCount = -1
                        For Each tliParameterInfo In .Me
                            If Not bFirstParameter Then
                                sOutput = sOutput & ", "
                            End If
                            bFirstParameter = False
                            bDefault = tliParameterInfo.Default
                            bOptional = bDefault Or tliParameterInfo.Optional
                            If bOptional Then
                                If bParamArray Then
                                    'This will be the only optional parameter
                                    sOutput = sOutput & "[ParamArray "
                                Else
                                    sOutput = sOutput & "["
                                End If
                            End If
                        
                            With tliParameterInfo.VarTypeInfo
                                Set tliTypeInfo = Nothing
                                Set tliResolvedTypeInfo = Nothing
                                tliTypeKinds = TKIND_MAX
                                intVarTypeCur = .VarType
                                If (intVarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                                    On Error Resume Next
                                    Set tliTypeInfo = .TypeInfo
                                    If Not tliTypeInfo Is Nothing Then
                                        Set tliResolvedTypeInfo = tliTypeInfo
                                        tliTypeKinds = tliResolvedTypeInfo.TypeKind
                                        Do While tliTypeKinds = TKIND_ALIAS
                                            tliTypeKinds = TKIND_MAX
                                            Set tliResolvedTypeInfo = tliResolvedTypeInfo.ResolvedType
                                            If err Then
                                                err.Clear
                                            Else
                                                tliTypeKinds = tliResolvedTypeInfo.TypeKind
                                            End If
                                        Loop
                                    End If
                                
                                    'Determine whether parameters are ByVal or ByRef
                                    Select Case tliTypeKinds
                                        Case TKIND_INTERFACE, TKIND_COCLASS, TKIND_DISPATCH
                                            bByVal = .PointerLevel = 1
                                        Case TKIND_RECORD
                                            'Records not passed ByVal in VB
                                            bByVal = False
                                        Case Else
                                            bByVal = .PointerLevel = 0
                                    End Select
                                
                                    'Indicate ByVal
                                    If bByVal Then
                                        sOutput = sOutput & "ByVal "
                                    End If
                                
                                    'Display the parameter name
                                    sOutput = sOutput & tliParameterInfo.name
                                
                                    If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                        sOutput = sOutput & "()"
                                    End If
                                    
                                    If tliTypeInfo Is Nothing Then 'Information not available
                                        sOutput = sOutput & " As ?"
                                    Else
                                        If .IsExternalType Then
                                            sOutput = sOutput & " As " & .TypeLibInfoExternal.name & "." & tliTypeInfo.name
                                        Else
                                            sOutput = sOutput & " As " & tliTypeInfo.name
                                        End If
                                    End If
                                
                                    'Reset error handling
                                    On Error GoTo 0
                                Else
                                    If .PointerLevel = 0 Then
                                        sOutput = sOutput & "ByVal "
                                    End If
                                        
                                    sOutput = sOutput & tliParameterInfo.name
                                    If intVarTypeCur <> vbVariant Then
                                        strTypeName = TypeName(.TypedVariant)
                                        If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                            sOutput = sOutput & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                                        Else
                                            sOutput = sOutput & " As " & strTypeName
                                        End If
                                    End If
                                End If
                                    
                                If bOptional Then
                                    If bDefault Then
                                        sOutput = sOutput & ProduceDefaultValue(tliParameterInfo.DefaultValue, tliResolvedTypeInfo)
                                        'sOutput = sOutput & " = " & tliParameterInfo.DefaultValue
                                    End If
                                    sOutput = sOutput & "]"
                                End If
                            End With
                        Next
                        sOutput = sOutput & ")"
                    End If
                End With
                'return type
                If bIsConstant Then
                    ConstVal = .value
                    sOutput = sOutput & " = " & ConstVal
                    Select Case VarType(ConstVal)
                        Case vbInteger, vbLong
                            If ConstVal < 0 Or ConstVal > 15 Then
                                sOutput = sOutput & " (&H" & Hex$(ConstVal) & ")"
                            End If
                    End Select
                Else
                    With .ReturnType
                        intVarTypeCur = .VarType
                        If intVarTypeCur = 0 Or (intVarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                            On Error Resume Next
                            If Not .TypeInfo Is Nothing Then
                                If err Then 'Information not available
                                    sOutput = sOutput & " As ?"
                                Else
                                    If .IsExternalType Then
                                        sOutput = sOutput & " As " & .TypeLibInfoExternal.name & "." & .TypeInfo.name
                                    Else
                                        sOutput = sOutput & " As " & .TypeInfo.name
                                    End If
                                End If
                            End If
                            
                            If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                sOutput = sOutput & "()"
                            End If
                            On Error GoTo 0
                        Else
                            Select Case intVarTypeCur
                                Case VT_VARIANT, VT_VOID, VT_HRESULT
                                Case Else
                                    strTypeName = TypeName(.TypedVariant)
                                    If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                        sOutput = sOutput & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                                    Else
                                        sOutput = sOutput & " As " & strTypeName
                                    End If
                            End Select
                        End If
                    End With
                End If
            End With
        getEventInfo = sOutput
End Function

Public Function ReturnHelpString(ByVal SearchData As Long, ByVal InvokeKinds As InvokeKinds, Optional ByVal MemberName As String) As String
'*****************************
'Purpose: To return the help string used on textbox in form editor to describe function
'*****************************
    With tliTypeLibInfo
        'First, determine the type of member we're dealing with
        With .GetMemberInfo(SearchData, InvokeKinds, , MemberName)
            ReturnHelpString = .HelpString
        End With
    End With

End Function
Public Function OcxEventSig(ByVal ctlName As String, ByVal k As Integer, ByVal EventCount As Long) As String
'*****************************
'Purpose: Resolve a control event handler's name+signature when the event source
'   is an OCX's (not in VB6.OLB, so GetEventNumber returns -1). The form stores a
'   VB-synthesized sink IID for OCX controls (not the OCX's event dispinterface
'   IID), so we map the control to its OCX by NAME (the late-bound resolver's
'   heuristic - Winsock1 -> "Winsock") and read the coclass's DefaultEventInterface.
'
'   The hooked events occupy the LAST <member-count> slots of the control's
'   EventCount-slot table, in declaration order, so the 1-based member index is:
'       memNum = k - (EventCount - memberCount) + 1
'   verified against MSWINSCK.OCX (Winsock EventCount=16, 7 events -> k 9/10/11 =
'   Error/DataArrival/Connect). Returns e.g. "DataArrival (ByVal bytesTotal As
'   Long)" formatted by getEventInfo, or "" if unresolved.
'*****************************
    On Error GoTo done
    If UBound(gOcxList) < 1 Then Exit Function
    If mOcxEvtTlbCache Is Nothing Then Set mOcxEvtTlbCache = New Collection
    Dim base As String
    base = StripTrailingDigits(ctlName)
    If Len(base) = 0 Then Exit Function
    Dim i As Long
    For i = 0 To UBound(gOcxList) - 1
        If Len(gOcxList(i).strGuid) > 0 And Len(gOcxList(i).strLibname) > 0 Then
            Dim coName As String
            coName = CoClassPart(gOcxList(i).strLibname)      '"MSWinsockLib.Winsock" -> "Winsock"
            If StrComp(coName, base, vbTextCompare) = 0 Or InStr(1, gOcxList(i).strLibname, base, vbTextCompare) > 0 Then
                Dim f As String
                f = OcxFileFromClsid(gOcxList(i).strGuid)
                If Len(f) > 0 Then
                    Dim sig As String
                    sig = OcxEventSigFromTlb(f, coName, k, EventCount)
                    If Len(sig) > 0 Then OcxEventSig = sig: Exit Function
                End If
            End If
        End If
    Next
done:
End Function

Private Function OcxEventSigFromTlb(ByVal f As String, ByVal coName As String, ByVal k As Integer, ByVal EventCount As Long) As String
    On Error GoTo done
    Dim tl As TypeLibInfo
    Set tl = Nothing
    On Error Resume Next
    Set tl = mOcxEvtTlbCache(f)
    On Error GoTo done
    If tl Is Nothing Then
        Set tl = New TypeLibInfo
        tl.ContainingFile = f
        On Error Resume Next
        mOcxEvtTlbCache.Add tl, f
        On Error GoTo done
    End If
    Dim cc As CoClassInfo, ci As Long
    Set cc = Nothing
    For ci = 1 To tl.CoClasses.count
        If StrComp(tl.CoClasses.Item(ci).name, coName, vbTextCompare) = 0 Then
            Set cc = tl.CoClasses.Item(ci)
            Exit For
        End If
    Next
    If cc Is Nothing Then Exit Function
    Dim ev As InterfaceInfo
    Set ev = cc.DefaultEventInterface
    If ev Is Nothing Then Exit Function
    Dim memNum As Long
    memNum = (k - (EventCount - ev.Members.count)) + 1        '1-based member index
    If memNum >= 1 And memNum <= ev.Members.count Then
        OcxEventSigFromTlb = getEventInfo(ev.Members.Item(memNum), "", False)
    End If
done:
End Function

'Strip trailing digits from a control name (Winsock1 -> Winsock) for OCX matching.
Private Function StripTrailingDigits(ByVal s As String) As String
    Dim n As Long
    n = Len(s)
    Do While n > 0
        Dim c As Integer
        c = Asc(Mid$(s, n, 1))
        If c < 48 Or c > 57 Then Exit Do
        n = n - 1
    Loop
    StripTrailingDigits = Left$(s, n)
End Function

'Coclass part of an OCX class string ("MSWinsockLib.Winsock" -> "Winsock").
Private Function CoClassPart(ByVal libClass As String) As String
    Dim p As Long
    p = InStrRev(libClass, ".")
    If p > 0 Then CoClassPart = Mid$(libClass, p + 1) Else CoClassPart = libClass
End Function

Public Function getEventComplete(sFileName As String, strGuid As String, EventNum As Integer) As String
'*****************************
'Purpose: To return all the events from a filename by COM
'*****************************
    'On Error Resume Next
    Dim srT As SearchResults
    Dim srM As SearchResults
    Dim MI As MemberInfo
    Dim lSearchData As Long
    Dim m As Long, t As Long


    Dim tliTypeInfo As TypeInfo
    With tliTypeLibInfo
    
        .ContainingFile = sFileName

         Set srT = .GetTypes(, tliStEvents, False)
        For t = 1 To srT.count
        
            lSearchData = srT(t).SearchData

            Set tliTypeInfo = tliTypeLibInfo.GetTypeInfo(Replace(Replace(srT(t).name, "<", ""), ">", ""))

            'frmMain.txtCode.Text = frmMain.txtCode.Text & "'==================== " & srT(t).Name & "====================" & tliTypeLibInfo.GUID & vbCrLf & vbCrLf
            If tliTypeInfo.GUID = strGuid Then
              ' MsgBox "GuidFound " & srT(t).Name
               Set srM = tliTypeLibInfo.GetMembers(lSearchData)
            
            For m = 1 To srM.count
            
                DoEvents
                If m = EventNum Then
                Set MI = .GetMemberInfo(lSearchData, srM(m).InvokeKinds, srM(m).MemberId, srM(m).name)
                
                'frmMain.txtCode.Text = frmMain.txtCode.Text & getEventInfo(mi, srT(t).Name, False) & vbCrLf
                getEventComplete = getEventInfo(MI, srT(t).name, False)
                
                Exit Function
                End If
                Next m
           End If
            '
        Next t
End With

End Function

