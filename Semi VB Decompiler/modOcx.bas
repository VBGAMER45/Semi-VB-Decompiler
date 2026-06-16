Attribute VB_Name = "modOcx"
'*****************************
'modOcx.bas
'Purpose: Decode an external (ActiveX/OCX) control's IPersistStream property blob
'   into readable VB form properties, beating the commercial decompiler (which
'   punts the whole blob to an opaque OleObjectBlob="frx":offset reference).
'
'   How it works (validated against richtx32.ocx + MSWINSCK.OCX):
'     1. VB frames each external control's persisted data with the signature
'        0x12344321; the 4 bytes before it are the payload length.
'     2. We instantiate the control (CreateObject by its registered CLSID),
'        QueryInterface for IPersistStreamInit, and Load() the blob from an
'        in-memory IStream - exactly what the VB6 IDE / runtime does.
'     3. We enumerate the live control's writable, browsable properties (via the
'        TLI type library) and emit the ones whose value differs from a fresh
'        (InitNew) instance - i.e. the non-default properties VB itself persists.
'
'   Limitations (see DEVELOPMENT.md): a property the control saved at a value
'   equal to its fresh default can't be distinguished from "not saved" by read-
'   back, so it may be missed; and a few runtime-only properties can leak. The
'   control must be registered on the machine running the decompiler; when it
'   isn't (or Load fails), the caller falls back to OleObjectBlob.
'*****************************
Option Explicit

Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef pclsid As Any) As Long
Private Declare Function ProgIDFromCLSID Lib "ole32" (ByRef clsid As Any, ByRef lplpszProgID As Long) As Long
Private Declare Function SHCreateMemStream Lib "shlwapi" Alias "#12" (ByVal pInit As Long, ByVal cbInit As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Integer, ByRef prgpvarg As Long, ByRef pvargResult As Variant) As Long
Private Declare Sub CopyMemoryOcx Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Private Const CC_STDCALL As Long = 4
Private Const VT_I4 As Integer = 3

' IPersistStreamInit vtable slots (after IUnknown 0,1,2 + IPersist GetClassID=3):
'   4 IsDirty, 5 Load, 6 Save, 7 GetSizeMax, 8 InitNew
Private Const VT_QueryInterface As Long = 0
Private Const VT_Release As Long = 2
Private Const VT_PSI_Load As Long = 5
Private Const VT_PSI_InitNew As Long = 8

' INVOKEKIND
Private Const INV_FUNC As Long = 1
Private Const INV_GET As Long = 2
Private Const INV_PUT As Long = 4
Private Const INV_PUTREF As Long = 8
' FUNCFLAGS used to decide the persisted set
Private Const FUNCFLAG_FRESTRICTED As Long = &H1
Private Const FUNCFLAG_FBINDABLE As Long = &H4
Private Const FUNCFLAG_FDEFAULTBIND As Long = &H20
Private Const FUNCFLAG_FHIDDEN As Long = &H40
Private Const FUNCFLAG_FNONBROWSABLE As Long = &H400

'*****************************
'Purpose: Main entry. Reads the control's IPersistStream blob from [startPos,endPos),
'   decodes its properties, and AddText()s them at the current indentation. Returns
'   True on success (caller skips its fallback), False to fall back to OleObjectBlob.
'*****************************
Public Function EmitOcxProperties(ByVal F As Variant, ByVal className As String, _
                                  ByVal startPos As Long, ByVal endPos As Long) As Boolean
    On Error GoTo fail
    EmitOcxProperties = False

    ' --- 1. Extract the IPersistStream payload [magic, magic+payloadLen) ---
    Dim blob() As Byte
    If Not ReadOcxBlob(F, startPos, endPos, blob) Then Exit Function

    ' --- 2. Resolve the control's creatable CLSID + ProgID ---
    Dim clsid As String
    clsid = ClsidForClass(className)
    If Len(clsid) = 0 Then Exit Function
    Dim progid As String
    progid = ProgIdFromClsidStr(clsid)
    If Len(progid) = 0 Then Exit Function

    ' --- 3. Instantiate + Load the blob; instantiate a fresh defaults instance ---
    Dim ctl As Object, fresh As Object
    Set ctl = CreateObject(progid)
    If ctl Is Nothing Then Exit Function
    If Not LoadBlobIntoControl(ctl, blob) Then Exit Function

    Set fresh = CreateObject(progid)
    If Not fresh Is Nothing Then InitNewControl fresh

    ' --- 4. Enumerate writable browsable properties, emit the non-default ones ---
    Dim emitted As Long
    emitted = EmitChangedProperties(ctl, fresh, className)

    Set ctl = Nothing
    Set fresh = Nothing
    EmitOcxProperties = True
    Exit Function
fail:
    EmitOcxProperties = False
End Function

'*****************************
'Purpose: Locate the 0x12344321 signature in [startPos,endPos), read the 4-byte
'   payload length stored immediately before it, and copy [magic, magic+len) into
'   blob(). Restores the file pointer. Returns False if no signature.
'*****************************
Private Function ReadOcxBlob(ByVal F As Variant, ByVal startPos As Long, ByVal endPos As Long, ByRef blob() As Byte) As Boolean
    ReadOcxBlob = False
    Dim savePos As Long
    savePos = Loc(F)
    On Error GoTo done
    If endPos <= startPos + 8 Then GoTo done
    Dim n As Long
    n = endPos - startPos
    Dim scan() As Byte
    ReDim scan(n - 1)
    Seek F, startPos
    Get F, , scan
    Dim i As Long, payLen As Long
    For i = 4 To n - 5
        If scan(i) = &H21 And scan(i + 1) = &H43 And scan(i + 2) = &H34 And scan(i + 3) = &H12 Then
            payLen = scan(i - 4) Or (CLng(scan(i - 3)) * &H100&) Or (CLng(scan(i - 2)) * &H10000) Or (CLng(scan(i - 1)) * &H1000000)
            If payLen < 8 Or payLen > n - i Then payLen = n - i      'sanity: clamp to available
            ReDim blob(payLen - 1)
            CopyMemoryOcx blob(0), scan(i), payLen
            ReadOcxBlob = True
            GoTo done
        End If
    Next
done:
    On Error Resume Next
    Seek F, savePos
End Function

'*****************************
'Purpose: Recover an invisible external control's design-time Left/Top. VB stores
'   these for windowless OCX controls (Winsock, CommonDialog) in a trailer AFTER
'   the IPersistStream payload, as opcode records: 0x39 + Long (Left), 0x3A + Long
'   (Top), 0x00 padding, terminated by the control separator. Returns True if a
'   Left or Top was found (in-range). Restores the file pointer.
'*****************************
Public Function GetOcxTrailerLeftTop(ByVal F As Variant, ByVal startPos As Long, ByVal endPos As Long, _
                                     ByRef leftV As Long, ByRef topV As Long, ByRef hasLeft As Boolean, ByRef hasTop As Boolean) As Boolean
    GetOcxTrailerLeftTop = False
    hasLeft = False: hasTop = False
    Dim savePos As Long
    savePos = Loc(F)
    On Error GoTo done
    If endPos <= startPos + 8 Then GoTo done
    Dim n As Long
    n = endPos - startPos
    Dim scan() As Byte
    ReDim scan(n - 1)
    Seek F, startPos
    Get F, , scan

    ' Locate signature + payload length, then start parsing after the payload.
    Dim i As Long, magicAt As Long, payLen As Long
    magicAt = -1
    For i = 4 To n - 5
        If scan(i) = &H21 And scan(i + 1) = &H43 And scan(i + 2) = &H34 And scan(i + 3) = &H12 Then
            magicAt = i
            payLen = scan(i - 4) Or (CLng(scan(i - 3)) * &H100&) Or (CLng(scan(i - 2)) * &H10000) Or (CLng(scan(i - 1)) * &H1000000)
            Exit For
        End If
    Next
    If magicAt = -1 Then GoTo done
    Dim p As Long
    p = magicAt + payLen
    If payLen < 8 Or p < 0 Or p > n - 1 Then GoTo done

    ' Parse the trailer opcode records.
    Do While p <= n - 5
        Select Case scan(p)
            Case &H39   'Left
                leftV = DwordAt(scan, p + 1): hasLeft = True: p = p + 5
            Case &H3A   'Top
                topV = DwordAt(scan, p + 1): hasTop = True: p = p + 5
            Case &H0    'padding
                p = p + 1
            Case Else
                Exit Do  'separator / unknown -> stop
        End Select
    Loop

    ' Sanity: reject out-of-range coordinates.
    If hasLeft Then If leftV < -2000 Or leftV > 30000 Then hasLeft = False
    If hasTop Then If topV < -2000 Or topV > 30000 Then hasTop = False
    GetOcxTrailerLeftTop = hasLeft Or hasTop
done:
    On Error Resume Next
    Seek F, savePos
End Function

Private Function DwordAt(ByRef b() As Byte, ByVal idx As Long) As Long
    Dim v As Currency
    v = CCur(b(idx)) + CCur(b(idx + 1)) * 256@ + CCur(b(idx + 2)) * 65536@ + CCur(b(idx + 3)) * 16777216@
    If v > 2147483647@ Then v = v - 4294967296@
    DwordAt = CLng(v)
End Function

'*****************************
'Purpose: Find the registered coclass CLSID for a control's external class string
'   (e.g. "RichTextLib.RichTextBox") via gOcxList (same source the late-bound call
'   resolver uses). Returns "{...}" or "".
'*****************************
Private Function ClsidForClass(ByVal className As String) As String
    On Error Resume Next
    Dim i As Long
    If UBound(gOcxList) < 1 Then Exit Function
    For i = 0 To UBound(gOcxList) - 1
        If Len(gOcxList(i).strGuid) > 0 Then
            If StrComp(gOcxList(i).strLibname, className, vbTextCompare) = 0 Then
                ClsidForClass = gOcxList(i).strGuid
                Exit Function
            End If
        End If
    Next
End Function

'*****************************
'Purpose: CLSID string -> ProgID string via ProgIDFromCLSID.
'*****************************
Private Function ProgIdFromClsidStr(ByVal clsid As String) As String
    On Error Resume Next
    Dim guid(15) As Byte
    If CLSIDFromString(StrPtr(clsid), guid(0)) <> 0 Then Exit Function
    Dim pOle As Long
    If ProgIDFromCLSID(guid(0), pOle) <> 0 Then Exit Function
    If pOle = 0 Then Exit Function
    ProgIdFromClsidStr = LpwstrToStr(pOle)
    CoTaskMemFree pOle
End Function

'*****************************
'Purpose: Copy a system-allocated OLESTR (wide, null-terminated) to a VB String.
'*****************************
Private Function LpwstrToStr(ByVal p As Long) As String
    Dim ch As Integer, n As Long, pc As Long
    pc = p
    Do
        CopyMemoryOcx ch, ByVal pc, 2
        If ch = 0 Then Exit Do
        n = n + 1
        pc = pc + 2
    Loop While n < 1024
    If n = 0 Then Exit Function
    LpwstrToStr = Space$(n)
    CopyMemoryOcx ByVal StrPtr(LpwstrToStr), ByVal p, n * 2
End Function

'*****************************
'Purpose: QueryInterface ctl for IPersistStreamInit and Load() the blob from an
'   in-memory IStream. Returns True on S_OK.
'*****************************
Private Function LoadBlobIntoControl(ByVal ctl As Object, ByRef blob() As Byte) As Boolean
    On Error GoTo fail
    LoadBlobIntoControl = False
    Dim pPSI As Long
    pPSI = QueryPersistStreamInit(ctl)
    If pPSI = 0 Then Exit Function

    Dim pStm As Long
    pStm = SHCreateMemStream(VarPtr(blob(0)), UBound(blob) + 1)
    If pStm = 0 Then VtblRelease pPSI: Exit Function

    Dim hr As Long
    hr = VtblCall1(pPSI, VT_PSI_Load, pStm)

    VtblRelease pStm
    VtblRelease pPSI
    LoadBlobIntoControl = (hr = 0)
    Exit Function
fail:
    LoadBlobIntoControl = False
End Function

'*****************************
'Purpose: Call IPersistStreamInit::InitNew on a fresh control so its properties
'   read back as the control's true defaults.
'*****************************
Private Sub InitNewControl(ByVal ctl As Object)
    On Error Resume Next
    Dim pPSI As Long
    pPSI = QueryPersistStreamInit(ctl)
    If pPSI = 0 Then Exit Sub
    VtblCall0 pPSI, VT_PSI_InitNew
    VtblRelease pPSI
End Sub

'*****************************
'Purpose: QI a VB object for IPersistStreamInit {7FD52380-4E07-101B-AE2D-08002B2EC713}.
'   Returns the interface pointer (caller releases) or 0.
'*****************************
Private Function QueryPersistStreamInit(ByVal ctl As Object) As Long
    On Error GoTo fail
    Dim iid(15) As Byte
    If CLSIDFromString(StrPtr("{7FD52380-4E07-101B-AE2D-08002B2EC713}"), iid(0)) <> 0 Then Exit Function
    Dim pUnk As Long, pOut As Long, hr As Long
    pUnk = ObjPtr(ctl)
    Dim vts(1) As Integer, pv(1) As Long, a0 As Variant, a1 As Variant, res As Variant
    a0 = VarPtr(iid(0)): vts(0) = VT_I4: pv(0) = VarPtr(a0)
    a1 = VarPtr(pOut): vts(1) = VT_I4: pv(1) = VarPtr(a1)
    hr = DispCallFunc(pUnk, VT_QueryInterface * 4, CC_STDCALL, VT_I4, 2, vts(0), pv(0), res)
    If hr = 0 And res = 0 Then QueryPersistStreamInit = pOut
    Exit Function
fail:
    QueryPersistStreamInit = 0
End Function

'--- Raw vtable call helpers (DispCallFunc) ------------------------------------
Private Function VtblCall1(ByVal pIntf As Long, ByVal slot As Long, ByVal arg0 As Long) As Long
    Dim vts(0) As Integer, pv(0) As Long, a0 As Variant, res As Variant
    a0 = arg0: vts(0) = VT_I4: pv(0) = VarPtr(a0)
    If DispCallFunc(pIntf, slot * 4, CC_STDCALL, VT_I4, 1, vts(0), pv(0), res) = 0 Then VtblCall1 = res
End Function

Private Function VtblCall0(ByVal pIntf As Long, ByVal slot As Long) As Long
    Dim dvt As Integer, dpv As Long, res As Variant
    If DispCallFunc(pIntf, slot * 4, CC_STDCALL, VT_I4, 0, dvt, dpv, res) = 0 Then VtblCall0 = res
End Function

Private Sub VtblRelease(ByVal pIntf As Long)
    If pIntf <> 0 Then VtblCall0 pIntf, VT_Release
End Sub

'*****************************
'Purpose: Enumerate the live control's writable browsable properties (TLI) and
'   AddText() each whose value differs from the fresh defaults instance. Returns
'   the number emitted.
'*****************************
Private Function EmitChangedProperties(ByVal ctl As Object, ByVal fresh As Object, ByVal className As String) As Long
    On Error GoTo done
    Dim app As Object
    Set app = New TLIApplication
    Dim ii As Object
    Set ii = app.InterfaceInfoFromObject(ctl)
    If ii Is Nothing Then GoTo done

    ' Pass 1: collect which member IDs have a property-put (writable).
    Dim mi As Object
    Dim putIds As Collection
    Set putIds = New Collection
    For Each mi In ii.Members
        If (mi.InvokeKind = INV_PUT) Or (mi.InvokeKind = INV_PUTREF) Then
            On Error Resume Next
            putIds.Add True, "K" & mi.MemberId
            On Error GoTo done
        End If
    Next

    ' Pass 2: for each browsable property-get that is also writable, diff + emit.
    Dim emitted As Long
    Dim seen As Collection
    Set seen = New Collection
    For Each mi In ii.Members
        If mi.InvokeKind = INV_GET Then
            If HasKey(putIds, "K" & mi.MemberId) Then
                If Not IsHiddenMember(mi) Then
                    If Not HasKey(seen, mi.Name) Then
                        seen.Add True, mi.Name
                        If EmitOneProperty(ctl, fresh, mi.Name, className) Then emitted = emitted + 1
                    End If
                End If
            End If
        End If
    Next
    EmitChangedProperties = emitted
done:
End Function

Private Function IsHiddenMember(ByVal mi As Object) As Boolean
    'Exclude hidden/restricted members, and non-browsable members UNLESS they are
    'bindable/persisted. This keeps e.g. RichTextBox.TextRTF (non-browsable but
    'bindable - VB persists it) while dropping the Sel* runtime properties (pure
    'non-browsable, no persistence flags).
    On Error Resume Next
    Dim m As Long
    m = mi.AttributeMask
    Dim hidden As Boolean, nonBrowse As Boolean, bindable As Boolean
    hidden = ((m And FUNCFLAG_FHIDDEN) <> 0) Or ((m And FUNCFLAG_FRESTRICTED) <> 0)
    nonBrowse = ((m And FUNCFLAG_FNONBROWSABLE) <> 0)
    bindable = ((m And FUNCFLAG_FBINDABLE) <> 0) Or ((m And FUNCFLAG_FDEFAULTBIND) <> 0)
    IsHiddenMember = hidden Or (nonBrowse And Not bindable)
End Function

Private Function HasKey(ByVal c As Collection, ByVal k As String) As Boolean
    On Error GoTo no
    Dim v As Variant
    v = c.Item(k)
    HasKey = True
    Exit Function
no:
    HasKey = False
End Function

'*****************************
'Purpose: Read one property from the loaded + fresh controls; if changed, format
'   and AddText() it. Returns True if a line was emitted.
'*****************************
Private Function EmitOneProperty(ByVal ctl As Object, ByVal fresh As Object, ByVal propName As String, ByVal className As String) As Boolean
    On Error GoTo skip
    EmitOneProperty = False

    Dim vLoad As Variant, vDef As Variant
    On Error Resume Next
    Dim okLoad As Boolean, okDef As Boolean
    okLoad = ReadProp(ctl, propName, vLoad)
    okDef = ReadProp(fresh, propName, vDef)
    On Error GoTo skip
    If Not okLoad Then Exit Function

    ' Object-valued properties: Font -> BeginProperty block; others (Picture,
    ' MouseIcon) deferred to the frx path (Build 3).
    If IsObject(vLoad) Then
        EmitOneProperty = EmitFontProperty(propName, vLoad, vDef)
        Exit Function
    End If

    ' Skip when unchanged from the control's fresh default.
    If okDef Then
        If Not IsObject(vDef) Then
            If VarsEqual(vLoad, vDef) Then Exit Function
        End If
    End If

    Dim rhs As String, ok As Boolean
    rhs = FormatScalar(vLoad, ok)
    If Not ok Then Exit Function

    Call AddText(propName & " = " & rhs)
    EmitOneProperty = True
    Exit Function
skip:
    EmitOneProperty = False
End Function

Private Function ReadProp(ByVal obj As Object, ByVal propName As String, ByRef outVal As Variant) As Boolean
    On Error GoTo fail
    If IsObject(CallByName(obj, propName, VbGet)) Then
        Set outVal = CallByName(obj, propName, VbGet)
    Else
        outVal = CallByName(obj, propName, VbGet)
    End If
    ReadProp = True
    Exit Function
fail:
    ReadProp = False
End Function

Private Function VarsEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    'Compare via CStr - robust against OLE Automation types VB can't use in
    'arithmetic/relational expressions (VT_UI4 OLE_COLOR, VT_UI2, VT_I1...).
    On Error GoTo no
    VarsEqual = (CStr(a) = CStr(b))
    Exit Function
no:
    VarsEqual = False
End Function

'*****************************
'Purpose: Format a scalar property value as VB form-file RHS text. Sets ok=False
'   for values we can't render readably yet (long/binary strings -> frx in Build 3).
'   Uses CStr for whole-number types (locale-free for integers, and the only path
'   that works for VB-unsupported VTs like VT_UI4) and Str$ for floats (invariant
'   "." decimal).
'*****************************
Private Function FormatScalar(ByVal v As Variant, ByRef ok As Boolean) As String
    ok = True
    On Error GoTo fail
    Select Case VarType(v)
        Case vbBoolean
            FormatScalar = IIf(v, "-1  'True", "0  'False")
        Case vbString
            If IsBlobString(CStr(v)) Then
                ok = False                 'long/multiline/binary -> frx (Build 3)
            Else
                FormatScalar = """" & Replace(CStr(v), """", """""") & """"
            End If
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            FormatScalar = Trim$(Str$(v))
        Case vbInteger, vbLong, vbByte
            FormatScalar = FormatIntegerLike(v, ok)
        Case vbEmpty, vbNull
            ok = False
        Case Else
            ' Whole-number OLE types (VT_UI4=19, VT_UI2=18, VT_I1=16, VT_UI1=17,
            ' VT_INT=22, VT_UINT=23) and anything else numeric.
            FormatScalar = FormatIntegerLike(v, ok)
    End Select
    Exit Function
fail:
    ok = False
End Function

'*****************************
'Purpose: Render an integer-like value (incl. unsigned OLE types) as VB form text.
'   System colors / high-bit flags (unsigned >= 0x80000000) become &H........&.
'*****************************
Private Function FormatIntegerLike(ByVal v As Variant, ByRef ok As Boolean) As String
    On Error GoTo fail
    Dim d As Double
    d = CDbl(CStr(v))                       'CStr first: works on VB-unsupported VTs
    If d >= 2147483648# And d <= 4294967295# Then
        FormatIntegerLike = "&H" & Hex$(CLng(d - 4294967296#)) & "&"   'system color / flag
    Else
        FormatIntegerLike = CStr(v)
    End If
    ok = True
    Exit Function
fail:
    ok = False
End Function

Private Function IsBlobString(ByVal s As String) As Boolean
    If Len(s) > 255 Then IsBlobString = True: Exit Function
    Dim i As Long, c As Long
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 32 And c <> 9 Then IsBlobString = True: Exit Function   'control char (tab allowed)
    Next
End Function

'*****************************
'Purpose: Emit a Font property as VB's BeginProperty Font {GUID} ... EndProperty
'   block when the loaded font differs from the control's default font. Returns
'   True if a block was written. vLoad is the loaded font object; vDef may hold the
'   fresh control's default font (or not be a font - then we always emit).
'*****************************
Private Function EmitFontProperty(ByVal propName As String, ByVal vLoad As Variant, ByVal vDef As Variant) As Boolean
    On Error GoTo skip
    EmitFontProperty = False
    If Not IsObject(vLoad) Then Exit Function

    ' Only handle objects that look like a font (have Name + Size).
    Dim lName As String, lSize As String
    If Not FontSig(vLoad, lName, lSize) Then Exit Function

    ' Compare against the default font; emit only when something changed.
    Dim changed As Boolean
    changed = True
    If IsObject(vDef) Then
        Dim dName As String, dSize As String
        If FontSig(vDef, dName, dSize) Then
            changed = (FontSignature(vLoad) <> FontSignature(vDef))
        End If
    End If
    If Not changed Then Exit Function

    ' Standard StdFont/IFontDisp dispinterface GUID (fixed).
    Call AddText("BeginProperty " & propName & " {0BE35203-8F91-11CE-9DE3-00AA004BB851} ")
    gIdentSpaces = gIdentSpaces + 1
    Call AddText("Name = """ & lName & """")
    Call AddText("Size = " & lSize)
    Call AddText("Charset = " & FontNum(vLoad, "Charset", 0))
    Call AddText("Weight = " & FontWeight(vLoad))
    Call AddText("Underline = " & FontBool(vLoad, "Underline"))
    Call AddText("Italic = " & FontBool(vLoad, "Italic"))
    Call AddText("Strikethrough = " & FontBool(vLoad, "Strikethrough"))
    gIdentSpaces = gIdentSpaces - 1
    Call AddText("EndProperty")
    EmitFontProperty = True
    Exit Function
skip:
    EmitFontProperty = False
End Function

'*****************************
'Purpose: True if obj is a font-like object; returns its Name + formatted Size.
'*****************************
Private Function FontSig(ByVal obj As Variant, ByRef outName As String, ByRef outSize As String) As Boolean
    On Error GoTo no
    outName = CStr(CallByName(obj, "Name", VbGet))
    outSize = Trim$(Str$(CDbl(CallByName(obj, "Size", VbGet))))
    FontSig = True
    Exit Function
no:
    FontSig = False
End Function

'*****************************
'Purpose: A change-detection signature over a font's persisted fields.
'*****************************
Private Function FontSignature(ByVal obj As Variant) As String
    On Error Resume Next
    Dim nm As String, sz As String
    Call FontSig(obj, nm, sz)
    FontSignature = nm & "|" & sz & "|" & FontNum(obj, "Charset", 0) & "|" & FontWeight(obj) & "|" & _
                    FontBool(obj, "Underline") & "|" & FontBool(obj, "Italic") & "|" & FontBool(obj, "Strikethrough")
End Function

Private Function FontNum(ByVal obj As Variant, ByVal field As String, ByVal dflt As Long) As Long
    On Error GoTo d
    FontNum = CLng(CDbl(CStr(CallByName(obj, field, VbGet))))
    Exit Function
d:
    FontNum = dflt
End Function

'*****************************
'Purpose: Font weight - prefer .Weight; fall back to .Bold (700/400).
'*****************************
Private Function FontWeight(ByVal obj As Variant) As Long
    On Error GoTo tryBold
    FontWeight = CLng(CDbl(CStr(CallByName(obj, "Weight", VbGet))))
    If FontWeight > 0 Then Exit Function
tryBold:
    On Error GoTo d
    If CBool(CallByName(obj, "Bold", VbGet)) Then FontWeight = 700 Else FontWeight = 400
    Exit Function
d:
    FontWeight = 400
End Function

'*****************************
'Purpose: A font Boolean field as VB form text ("0  'False" / "-1  'True").
'*****************************
Private Function FontBool(ByVal obj As Variant, ByVal field As String) As String
    On Error GoTo f
    If CBool(CallByName(obj, field, VbGet)) Then FontBool = "-1  'True" Else FontBool = "0  'False"
    Exit Function
f:
    FontBool = "0  'False"
End Function
