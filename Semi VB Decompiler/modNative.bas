Attribute VB_Name = "modNative"
'*********************************************
'modNative
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************
'Module for Processing Native Code
Option Explicit
Private Type API_VBDEF
    rva As Long
    Ordinal As Long
    uName As String
    uDescr As String
End Type
Public exeVB6_APIDEF() As API_VBDEF

Private Type NativeProcType
    offset As Long
    sName As String
End Type
Global gNativeProcArray() As NativeProcType
Global NativeShowOffsets As Boolean
Global NativeShowHexInformation As Boolean

'Decompiled code, grouped by owning object (form/module/class), built once at
'load so the main tree can show it inline.  Keyed "OBJ_<UPPER object name>".
Public gNativeCodeCache As Collection

'*****************************
'ScanNativeProcsByPrologue
'Find procedures that have no event-link table entry - most importantly the
'Sub/Function procedures in .bas modules, but also any private class procedure.
'
'VB6 lays out every object's native code contiguously, in object-table order,
'and every procedure starts with the standard prologue:
'        55          push ebp
'        8B EC       mov  ebp, esp
'So we read the whole native code range (gProjectInfo.aStartOfCode ..
'aEndOfCode) and add any prologue address we did not already discover through
'the event tables.  Ownership is inferred from address range: procedures sort
'by object, so a new proc belongs to the object whose first known (event-table)
'proc is the greatest one at or below it; anything below the first such anchor
'belongs to the leading module.
'*****************************
Public Sub ScanNativeProcsByPrologue(ByVal F As Integer)
    On Error GoTo done

    Dim startVA As Long, endVA As Long, codeLen As Long
    startVA = gProjectInfo.aStartOfCode
    endVA = gProjectInfo.aEndOfCode
    If startVA = 0 Or endVA <= startVA Then Exit Sub
    codeLen = endVA - startVA
    If codeLen < 3 Or codeLen > 16000000 Then Exit Sub

    'Per-object anchor = lowest event-table proc address already found.
    Dim anchorMin() As Long
    ReDim anchorMin(UBound(gObjectNameArray))
    Dim p As Long, dotPos As Long, objName As String, oi As Long
    For p = 0 To UBound(gNativeProcArray) - 1
        If gNativeProcArray(p).offset <> 0 Then
            dotPos = InStr(gNativeProcArray(p).sName, ".")
            If dotPos > 0 Then
                objName = Left$(gNativeProcArray(p).sName, dotPos - 1)
                For oi = 0 To UBound(gObjectNameArray)
                    If gObjectNameArray(oi) = objName Then
                        If anchorMin(oi) = 0 Or gNativeProcArray(p).offset < anchorMin(oi) Then _
                            anchorMin(oi) = gNativeProcArray(p).offset
                        Exit For
                    End If
                Next oi
            End If
        End If
    Next p

    'Set of addresses we already know (event procs + SubMain) so we never dup.
    Dim seen As Collection
    Set seen = New Collection
    On Error Resume Next
    For p = 0 To UBound(gNativeProcArray) - 1
        If gNativeProcArray(p).offset <> 0 Then seen.Add 1, "k" & gNativeProcArray(p).offset
    Next p
    If gVBHeader.aSubMain <> 0 Then seen.Add 1, "k" & gVBHeader.aSubMain
    On Error GoTo done

    'Read the whole native code blob once (file offset = VA - ImageBase).
    Dim b() As Byte
    ReDim b(codeLen - 1)
    Seek F, startVA + 1 - OptHeader.ImageBase
    Get F, , b

    Dim j As Long, va As Long, owner As String
    For j = 0 To codeLen - 3
        If b(j) = &H55 And b(j + 1) = &H8B And b(j + 2) = &HEC Then
            va = startVA + j
            If Not KeyExists(seen, "k" & va) Then
                seen.Add 1, "k" & va
                owner = OwnerForAddress(va, anchorMin())
                gNativeProcArray(UBound(gNativeProcArray)).sName = owner & ".proc_" & Hex$(va)
                gNativeProcArray(UBound(gNativeProcArray)).offset = va
                ReDim Preserve gNativeProcArray(UBound(gNativeProcArray) + 1)
            End If
        End If
    Next j

done:
End Sub

'Pick the owning object name for a scanned procedure address.
Private Function OwnerForAddress(ByVal va As Long, ByRef anchorMin() As Long) As String
    Dim oi As Long, bestIdx As Long, bestVal As Long
    bestIdx = -1: bestVal = 0
    For oi = 0 To UBound(anchorMin)
        If anchorMin(oi) <> 0 And anchorMin(oi) <= va Then
            If anchorMin(oi) > bestVal Then bestVal = anchorMin(oi): bestIdx = oi
        End If
    Next oi
    If bestIdx <> -1 Then OwnerForAddress = gObjectNameArray(bestIdx): Exit Function

    'Below every anchor -> a leading module (object with no event table).
    For oi = 0 To UBound(anchorMin)
        If anchorMin(oi) = 0 Then OwnerForAddress = gObjectNameArray(oi): Exit Function
    Next oi
    OwnerForAddress = "Module1"
End Function

Private Function KeyExists(ByRef col As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = col(key)
    KeyExists = (Err.Number = 0)
    Err.Clear
End Function

Public Sub Decode(ByVal Filename As String)
'*****************************
'Purpose: To Get the procdures of a Native Exe and produce a report
'*****************************

Dim FileNum As Integer
    Dim F As Long
    F = FreeFile
    Open App.Path & "\dump\" & SFile & "\NativeOut.txt" For Output As #F
        Print #F, "Semi VB Decompiler - VisualBasicZone.com"
        Print #F, "Native Output for : " & Filename
        Print #F, "---------------------------------"
       
        Print #F, "Procedure Offsets:"
        If gProjectInfo.aNativeCode <> 0 Then
            If gVBHeader.aSubMain <> 0 Then
                 Print #F, gVBHeader.aSubMain
            End If
        End If
        Dim i As Integer
        For i = 0 To UBound(gNativeProcArray) - 1
             Print #F, gNativeProcArray(i).offset
        Next i
    Close #F

    'Decompile every procedure to readable VB, grouped by object, for the tree.
    Call BuildNativeCodeCache

End Sub

Public Sub BuildNativeCodeCache()
'*****************************
'Decompile each procedure (grouped by its owning object) and cache the result
'so the main tree can display per-object code without re-running the engine.
'Addresses come from gNativeProcArray - the same authoritative list the Native
'Procedure Decompile window uses - so they are always valid (SubNamelist's
'event-proc addresses are unreliable).  The procedure name in each header is
'resolved internally by the engine from the address.
'*****************************
    On Error Resume Next
    Dim objNames() As String, objCode() As String, objCount As Long
    Dim p As Long, oi As Long, found As Long, addr As Long, body As String, total As Long, done As Long
    Dim sn As String, objName As String, dotPos As Long

    Set gNativeCodeCache = New Collection
    Dim ub As Long
    ub = -1
    ub = UBound(gNativeProcArray)                      '-1 stays if not dimensioned
    If ub < 0 Then Exit Sub
    total = ub
    ReDim objNames(64): ReDim objCode(64): objCount = 0

    For p = 0 To ub - 1
        If CancelDecompile = True Then Exit For
        addr = gNativeProcArray(p).offset
        If addr = 0 Then GoTo nextProc

        'Owning object is the prefix of the synthetic name "Object.proc:addr"
        sn = gNativeProcArray(p).sName
        dotPos = InStr(sn, ".")
        If dotPos > 0 Then objName = UCase$(Left$(sn, dotPos - 1)) Else objName = "MODULE1"

        found = -1
        For oi = 0 To objCount - 1
            If objNames(oi) = objName Then found = oi: Exit For
        Next
        If found = -1 Then
            If objCount > UBound(objNames) Then
                ReDim Preserve objNames(objCount + 64): ReDim Preserve objCode(objCount + 64)
            End If
            objNames(objCount) = objName
            objCode(objCount) = ""
            found = objCount: objCount = objCount + 1
        End If

        done = done + 1
        frmMain.txtStatus.Text = frmMain.txtStatus.Text & "Decompiling " & sn & " (" & done & "\" & total & ")" & vbCrLf
        frmMain.txtStatus.Refresh
        DoEvents

        body = modNativeToVB.DecompileNativeProcToVB(addr)
        objCode(found) = objCode(found) & body & vbCrLf
nextProc:
    Next p

    For oi = 0 To objCount - 1
        gNativeCodeCache.Add objCode(oi), "OBJ_" & objNames(oi)
    Next
End Sub

Public Function GetNativeObjectCode(ByVal objName As String) As String
'*****************************
'Return the cached decompiled code for an object, or empty stubs built from the
'procedure list when nothing was cached (e.g. P-Code projects).
'*****************************
    Dim code As String, p As Long
    On Error Resume Next
    If Not gNativeCodeCache Is Nothing Then
        code = gNativeCodeCache("OBJ_" & UCase$(objName))
        If Len(code) > 0 Then GetNativeObjectCode = code: Exit Function
    End If
    'Fallback: signature-only stubs
    For p = 0 To UBound(gProcedureList)
        If UCase$(gProcedureList(p).strParent) = UCase$(objName) And gProcedureList(p).strProcedureName <> "" Then
            If Right$(gProcedureList(p).strProcedureName, 1) = ")" Then
                code = code & "Private Sub " & gProcedureList(p).strProcedureName & vbCrLf
            Else
                code = code & "Private Sub " & gProcedureList(p).strProcedureName & "()" & vbCrLf
            End If
            code = code & "End Sub" & vbCrLf
        End If
    Next
    GetNativeObjectCode = code
End Function

Sub VBFunction_Description_Init(ByVal fRes As String)
'*****************************
'Purpose: To load the Msvbvm60.dll api list from a file
'*****************************
Dim lfp As Integer, i As Long
Dim sAdr As String, sOrd As String, sName As String, sDef As String
lfp = FreeFile
Erase exeVB6_APIDEF()

    Open fRes For Input Access Read As #lfp
        i = 0
        Do
        i = i + 1
            Input #lfp, sAdr, sOrd, sName, sDef
            If LCase$(sAdr) <> "eof" Then
                ReDim Preserve exeVB6_APIDEF(1 To i)
                exeVB6_APIDEF(i).rva = Val("&H" & sAdr)

                exeVB6_APIDEF(i).Ordinal = CLng(sOrd)
                exeVB6_APIDEF(i).uName = sName
                exeVB6_APIDEF(i).uDescr = sDef
            Else
                Exit Do
            End If
        Loop Until EOF(1)
    
    Close #lfp

End Sub
Public Function VBFunction_Description(ByVal inOrdinal As Long, ByVal inAPIname As String, ByRef outRName As String) As String
'*****************************
'Purpose: To return the description of a function
'*****************************
Dim i As Long


If inOrdinal > 0 And inAPIname = vbNullString Then
    'by ordinal :
    For i = 1 To UBound(exeVB6_APIDEF)
        If exeVB6_APIDEF(i).Ordinal = inOrdinal Then
            VBFunction_Description = exeVB6_APIDEF(i).uDescr
            outRName = exeVB6_APIDEF(i).uName
            Exit Function
        End If
    Next i

Else
    'by name:
   
    For i = 1 To UBound(exeVB6_APIDEF)
        If exeVB6_APIDEF(i).uName = inAPIname Then
            VBFunction_Description = exeVB6_APIDEF(i).uDescr
            
            Exit Function
        End If
    Next i
End If

VBFunction_Description = "Error API incorrect or not present in msvbvm60.dll"

End Function


