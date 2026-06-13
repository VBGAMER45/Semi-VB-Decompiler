Attribute VB_Name = "modPE"
'*********************************************
' modPE - lightweight PE identifier for VB Scanner
'
' Mirrors the runtime detection used by Semi VB Decompiler
' (modPeSkeleton.bas) but as a small, self contained reader:
'
'   .NET   -> PE optional header data directory 14 (COM/CLR
'             descriptor) has a non-zero RVA.
'   VB4    -> imports VB40032.DLL
'   VB5    -> imports MSVBVM50.DLL
'   VB6    -> imports MSVBVM60.DLL
'
' Everything else is reported as "not a target" and is skipped.
' Only 32-bit PE files carry the VB4/5/6 runtimes; 16-bit (NE)
' VB1/2/3 and VB4-16 are not PE and are therefore not detected.
'*********************************************
Option Explicit

Public Const RT_NONE As Long = 0
Public Const RT_VB4 As Long = 4
Public Const RT_VB5 As Long = 5
Public Const RT_VB6 As Long = 6
Public Const RT_NET As Long = 7

' Identify one file. Returns one of the RT_* constants.
' RT_NONE means the file is not a VB4/5/6 or .NET binary.
Public Function IdentifyFile(ByVal sPath As String) As Long
    Dim f As Integer
    Dim resultKind As Long
    resultKind = RT_NONE

    On Error GoTo done
    f = FreeFile
    Open sPath For Binary Access Read As #f

    If LOF(f) < 64 Then GoTo done

    ' --- DOS header: "MZ" ---
    If ReadWordAt(f, 0) <> &H5A4D Then GoTo done

    Dim peOff As Long
    peOff = ReadDWordAt(f, &H3C)
    If peOff <= 0 Or (peOff + 248) > LOF(f) Then
        ' Header runs past EOF; still let the data-dir reads be guarded.
        If peOff <= 0 Or (peOff + 24) > LOF(f) Then GoTo done
    End If

    ' --- PE signature: "PE\0\0" (low word = &H4550) ---
    If ReadDWordAt(f, peOff) <> &H4550 Then GoTo done

    Dim numSections As Long, sizeOptHdr As Long
    numSections = ReadWordAt(f, peOff + 6)
    sizeOptHdr = ReadWordAt(f, peOff + 20)

    Dim optOff As Long
    optOff = peOff + 24                 ' start of the optional header

    Dim magic As Long
    magic = ReadWordAt(f, optOff)

    Dim dirOff As Long
    If magic = &H20B Then
        dirOff = optOff + 112           ' PE32+ (64-bit)
    Else
        dirOff = optOff + 96            ' PE32
    End If

    ' --- Data directory 14 = COM/CLR descriptor => managed (.NET) ---
    Dim clrRva As Long
    clrRva = ReadDWordAt(f, dirOff + 14 * 8)
    If clrRva <> 0 Then
        resultKind = RT_NET
        GoTo done
    End If

    ' --- Import directory = data directory 1 ---
    Dim impRva As Long
    impRva = ReadDWordAt(f, dirOff + 1 * 8)
    If impRva = 0 Then GoTo done

    ' Section headers follow the optional header.
    Dim secOff As Long
    secOff = optOff + sizeOptHdr

    Dim impOff As Long
    impOff = RvaToOffset(f, impRva, secOff, numSections)
    If impOff = 0 Then GoTo done

    ' Walk every import descriptor (20 bytes each) and look at the
    ' DLL name; any of the three VB runtimes wins.
    Dim i As Long, descOff As Long
    Dim nameRva As Long, origThunk As Long, firstThunk As Long
    Dim nameOff As Long, dllName As String
    For i = 0 To 2000
        descOff = impOff + i * 20
        If (descOff + 20) > LOF(f) Then Exit For
        origThunk = ReadDWordAt(f, descOff)
        nameRva = ReadDWordAt(f, descOff + 12)
        firstThunk = ReadDWordAt(f, descOff + 16)
        ' Null descriptor terminates the table.
        If origThunk = 0 And nameRva = 0 And firstThunk = 0 Then Exit For
        If nameRva <> 0 Then
            nameOff = RvaToOffset(f, nameRva, secOff, numSections)
            If nameOff > 0 Then
                dllName = UCase$(ReadAsciiZ(f, nameOff, 32))
                Select Case dllName
                    Case "VB40032.DLL"
                        resultKind = RT_VB4: GoTo done
                    Case "MSVBVM50.DLL"
                        resultKind = RT_VB5: GoTo done
                    Case "MSVBVM60.DLL"
                        resultKind = RT_VB6: GoTo done
                End Select
            End If
        End If
    Next i

done:
    On Error Resume Next
    Close #f
    IdentifyFile = resultKind
End Function

' Friendly label for a RT_* constant.
Public Function RuntimeName(ByVal kind As Long) As String
    Select Case kind
        Case RT_VB4: RuntimeName = "VB4"
        Case RT_VB5: RuntimeName = "VB5"
        Case RT_VB6: RuntimeName = "VB6"
        Case RT_NET: RuntimeName = ".NET"
        Case Else:   RuntimeName = ""
    End Select
End Function

' Translate an RVA to a raw file offset using the section table.
Private Function RvaToOffset(ByVal f As Integer, ByVal rva As Long, _
                             ByVal secOff As Long, ByVal numSections As Long) As Long
    Dim i As Long, baseOff As Long
    Dim vSize As Long, vAddr As Long, rawSize As Long, rawPtr As Long, span As Long
    For i = 0 To numSections - 1
        baseOff = secOff + i * 40
        If (baseOff + 40) > LOF(f) Then Exit For
        vSize = ReadDWordAt(f, baseOff + 8)
        vAddr = ReadDWordAt(f, baseOff + 12)
        rawSize = ReadDWordAt(f, baseOff + 16)
        rawPtr = ReadDWordAt(f, baseOff + 20)
        span = vSize
        If rawSize > span Then span = rawSize
        If span <= 0 Then span = 1
        If rva >= vAddr And rva < (vAddr + span) Then
            RvaToOffset = (rva - vAddr) + rawPtr
            Exit Function
        End If
    Next i
    RvaToOffset = 0
End Function

' --- Little-endian readers (pos is a 0-based byte offset) ---

Private Function ReadDWordAt(ByVal f As Integer, ByVal pos As Long) As Long
    Dim v As Long
    Get #f, pos + 1, v
    ReadDWordAt = v
End Function

Private Function ReadWordAt(ByVal f As Integer, ByVal pos As Long) As Long
    Dim v As Integer
    Get #f, pos + 1, v
    ReadWordAt = v And &HFFFF&
End Function

' Read an ASCII, NUL-terminated string of at most maxLen characters.
Private Function ReadAsciiZ(ByVal f As Integer, ByVal pos As Long, ByVal maxLen As Long) As String
    Dim b As Byte, s As String, i As Long
    s = ""
    For i = 0 To maxLen - 1
        If (pos + 1 + i) > LOF(f) Then Exit For
        Get #f, pos + 1 + i, b
        If b = 0 Then Exit For
        s = s & Chr$(b)
    Next i
    ReadAsciiZ = s
End Function
