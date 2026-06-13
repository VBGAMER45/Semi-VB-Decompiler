Attribute VB_Name = "modScan"
'*********************************************
' modScan - helpers for VB Scanner
'   * INI persistence (last folder + decompiler path)
'   * folder picker  (SHBrowseForFolder)
'   * exe picker     (GetOpenFileName)
'   * small path utilities
' No OCX dependencies - pure Win32 API.
'*********************************************
Option Explicit

' ---- Folder browse ----
Private Type BROWSEINFO
    hOwner          As Long
    pidlRoot        As Long
    pszDisplayName  As String
    lpszTitle       As String
    ulFlags         As Long
    lpfn            As Long
    lParam          As Long
    iImage          As Long
End Type
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Const BIF_RETURNONLYFSDIRS As Long = 1

' ---- File open ----
Private Type OPENFILENAME
    lStructSize         As Long
    hwndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_HIDEREADONLY  As Long = &H4

' ---- INI ----
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

'-------------------------------------------------

Public Function IniPath() As String
    IniPath = AddSlash(App.Path) & "VBScanner.ini"
End Function

Public Function AddSlash(ByVal s As String) As String
    If Len(s) = 0 Then
        AddSlash = s
    ElseIf Right$(s, 1) <> "\" Then
        AddSlash = s & "\"
    Else
        AddSlash = s
    End If
End Function

' Parent folder of a path ("C:\a\b" -> "C:\a").
Public Function ParentFolder(ByVal s As String) As String
    Dim p As Long
    If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
    p = InStrRev(s, "\")
    If p > 0 Then ParentFolder = Left$(s, p - 1) Else ParentFolder = s
End Function

Public Function IniGet(ByVal section As String, ByVal key As String, ByVal def As String) As String
    Dim buf As String, n As Long
    buf = String$(1024, 0)
    n = GetPrivateProfileString(section, key, def, buf, 1024, IniPath())
    If n > 0 Then IniGet = Left$(buf, n) Else IniGet = def
End Function

Public Sub IniPut(ByVal section As String, ByVal key As String, ByVal val As String)
    WritePrivateProfileString section, key, val, IniPath()
End Sub

' Show the system folder picker; returns "" if cancelled.
Public Function BrowseFolder(ByVal hWnd As Long, ByVal sTitle As String) As String
    Dim bi As BROWSEINFO, pidl As Long, sPath As String, z As Long
    bi.hOwner = hWnd
    bi.lpszTitle = sTitle
    bi.ulFlags = BIF_RETURNONLYFSDIRS
    pidl = SHBrowseForFolder(bi)
    If pidl <> 0 Then
        sPath = String$(260, 0)
        If SHGetPathFromIDList(pidl, sPath) <> 0 Then
            z = InStr(sPath, vbNullChar)
            If z > 0 Then sPath = Left$(sPath, z - 1)
            BrowseFolder = sPath
        End If
        CoTaskMemFree pidl
    End If
End Function

' Show the open-file dialog to locate the decompiler exe.
Public Function BrowseForExe(ByVal hWnd As Long, ByVal sInitial As String) As String
    Dim ofn As OPENFILENAME, z As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = hWnd
    ofn.lpstrFilter = "Programs (*.exe)" & vbNullChar & "*.exe" & vbNullChar & _
                      "All files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    ofn.lpstrFile = "SemiVBDecompiler.exe" & String$(260, 0)
    ofn.nMaxFile = 260
    ofn.lpstrFileTitle = String$(260, 0)
    ofn.nMaxFileTitle = 260
    ofn.lpstrInitialDir = sInitial
    ofn.lpstrTitle = "Locate SemiVBDecompiler.exe"
    ofn.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY
    If GetOpenFileName(ofn) <> 0 Then
        z = InStr(ofn.lpstrFile, vbNullChar)
        If z > 0 Then
            BrowseForExe = Left$(ofn.lpstrFile, z - 1)
        Else
            BrowseForExe = ofn.lpstrFile
        End If
    End If
End Function
