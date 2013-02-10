Attribute VB_Name = "MODULE5"
' Module5
Option Explicit
Declare Function extfn0529 Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$, ByVal p5%, ByVal p6$) As Integer
Declare Function extfn0544 Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn055A Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal p1$, ByVal p2%) As Integer
'Declare Sub extsub0571 Lib "User" Alias "WinHelp" ()
Const mc0044 = 48 ' &H30%
Const mc0050 = 64 ' &H40%

Function fn057C(ByVal p0024 As String) As String
Dim l0026 As Integer
Dim l0028 As String
  l0028 = Space$(128)
  l0026 = extfn0529("Visual Basic" & Chr$(0), p0024 & Chr$(0), "", l0028, 128, "vb.ini")
  fn057C = Left$(l0028, l0026)
End Function

Sub sub058D(p002A As Integer)
Dim l002C As Integer
Dim l0030 As String
Dim l0032 As Integer
' Const mc0044 = 48 ' &H30%
' Const mc0050 = 64 ' &H40%
On Error Resume Next
  'gv0BBA = fn02F0((App.Path))
  gv0BBA = fn02F0(m_SetDataPath)
  gv7B4A = gc1208
  If p002A Then sub047C p002A, gc03D2 & gv7B4A
  While Len(Dir$(gv0BBA & gv7B4A, 0)) = 0
    l0032 = Screen.MousePointer
    Screen.MousePointer = 0
      MsgBox gv0BBA & gv7B4A, mc0044, gc0396
      'form3.Tag = gv7B4A
      'form3.Show 1
    Screen.MousePointer = l0032
    'l0030 = form3.Tag
    l0030 = m_Filename
    If Len(l0030) = 0 Then
        MsgBox "No filename set!!", vbCritical
        Exit Sub
    End If
    
    If Right$(l0030, 3) = "300" Then
      l002C = fn02FF(l0030, "\")
      If l002C Then gv0BBA = Left$(l0030, l002C)
    End If
  Wend
  l0030 = Space$(256)
  l002C = extfn055A(l0030, Len(l0030))
  gv0BAE = fn02F0(Left$(l0030, l002C))
  l002C = extfn0544(l0030, Len(l0030))
  gv0BB2 = fn02F0(Left$(l0030, l002C))
  gv0BB6 = fn057C("vbpath")
  If Len(gv0BB6) = 0 Then
    gv0BB6 = gv0BBA
    If Len(Dir$(gv0BB6 & "vb.exe")) = 0 Then
      gv0BB6 = ""
    End If
  End If
  If Len(gv0BB6) Then
    If Right$(gv0BB6, 1) <> "\" Then gv0BB6 = gv0BB6 & "\"
  Else
    MsgBox gc0392, mc0050
  End If
  Err = 0
End Sub

