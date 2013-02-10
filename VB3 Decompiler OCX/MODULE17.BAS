Attribute VB_Name = "MODULE17"
' Module17
Option Explicit
'Declare Function extfn0E57 Lib "Kernel" Alias "GetModuleUsage" (ByVal p1%) As Integer
Const mc006E = 1 ' &H1%
Const mc0070 = 2 ' &H2%
Const mc0072 = 3 ' &H3%
Const mc0074 = 4 ' &H4%
Const mc0076 = 5 ' &H5%
Const mc0078 = 8 ' &H8%
Const mc007A = 9 ' &H9%
Const mc007C = 6 ' &H6%
Const mc007E = 7 ' &H7%
Const mc0080 = 10 ' &HA%
Const mc0082 = 11 ' &HB%
Const mc0084 = 12 ' &HC%
Const mc0086 = 13 ' &HD%
Const mc0088 = 14 ' &HE%
Const mc008A = 15 ' &HF%
Const mc008C = 16 ' &H10%
Const mc008E = 17 ' &H11%
Const mc0090 = 18 ' &H12%
Const mc0092 = 19 ' &H13%
Const mc0094 = 20 ' &H14%
Const mc0096 = 21 ' &H15%
Const mc00E0 = 4 ' &H4%
Const mc00EC = 16 ' &H10%
Private Type STARTUPINFO
      cb As Long
      lpReserved As String
      lpDesktop As String
      lpTitle As String
      dwX As Long
      dwY As Long
      dwXSize As Long
      dwYSize As Long
      dwXCountChars As Long
      dwYCountChars As Long
      dwFillAttribute As Long
      dwFlags As Long
      wShowWindow As Integer
      cbReserved2 As Integer
      lpReserved2 As Long
      hStdInput As Long
      hStdOutput As Long
      hStdError As Long
   End Type

   Private Type PROCESS_INFORMATION
      hProcess As Long
      hThread As Long
      dwProcessID As Long
      dwThreadID As Long
   End Type

   Private Declare Function WaitForSingleObject Lib "Kernel32" (ByVal _
      hHandle As Long, ByVal dwMilliseconds As Long) As Long

   Private Declare Function CreateProcessA Lib "Kernel32" (ByVal _
      lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Long

   Private Declare Function CloseHandle Lib "Kernel32" _
      (ByVal hObject As Long) As Long

   Private Declare Function GetExitCodeProcess Lib "Kernel32" _
      (ByVal hProcess As Long, lpExitCode As Long) As Long

   Private Const NORMAL_PRIORITY_CLASS = &H20&
   Private Const INFINITE = -1&

   Public Function ExecCmd(cmdline$) As Long
   
      Dim proc As PROCESS_INFORMATION
      Dim start As STARTUPINFO
Dim ret As Long

      ' Initialize the STARTUPINFO structure:
      start.cb = Len(start)

      ' Start the shelled application:
      ret& = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, _
         NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)

      ' Wait for the shelled application to finish:
         ret& = WaitForSingleObject(proc.hProcess, 0)
           While ret& = &H102 ' status timeout
            DoEvents
            ret& = WaitForSingleObject(proc.hProcess, 10)
          Wend
         
         
         Call GetExitCodeProcess(proc.hProcess, ret&)
         Call CloseHandle(proc.hThread)
         Call CloseHandle(proc.hProcess)
         ExecCmd = ret&
   End Function


Function fn0E69(ByVal pv009A As Integer) As Integer
  fn0E69 = (fn032D(pv009A, 8) And &H1FFF&) + ((pv009A And 6) * &H1000)
End Function

Function fn0E71(ByVal pv00A0 As Integer) As Long
Dim l00A2 As Integer
  l00A2 = (pv00A0 \ &H2000) And 3
  fn0E71 = 1 + fn012B(l00A2 + 3) + CLng(pv00A0 And &H1FFF) * 8
End Function

Function fn0E7C(ByVal pv00A8 As Integer) As Integer
  fn0E7C = fn02CF(fn0335(pv00A8 And &H1FFF) * 8 + ((pv00A8 \ &H1000) And 6))
End Function

Sub sub01F8()
    Beep
End Sub

Sub sub0201(p0058 As Integer, p005A As String)
  frmMain.control5.AddItem p005A
  p0058 = frmMain.control5.ListCount
  sub0EA5 p005A
End Sub

Sub sub0220(p0056 As Integer)
  Screen.MousePointer = 0
  While frmMain.control5.ListCount >= p0056
    If frmMain.control5.ListCount = 0 Then Exit Sub
    frmMain.control5.RemoveItem frmMain.control5.ListCount - 1
  Wend
End Sub

Sub sub047C(p004C As Integer, ByVal p004E As String)
Dim l0050 As Integer
  l0050 = frmMain.control5.ListCount
  If p004C > l0050 Then
    p004C = l0050
  Else
    While p004C < frmMain.control5.ListCount
      frmMain.control5.RemoveItem p004C
    Wend
  End If
  frmMain.control5.List(p004C - 1) = p004E
  sub0EA5 p004E
End Sub

Sub sub0E86(p00AE As Integer, p00B2 As Form)
Dim l00B4 As Integer
Dim l00B6 As String
Dim l00B8 As Integer
Dim l00BA  As Integer
' Const mc00E0 = 4 ' &H4%
' Const mc00EC = 16 ' &H10%
  Select Case p00AE
  Case gc7A1E
    gv7A1C = 0
  Case gc7A22
    gv7A1C = gv7A1C And Not (gc7A22 Or gc7A26 Or gc7A28)
  Case gc7A26
    gv7A1C = gv7A1C And Not (gc7A28 Or gc7A26)
  Case gc7A28
    If gv7A1C And p00AE Then
      gv7A1C = gv7A1C And Not gc7A26
    End If
  Case Else
    If p00AE And gv7A1C Then
      MsgBox gc039E
      Exit Sub
    End If
  End Select
  l00B8 = p00AE Or (p00AE - 1)
  l00B8 = l00B8 And Not gv7A1C
  If l00B8 And gc7A1E Then
    sub0745
    gv7B8C = ""
    p00B2.Caption = gc03AA
    Screen.MousePointer = 0
    'form3.Tag = "*.EXE"
    'form3.Show 1
    'l00B6 = form3.Tag
    l00B6 = m_Filename
    If Len(l00B6) = 0 Then
        MsgBox "No filename set!", vbCritical
        Exit Sub
    End If
    gv0882 = l00B6
    l00B4 = fn02FF(l00B6, "\")
    If l00B4 Then
      gv0886 = Left$(l00B6, l00B4)
      l00B6 = Mid$(l00B6, l00B4 + 1)
    Else
      sub01F8
    End If
    gv088A = l00B6
    l00B4 = InStr(l00B6, ".")
    If l00B4 Then l00B6 = Left$(l00B6, l00B4 - 1)
    gv7B90 = l00B6
    gv7A1C = gv7A1C Or gc7A1E
    p00B2.Caption = "VBDis " & gv7B90
    frmMain.control5.Clear
  End If
  If l00B8 And gc7A20 Then
    'frmMain.Show
    frmMain.control7.Value = 1
    sub0201 l00BA, gc03D2 & gv088A
    sub0097
    If iVBVersion Then
      sub0BBC l00BA
      If gv098A Then
        iVBVersion = 0
        gv7A1C = 0
      End If
      sub0427 l00BA, "VBRUN" & Format(iVBVersion) & "00.VBX"
    Else
      MsgBox gc034E
      gv7A1C = 0
    End If
    Screen.MousePointer = 0
    sub0220 l00BA
    If gv7A1C Then
      gv7A1C = gv7A1C Or gc7A20
      gv7B8C = m_OutputFolder
      p00B2.Caption = "VBDis " & gv088A
      'form2.control3 = gv7B90
      'form2.control4 = gv0886
      'form2.Show 1
    Else
      p00B2.Caption = gc03AA
      Exit Sub
    End If
    Screen.MousePointer = 11
  End If
  If l00B8 And gc7A22 Then
    sub0201 l00BA, gc03D6 & gv088A
    sub0D3A
    frmMain.control6.Value = 1
    sub0220 l00BA
    If gv7B30 And mc00E0 Then
      gv7A1C = 0
      Screen.MousePointer = 0
      Exit Sub
    Else
      gv7A1C = gv7A1C Or gc7A22
    End If
  End If
  If l00B8 And gc7A24 Then
    If (gv7A1C And gc7A24) = 0 Then
      frmMain.control5.Clear
      sub0201 l00BA, gc03D6 & gc0406
      sub0AC3
      gv7A1C = gv7A1C Or gc7A24
      sub0220 l00BA
    End If
  End If
  If l00B8 And gc7A26 Then
    sub0201 l00BA, gc03B6   'create project file
    sub0E93
    sub047C l00BA, gc03BE   'save form
    gv7A1C = gv7A1C Or gc7A26
    gv091A = True
    l00B6 = fn0FBE("", "mak")
    l00B4 = 0
    On Error Resume Next
      'l00B4 = Shell(gv0BBA & "frms2txt " & l00B6)
      l00B4 = ExecCmd(gv0BBA & "frms2txt " & l00B6)
      l00B4 = (l00B4 = 0)
      'If l00B4 Then
      '  If extfn0E57(l00B4) = 1 Then
      '    While extfn0E57(l00B4) = 1
      '      DoEvents
      '    Wend
      '  Else
      '    l00B4 = 0
      '  End If
      'End If
    On Error GoTo 0
    If l00B4 = 0 Or Err Then
      l00B8 = 0
      p00B2.Caption = l00B6
      MsgBox l00B6 & Chr$(10) & gc03BE, mc00EC, gc03C2
    Else
      l00B8 = gc7A28
    End If
  End If
  If l00B8 And gc7A28 Then
    frmMain.control5.Clear
    sub0201 l00BA, gc03C6
    gv7A1C = gv7A1C Or gc7A28
    sub04F2
    gv091A = True
    If gv7A1C And gc7A28 Then
      sub047C l00BA, gc03CA
    Else
      sub047C l00BA, gc03BE
    End If
  End If
  Screen.MousePointer = 0
End Sub

Private Sub sub0EA5(d1 As String)
  Screen.MousePointer = 11
  DoEvents
End Sub

Sub sub0EAF(p005C As Integer)
Static m0066 As Integer
' Const mc006E = 1 ' &H1%
' Const mc0070 = 2 ' &H2%
' Const mc0072 = 3 ' &H3%
' Const mc0074 = 4 ' &H4%
' Const mc0076 = 5 ' &H5%
' Const mc0078 = 8 ' &H8%
' Const mc007A = 9 ' &H9%
' Const mc007C = 6 ' &H6%
' Const mc007E = 7 ' &H7%
' Const mc0080 = 10 ' &HA%
' Const mc0082 = 11 ' &HB%
' Const mc0084 = 12 ' &HC%
' Const mc0086 = 13 ' &HD%
' Const mc0088 = 14 ' &HE%
' Const mc008A = 15 ' &HF%
' Const mc008C = 16 ' &H10%
' Const mc008E = 17 ' &H11%
' Const mc0090 = 18 ' &H12%
' Const mc0092 = 19 ' &H13%
' Const mc0094 = 20 ' &H14%
' Const mc0096 = 21 ' &H15%
  If Len(gv0BBA) Then Exit Sub
  sub058D p005C
  sub0BA0 p005C
  sub0B3E p005C
  gv0EE2(1) = -22 - 2
  gv0EE2(2) = -22 - 4
  gv0EE2(3) = gv0EE2(2)
  gv0EE2(4) = -22 - 8
  gv0EE2(5) = gv0EE2(4)
  gv0EE2(6) = -22 - &H10
  gv0EE2(7) = 1
  gv793C(m0066) = "g?"
  gv793C(mc006E) = "gn"
  gv793C(mc0070) = "gv"
  gv793C(mc0072) = "gc"
  gv793C(mc0074) = "T"
  gv793C(mc0076) = "M"
  gv793C(mc0078) = "extsub"
  gv793C(mc007A) = "extfn"
  gv793C(mc007C) = "sub"
  gv793C(mc007E) = "fn"
  gv793C(mc0080) = "event"
  gv793C(mc0082) = "v?"
  gv793C(mc0084) = "v"
  gv793C(mc0086) = "m"
  gv793C(mc0088) = "g"
  gv793C(mc008A) = "f"
  gv793C(mc008C) = "l"
  gv793C(mc008E) = "p"
  gv793C(mc0090) = "pv"
  gv793C(mc0092) = "mc"
  gv793C(mc0094) = "o?"
  gv793C(mc0096) = "f?"
End Sub

