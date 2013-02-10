Attribute VB_Name = "Module18"
' Module18
Option Explicit
Const mc0060 = -32760 ' &H8008%
Const mc0062 = 16768 ' &H4180%
Const mc0064 = 8016 ' &H1F50%
Const mc0070 = 16 ' &H10%
Const mc0074 = 48 ' &H30%
Const mc0076 = 1 ' &H1%
Const mc0094 = 2 ' &H2%
Const mc00AC = 4 ' &H4%
Const mc00AE = 2 ' &H2%
Const mc00B0 = 3 ' &H3%
Const mc00B2 = 7 ' &H7%

Function fn0EE1() As Integer
  fn0EE1 = gv7B30
End Function

Function fn0EFB() As Integer
  fn0EFB = gv7B2E
End Function

Sub sub022F(p0068 As Variant)
' Const mc0070 = 16 ' &H10%
' Const mc0074 = 48 ' &H30%
' Const mc0076 = 1 ' &H1%
  If gv7B34 = 0 Then
    gv7B34 = gv0988: gv7B36 = gv0924
  End If
  If VarType(p0068) = 8 Then
    MsgBox p0068, mc0070, gc033A
    gv7B2E = -1
  Else
    If gv7B30 And p0068 Then Exit Sub
    sub0EF0 Val(p0068), mc0074, mc0076
  End If
End Sub

Sub sub0418()
  gv7B2E = 0: gv7B34 = 0
End Sub

Sub sub0EF0(p007A As Integer, p007C As Integer, p007E As Integer)
Dim l0080 As Integer
Dim l0082 As String
' Const mc0094 = 2 ' &H2%
' Const mc00AC = 4 ' &H4%
' Const mc00AE = 2 ' &H2%
' Const mc00B0 = 3 ' &H3%
' Const mc00B2 = 7 ' &H7%
  If gv7B2E = 0 Then gv7B2E = p007A
  If p007A And mc0060 Then l0080 = l0080 Or gc7B80
  If p007A And mc0062 Then l0080 = l0080 Or gc7B7E
  If p007A And mc0064 Then l0080 = l0080 Or gc7B7C
  If l0080 = 0 Then l0080 = gc7B7A
  gv7B30 = gv7B30 Or p007A
  Select Case p007A
  Case 0
  Case gc7B4E: l0082 = gc034A
  Case gc7B50: l0082 = gc034E
  Case gc7B52: l0082 = "VB Version" & Str$(iVBVersion) & gc0352
  Case gc7B54: l0082 = gc0356
  Case gc7B56: l0082 = gc035A
  Case gc7B58: l0082 = gc035E
  Case gc7B5A
    If Right$(gv7B4A, 4) = ".300" Then
      l0082 = Left$(gv7B4A, Len(gv7B4A) - 3) & "VBX"
    Else
      l0082 = gv7B4A
    End If
    l0082 = gc0366 & l0082
    p007C = mc0070
    p007E = mc0094
  Case gc7B5C: l0082 = gv7B4A & gc036A
  Case gc7B5E: l0082 = gc036E
  Case gc7B60: l0082 = gc0372
  Case gc7B62: l0082 = gc0376
  Case gc7B64: l0082 = gc037A
  Case gc7B66: l0082 = gc037E
  Case gc7B68: l0082 = gc0382
  Case gc7B6A
     l0082 = gc0386 & Chr$(10) & gv7B4A
  Case gc7B6C: l0082 = gc0362
  Case Else: l0082 = gc0342 & Format$(p007A)
  End Select
  If Len(l0082) Then
    If Right$(l0082, 1) = "!" Then l0082 = l0082 & Chr$(10) & gc033E
    gv7B48 = MsgBox(l0082, p007C Or p007E, gc033A)
    Select Case gv7B48
    Case mc00AC
      Exit Sub
    Case mc00AE, mc00B0, mc00B2
      GoTo L1E436
    End Select
  End If
  l0080 = l0080 And Not gv7B32
  gv7B32 = gv7B32 Or l0080
  Select Case l0080
  Case 0
  Case gc7B80
    MsgBox gc0326, mc0070, gc0322
    gv7B48 = mc00B0
  Case gc7B7E
    gv7B48 = MsgBox(gc032A, 48 + 1, gc0322)
  Case gc7B7C
    gv7B48 = MsgBox(gc032E, 32 + 1, gc0322)
  Case Else
    gv7B48 = MsgBox(gc0332, 64 + 1, gc0322)
  End Select
  Select Case gv7B48
  Case mc00AE, mc00B0, mc00B2
  Case Else
    Exit Sub
  End Select

L1E436:
  gv7B32 = gv7B32 Or gc7B82
  Unload frmMain
End Sub

