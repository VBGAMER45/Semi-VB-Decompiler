' Module20
Option Explicit
Const mc004A = 2 ' &H2%
Const mc006A = 8192 ' &H2000%

Sub sub0E93 ()
  If gv7AA2 = 0 Then
    MsgBox gc039A
    Exit Sub
  End If
  sub0F50
  sub0F7C
  sub04E2 True
End Sub

Sub sub0F50 ()
Dim l0006 As String
Dim l0008  As Integer
  l0006 = fn0FBE(gv0F76, "MAK")
  sub0201 l0008, gc03DA & l0006
  sub089C l0006
  sub0F62
  sub087E
  sub0220 l0008
End Sub

Sub sub0F62 ()
Dim l0040 As Integer
Dim l0042   As Integer
Dim l0044 As String
' Const mc004A = 2 ' &H2%
' Const mc006A = 8192 ' &H2000%
  If gv0F88 Then
    l0040 = fn0E69(gv0F88)
    l0040 = gv0DE6(l0040 + mc004A)
    If l0040 = 0 Then sub01F8
    For l0042 = 1 To gv7AA2
      If gv7A70(l0042).M24AC = l0040 Then
        l0044 = fn0D20(l0042)
        If InStr(l0044, ".") = 0 Then l0044 = l0044 & ".BAS"
        sub05FF l0044
        sub0868
        Exit For
      End If
    Next
  ElseIf gv0F82 Then
    l0040 = fn0E69(gv0F86)
    sub05FF gv0F42(gv0F82): sub0868
  Else
  End If
  For l0042 = 1 To gv0F70
    sub05FF gv0F42(l0042)
    sub0868
  Next
  sub05FF "main.txt": sub0868
  For l0042 = 1 To gv7AA2
    If gv7A70(l0042).M24AC <> l0040 Then
      l0044 = fn0D20(l0042)
      If InStr(l0044, ".") = 0 Then l0044 = l0044 & ".BAS"
      sub05FF l0044
      sub0868
    End If
  Next
  If Len(gv0BB6) Then
    l0040 = FreeFile
On Error Resume Next
    Open gv0BB6 & "autoload.mak" For Input As l0040
    If Err = 0 Then
      While Not EOF(l0040)
        Line Input #l0040, l0044
        If UCase$(Left$(l0044, 7)) = "PROJWIN" Then
          sub05FF l0044: sub0868
        End If
      Wend
      Close l0040
    End If
On Error GoTo 0
  End If
  If gv0F84 Then
    sub022F mc006A
    sub05FF "IconForm="""
    sub05FF gv0F42(gv0F84)
    sub05FF """"
    sub0868
  End If
  sub05FF "Title="""
  sub05FF gv0F7A
  sub05FF """": sub0868
  l0044 = gv0882
  l0042 = fn02FF(l0044, "\")
  If l0042 Then l0044 = Mid$(l0044, l0042 + 1)
  sub05FF "ExeName="""
  sub05FF l0044
  sub05FF """": sub0868
  If Len(gv0F7E) Then
    sub05FF "HelpFile="""
    sub05FF gv0F7E
    sub05FF """": sub0868
  End If
End Sub

Sub sub0F6E (p0010 As Integer, p0012 As Integer)
' called (p1 As type, p2%)
Dim l0014 As String
Dim l0016   As Integer
Dim l0018 As Integer
Dim l001A   As Integer
  Screen.MousePointer = 11
  If p0010 <> gv7AB6 Then sub01A8 p0010
  l0014 = fn0FBE(fn0D20(p0010), "bas")
  sub089C l0014
  l0018 = gv0C52
  Print #gv0C52, "' "; fn0D20(p0010)
  sub0201 l001A, gc03BA
  sub0A21 p0010
  For l0016 = 1 To gv7AEE
    If gv7ABC(l0016).M24AC = gv7AB8 Then
      If gv7ABC(l0016).M13E4 <> -1 Then
        sub047C l001A, gc03DA & fn0D70(l0016)
        Print #gv0C52,
        sub01C8 l0016
      End If
    End If
  Next
  sub087E
  sub0220 l001A
End Sub

Sub sub0F7C ()
Dim l0036    As Integer
Dim l0038    As Integer
  sub0201 l0038, gc03B6
    sub047C l0038, gc03DA & "main.txt"
    sub089C fn0FBE("main", "txt")
    sub0DE9
    sub087E
  For l0036 = 1 To gv7AA2
    sub047C l0038, gc03DA & fn0D20(l0036)
    sub0F6E l0036, 0
  Next l0036
  sub0220 l0038
End Sub

