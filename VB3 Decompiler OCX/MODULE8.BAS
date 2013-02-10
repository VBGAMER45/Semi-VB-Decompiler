' Module8
Option Explicit

Function fn08ED (p0062 As Integer) As String
Dim l0064 As String
  l0064 = Space$(p0062)
  Get gv088E, , l0064
  fn08ED = fn0357(l0064)
End Function

Sub sub040D (p0050 As Integer)
Dim l0052 As Integer
  gv0DD4 = 0
  For l0052 = gv0DC0 To gv0DC2
    If gv0D8C(l0052).M239C = p0050 Then
      sub0460 l0052
      Exit Sub
    End If
  Next l0052
End Sub

Sub sub0460 (p0054 As Integer)
Dim l0058 As Long
  If p0054 = gv0DD4 Then Exit Sub
  gv0DC6 = gv0D8C(p0054)
  If gv0DC6.M13E4 = 0 Then Exit Sub
  gv0DD4 = p0054
  gv0DDA = 1 + fn02C7(gv0DC6.M13E4, gv0918)
  l0058 = fn02C7(gv0DC6.M2395, gv0918)
  gv0DDE = gv0DDA + l0058
  Seek gv088E, gv0DDA
End Sub

