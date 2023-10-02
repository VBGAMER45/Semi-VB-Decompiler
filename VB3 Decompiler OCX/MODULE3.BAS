Attribute VB_Name = "Module3"
' Module3
Option Explicit

Function fn027F (p0082 As String) As Long
  gv0A98.M1EE3 = p0082
  LSet gv0A9E = gv0A98
  fn027F = gv0A9E.M1F14
End Function

Function fn0292 (p0090 As Integer) As String
  fn0292 = Right$(Hex$(p0090 Or &H100), 2)
 End Function

Function fn029A (p0094 As Integer) As String
  fn029A = " " & Right$(Hex$(p0094 Or &H100), 2)
End Function

Function fn02A3 (p0098 As Long) As String
Dim l009C As String * 9
  RSet l009C = Hex$(p0098)
  fn02A3 = l009C
End Function

Function fn02AC (p00A0 As Integer) As String
  fn02AC = " " & Right$(Hex$(p00A0 Or &H10000), 4)
End Function

Function fn02B5 (ByVal p00A4 As Long) As String
  fn02B5 = Right$(Hex$(p00A4 Or &H10000), 4)
End Function

Function fn02BD (p00A8 As Long) As Integer
  fn02BD = (p00A8 And &HFFFF0000) \ &H10000
End Function

Function fn02C7 (p00AC As Integer, p00AE As Integer) As Long
  fn02C7 = (CLng(p00AC) And &HFFFF&) * p00AE
End Function

Function fn02CF (ByVal p00B2 As Long) As Integer
  If p00B2 And &H8000& Then
    fn02CF = CInt(p00B2 Or &HFFFF0000)
  Else
    fn02CF = p00B2 And &HFFFF&
  End If
End Function

Function fn02D9 (p00B6 As Integer) As String
  gv0A5A.M1EF7 = p00B6
  LSet gv0A46 = gv0A5A
  fn02D9 = gv0A46.M1EE3
End Function

Function fn02E0 (p00BA As Long) As String
  gv0A9E.M1F14 = p00BA
  LSet gv0A98 = gv0A9E
  fn02E0 = gv0A98.M1EE3
End Function

Function fn02E7 (p00BE As String, p00C0 As Integer) As Integer
  gv0A46.M1EE3 = Mid$(p00BE, p00C0, 2)
  LSet gv0A5A = gv0A46
  fn02E7 = gv0A5A.M1EF7
End Function

Function fn02F0 (ByVal p00C4 As String) As String
  If Right$(p00C4, 1) = "\" Then fn02F0 = p00C4 Else fn02F0 = p00C4 & "\"
End Function

Function fn02FF (p00C8 As String, ByVal p00CA As String) As Integer
Dim l00CC As Integer
  Do
    fn02FF = l00CC
    l00CC = InStr(l00CC + 1, p00C8, p00CA)
  Loop While l00CC
End Function

Function fn0309 () As String
End Function

Function fn0325 (p00D8 As Integer, p00DA As Integer) As Integer
Dim l00DC As Long
  l00DC = fn0335(p00D8) + p00DA
  If l00DC >= &H8000& Then l00DC = l00DC Or &HFFFF0000
  fn0325 = l00DC
End Function

Function fn032D (p00E0 As Integer, p00E2 As Integer) As Integer
  fn032D = fn0335(p00E0) \ p00E2
End Function

Function fn0335 (p00E6 As Integer) As Long
  fn0335 = CLng(p00E6) And &HFFFF&
End Function

Function fn033D (p00EA As Integer, p00EC As Integer) As Integer
Dim l00EE As Long
  l00EE = fn0335(p00EA) * p00EC
  If l00EE >= &H8000& Then l00EE = l00EE Or &HFFFF0000
  fn033D = l00EE
End Function

Function fn034F (p00F6 As String, p00F8 As Integer) As String
Dim l00FA As Integer
  If p00F8 < 0 Then
    fn034F = "$" & fn02B5(p00F8) & "?"
  Else
    l00FA = InStr(p00F8, p00F6, Chr$(0)) - p00F8
    fn034F = Mid$(p00F6, p00F8, l00FA)
  End If
End Function

Function fn0357 (p00FE As String) As String
Dim l0100 As Integer
  l0100 = InStr(p00FE, Chr$(0))
  If l0100 Then
    fn0357 = Left$(p00FE, l0100 - 1)
  Else
    fn0357 = p00FE
  End If
End Function

Sub sub0272 (p0076 As ComboBox, p0078 As Variant)
Dim l007C As Long
Dim l007E As Long
  If VarType(p0078) = 8 Then
    For l007C = p0076.ListCount - 1 To 0 Step -1
      If p0076.List(l007C) = p0078 Then Exit For
    Next
  Else
    l007E = p0078
    For l007C = p0076.ListCount - 1 To 0 Step -1
      If p0076.ItemData(l007C) = l007E Then Exit For
    Next
  End If
  p0076.ListIndex = l007C
End Sub

Sub sub0286 (p0084 As Integer, p0086 As Integer)
Dim l008C As Long
  gv0B16.M1F51 = p0086
  gv0B16.M1F59 = p0084
  l008C = p0086 And &HFFFF&
  gv0B16.M1F8E = (l008C And &H1F) * 2
  gv0B16.M1F87 = (l008C \ &H20) And &H3F
  gv0B16.M1F80 = (l008C \ &H800) And &H1F
  l008C = p0084 And &HFFFF&
  gv0B16.M1F79 = l008C And &H1F
  gv0B16.M1F72 = (l008C \ &H20) And &HF
  gv0B16.M1F6A = ((l008C \ &H200) And &H7F) + 1980
  gv0B16.M1F4C = DateSerial(gv0B16.M1F6A, gv0B16.M1F72, gv0B16.M1F79) + TimeSerial(gv0B16.M1F80, gv0B16.M1F87, gv0B16.M1F8E)
  gv0B16.M034F = Format$(gv0B16.M1F4C, "hh:mm:ss")
  gv0B16.M1F62 = Format$(gv0B16.M1F4C, "dd.mm.yyyy")
End Sub

Sub sub0319 (p00D2 As Variant)
  gv0B16.M1F4C = p00D2
  gv0B16.M1F6A = Year(p00D2)
  gv0B16.M1F72 = Month(p00D2)
  gv0B16.M1F79 = Day(p00D2)
  gv0B16.M1F80 = Hour(p00D2)
  gv0B16.M1F87 = Minute(p00D2)
  gv0B16.M1F8E = Second(p00D2)
  gv0B16.M1F51 = (fn033D(gv0B16.M1F80, &H800) Or gv0B16.M1F87 * &H20 Or gv0B16.M1F8E \ 2)
  gv0B16.M1F59 = (gv0B16.M1F6A - 1980) * &H200 + gv0B16.M1F72 * &H20 + gv0B16.M1F79
  gv0B16.M034F = Format$(gv0B16.M1F4C, "hh:mm:ss")
  gv0B16.M1F62 = Format$(gv0B16.M1F4C, "dd.mm.yyyy")
End Sub

Sub sub0345 (p00F0 As Variant)
  gv0B16.M1F4C = p00F0
End Sub

