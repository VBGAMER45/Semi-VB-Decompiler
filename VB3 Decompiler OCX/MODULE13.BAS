Attribute VB_Name = "Module13"
' Module13
Option Explicit
Const mc00C2 = 4 ' &H4%
Const mc00CA = 16384 ' &H4000%
Const mc00DC = 1 ' &H1%
Const mc00E4 = 1 ' &H1%
Const mc00E6 = 2 ' &H2%
Const mc00E8 = 2 ' &H2%
Const mc00EA = 3 ' &H3%
Const mc00EC = 3 ' &H3%
Const mc00EE = 4 ' &H4%
Const mc00F0 = 4 ' &H4%
Const mc00F2 = 5 ' &H5%
Const mc00F4 = 5 ' &H5%
Const mc00F6 = 6 ' &H6%
Const mc00F8 = 6 ' &H6%
Const mc00FA = 7 ' &H7%
Const mc00FC = 7 ' &H7%
Const mc00FE = 12 ' &HC%
Const mc0100 = 8 ' &H8%
Const mc0102 = 10 ' &HA%
Const mc0104 = 9 ' &H9%
Const mc0106 = 8 ' &H8%
Const mc0108 = 10 ' &HA%
Const mc010A = 9 ' &H9%

Function fn0BB2(p00B0 As Integer) As String
Dim l00B2 As Integer
Dim l00B4 As Integer
  l00B2 = gv1148 + (p00B0 And &HFF) * &H18
  l00B4 = Asc(Mid$(gv1142, l00B2, 1))
  fn0BB2 = Mid$(gv1142, l00B2 + 1, l00B4)
End Function

Function fn0BCF(p00D0 As Integer) As String
Dim l00D2 As Integer
Dim l00D4 As Integer
  l00D2 = Abs(gv1646.M2D10(p00D0).M2CBB)
  If l00D2 = 0 Then Exit Function
  l00D4 = InStr(l00D2, gv15DA, "§")
  fn0BCF = Mid$(gv15DA, l00D2, l00D4 - l00D2)
End Function

Function fn0BDF() As String
Dim l00A4 As Integer
Dim l00A6 As Integer
  l00A4 = gv15E0.M2CBB
  If l00A4 <= 0 Then Exit Function
  l00A6 = InStr(l00A4, gv15DA, "§")
  fn0BDF = Mid$(gv15DA, l00A4, l00A6 - l00A4)
End Function

Sub sub0BA0(p00D6 As Integer)
Dim l00D8 As Integer
Dim l00DA As Integer
' Const mc00DC = 1 ' &H1%
' Const mc00E4 = 1 ' &H1%
' Const mc00E6 = 2 ' &H2%
' Const mc00E8 = 2 ' &H2%
' Const mc00EA = 3 ' &H3%
' Const mc00EC = 3 ' &H3%
' Const mc00EE = 4 ' &H4%
' Const mc00F0 = 4 ' &H4%
' Const mc00F2 = 5 ' &H5%
' Const mc00F4 = 5 ' &H5%
' Const mc00F6 = 6 ' &H6%
' Const mc00F8 = 6 ' &H6%
' Const mc00FA = 7 ' &H7%
' Const mc00FC = 7 ' &H7%
' Const mc00FE = 12 ' &HC%
' Const mc0100 = 8 ' &H8%
' Const mc0102 = 10 ' &HA%
' Const mc0104 = 9 ' &H9%
' Const mc0106 = 8 ' &H8%
' Const mc0108 = 10 ' &HA%
' Const mc010A = 9 ' &H9%
  l00D8 = FreeFile
  gv7B4A = gc156A
  If p00D6 Then sub047C p00D6, gc03D2 & gv7B4A
  Open gv0BBA & gv7B4A For Binary As l00D8
  Get l00D8, , l00DA
  If l00DA = 0 Then sub022F mc00CA
  gv15DA = String$(l00DA, 0)
  Get l00D8, , gv1646
  Get l00D8, , gv15DA
  Close l00D8
  gv11F2(mc00DC) = 0
  gv11F2(mc00E6) = mc00E4
  gv11F2(mc00EA) = mc00E8
  gv11F2(mc00EE) = mc00EC
  gv11F2(mc00F2) = mc00F0
  gv11F2(mc00F6) = mc00F4
  gv11F2(mc00FA) = mc00F8
  gv11F2(mc00FE) = mc00FC
  gv11F2(mc0102) = mc0100
  gv11F2(mc0106) = mc0104
  gv11F2(mc010A) = mc0108
  gv11DC(0) = mc00DC
  gv11DC(mc00E4) = mc00E6
  gv11DC(mc00E8) = mc00EA
  gv11DC(mc00EC) = mc00EE
  gv11DC(mc00F0) = mc00F2
  gv11DC(mc00F4) = mc00F6
  gv11DC(mc00F8) = mc00FA
  gv11DC(mc00FC) = mc00FE
  gv11DC(mc0100) = mc0102
  gv11DC(mc0104) = mc0106
  gv11DC(mc0108) = mc010A
  gv157E(1) = "Int"
  gv157E(2) = "Lng"
  gv157E(3) = "Sng"
  gv157E(4) = "Dbl"
  gv157E(5) = "Cur"
  gv157E(6) = "Var"
  gv157E(7) = "Str"
  gv1594(1) = "Integer"
  gv1594(2) = "Long"
  gv1594(3) = "Single"
  gv1594(4) = "Double"
  gv1594(5) = "Currency"
  gv1594(6) = "Variant"
  gv1594(7) = "String"
  gv1594(8) = "String *"
  sub0AD1
End Sub

Sub sub0BBC(p00BA As Integer)
Dim l00bc As Integer
Dim l00BE As Integer
' Const mc00C2 = 4 ' &H4%
' Const mc00CA = 16384 ' &H4000%
  Select Case iVBVersion
  Case 2: gv1148 = &H7500
  Case 3: gv1148 = &H5600
  Case Else
    sub022F mc00C2
    Exit Sub
  End Select
  l00bc = FreeFile
  gv7B4A = "vbdis" & Format$(iVBVersion) & gc156E
  If p00BA Then sub047C p00BA, gc03D2 & gv7B4A
  Open gv0BBA & gv7B4A For Binary As l00bc
  Get l00bc, , l00BE
  l00BE = l00BE Xor (iVBVersion * &H100)
  If l00BE = 2 Then
    Get l00bc, , gv23AA
    gv7856 = String$(gc1574 * 3, 0): Get l00bc, , gv7856
    Get l00bc, , l00BE
    gv1142 = String$(l00BE, 0): Get l00bc, , gv1142
  Else
    sub022F mc00CA
    gv098A = True
  End If
  Close l00bc

End Sub

Sub sub0BEB()
End Sub

