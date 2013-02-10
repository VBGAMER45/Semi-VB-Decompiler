Attribute VB_Name = "Module1"
' Module1
Option Explicit
Const mc0084 = 1 ' &H1%
Const mc008C = 2 ' &H2%
Const mc008E = 3 ' &H3%
Const mc0090 = 4 ' &H4%
Const mc0092 = 5 ' &H5%
Const mc0094 = 6 ' &H6%
Const mc0096 = 7 ' &H7%
Const mc0098 = 8 ' &H8%
Const mc009A = 9 ' &H9%
Const mc009C = 10 ' &HA%
Const mc009E = 12 ' &HC%
Const mc00A0 = 14 ' &HE%
Const mc00A2 = 15 ' &HF%
Const mc00A4 = 16 ' &H10%

Function fn012B(p0130 As Integer) As Long
  fn012B = fn02C7(gv0784(p0130).M1CDD, gv0916)
End Function

Sub sub007C()
  gv0D04(mc0084) = "CURSOR"
  gv0D04(mc008C) = "BITMAP"
  gv0D04(mc008E) = "ICON"
  gv0D04(mc0090) = "MENU"
  gv0D04(mc0092) = "DIALOG"
  gv0D04(mc0094) = "STRING"
  gv0D04(mc0096) = "FONTDIR"
  gv0D04(mc0098) = "FONT"
  gv0D04(mc009A) = "ACCELERATOR"
  gv0D04(mc009C) = "RCDATA"
  gv0D04(mc009E) = "GROUP CURSOR"
  gv0D04(mc00A0) = "GROUP ICON"
  gv0D04(mc00A2) = "NAMETABLE"
  gv0D04(mc00A4) = "VERSIONINFO"
End Sub

Sub sub0088(p00A6 As Integer)
  gv06BC = p00A6
  If gv06BC Then
    gv06B4 = 1 + fn012B(p00A6)
    gv06B8 = fn0335(gv0784(p00A6).M1CE7)
    Seek gv088E, gv06B4
  Else
    gv06B4 = 0
    gv06B8 = 0
  End If
End Sub

Sub sub0097()
  gv08D6.M19CA = 0
  iVBVersion = 0
  If gv088E Then
    Close gv088E
  Else
    gv088E = FreeFile
  End If
  Open gv0882 For Binary Access Read As gv088E
  Get gv088E, , gv0894
  If gv0894.M19CA <> gc05D2 Or gv0894.M1A44 < &H40 Then Exit Sub
  Get gv088E, gv0894.M1A5F + 1, gv08D6
  If gv08D6.M19CA <> gc05D4 Then
    gv08D6.M19CA = 0
    Exit Sub
  End If
  gv0916 = 2 ^ gv08D6.M1B37
  sub00A3
  sub00B9
  sub00CA
  sub00DC
  sub00ED
  sub00FE
End Sub

Sub sub00A3()
Dim l00B8 As Long
Dim l00BA As Integer
Dim l00bc As Integer
Dim l00BE As Integer
Dim l00C2 As T1BE3
Dim l00C6 As T1BCE
  l00B8 = gv0894.M1A5F + gv08D6.M1A84 + 1
  gv067E = 0
  Get gv088E, l00B8, gv0A30: l00BA = Asc(gv0A30)
  While l00BA
    gv067E = gv067E + l00BA
    Get gv088E, , gv0A30: l00bc = Asc(gv0A30)
    Select Case l00bc
    Case 0
    Case &HFF
      While l00BA
        Get gv088E, , l00C2
        l00BA = l00BA - 1
      Wend
    Case Else
      While l00BA
        Get gv088E, , l00C6
        l00BA = l00BA - 1
      Wend
    End Select
    Get gv088E, , gv0A30: l00BA = Asc(gv0A30)
  Wend
  ReDim gv064C(gv067E)
  Get gv088E, l00B8, gv0A30: l00BA = Asc(gv0A30)
  l00BE = 0
  While l00BA
    Get gv088E, , gv0A30: l00bc = Asc(gv0A30)
    Select Case l00bc
    Case 0
      l00BE = l00BE + l00BA
    Case &HFF
      While l00BA
        Get gv088E, , l00C2
        gv064C(l00BE).M12D9 = Asc(l00C2.M12D9)
        gv064C(l00BE).M1C2B = Asc(l00C2.M1BF9)
        gv064C(l00BE).M13E4 = l00C2.M1BDB
        l00BA = l00BA - 1
        l00BE = l00BE + 1
      Wend
    Case Else
      While l00BA
        Get gv088E, , l00C6
        gv064C(l00BE).M12D9 = Asc(l00C6.M12D9)
        gv064C(l00BE).M1C2B = l00bc
        gv064C(l00BE).M13E4 = l00C6.M1BDB
        l00BA = l00BA - 1
        l00BE = l00BE + 1
      Wend
    End Select
    Get gv088E, , gv0A30: l00BA = Asc(gv0A30)
  Wend
End Sub

Sub sub00B9()
Dim l0110
  gv06BE = gv08D6.M1ADD
  ReDim gv0784(gv06BE)
  Seek gv088E, gv0894.M1A5F + gv08D6.M1AF8 + 1
  For l0110 = 1 To gv06BE
    Get gv088E, , gv0784(l0110)
  Next l0110
End Sub

Sub sub00CA()
Dim l00E6 As T227A
Dim l00EA As T238C
Dim l00EC As Long
Dim l00EE  As Integer
Dim l00F0 As Integer
Dim l00F2 As Integer
Dim l00F4  As Integer
Dim l00F6 As Integer
  If gv08D6.M1B01 = gv08D6.M1B0A Then Exit Sub
  l00EC = gv0894.M1A5F + gv08D6.M1B01 + 1
  Get gv088E, l00EC, gv0918
  gv0918 = 2 ^ gv0918
  ReDim gv0CD0(0): gv0D02 = 0
  ReDim gv0D8C(0): gv0DBE = 0
  Do
    Get gv088E, , l00E6
    If l00E6.M13DD = 0 Then Exit Do
    gv0D02 = gv0D02 + 1
    ReDim Preserve gv0CD0(gv0D02)
    gv0CD0(gv0D02).M13DD = l00E6.M13DD
    gv0CD0(gv0D02).M2283 = l00E6.M2283
    gv0CD0(gv0D02).M2295 = gv0DBE
    l00F0 = l00E6.M2283
    If l00E6.M13DD = (mc009C Or &H8000) Then
      gv0DC0 = gv0DBE + 1
      gv0DC2 = gv0DBE + l00F0
    End If
    ReDim Preserve gv0D8C(gv0DBE + l00F0)
    While l00F0 > 0
      gv0DBE = gv0DBE + 1
      Get gv088E, , l00EA
      gv0D8C(gv0DBE).M13DD = l00E6.M13DD
      gv0D8C(gv0DBE).M2395 = l00EA.M2395
      gv0D8C(gv0DBE).M12D9 = l00EA.M12D9
      gv0D8C(gv0DBE).M239C = l00EA.M239C
      gv0D8C(gv0DBE).M13E4 = l00EA.M13E4
      l00F0 = l00F0 - 1
    Wend
  Loop
  For l00EE = 1 To gv0D02
    l00F6 = gv0CD0(l00EE).M13DD
    If l00F6 < 0 Then
    Else
      Get gv088E, l00EC + l00F6, gv0A30
      gv0CD0(l00EE).M1D68 = Space$(Asc(gv0A30))
      Get gv088E, , gv0CD0(l00EE).M1D68
    End If
    l00F2 = gv0CD0(l00EE).M2295
    For l00F4 = l00F2 + 1 To l00F2 + gv0CD0(l00EE).M2283
      l00F6 = gv0D8C(l00F4).M239C
      If l00F6 < 0 Then
        gv0D8C(l00F4).M1D68 = Hex$(l00F6)
      Else
        Get gv088E, l00EC + l00F6, gv0A30
        gv0D8C(l00F4).M1D68 = Space$(Asc(gv0A30))
        Get gv088E, , gv0D8C(l00F4).M1D68
      End If
    Next l00F4
  Next l00EE
End Sub

Sub sub00DC()
Dim l00DC As Long
Dim l00DE As String
Dim l00E0
Dim l00E2
  gv0816 = 0
  l00DC = gv0894.M1A5F + gv08D6.M1B0A + 1
  Get gv088E, l00DC, gv0A30
  While Asc(gv0A30)
    gv0816 = gv0816 + 1
    l00DE = Space$(Asc(gv0A30))
    Get gv088E, , l00DE
    Get gv088E, , l00E0
    Get gv088E, , gv0A30
  Wend
  ReDim gv07E4(gv0816)
  Seek gv088E, l00DC
  For l00E2 = 1 To gv0816
    Get gv088E, , gv0A30
    l00DE = Space$(Asc(gv0A30))
    Get gv088E, , l00DE
    l00E0 = InStr(l00DE, Chr$(0))
    If l00E0 Then l00DE = Left$(l00DE, l00E0 - 1)
    gv07E4(l00E2).M1D68 = l00DE
    Get gv088E, , gv07E4(l00E2).M1D70
  Next l00E2
End Sub

Sub sub00ED()
Dim l00CC As Long
Dim l00CE As Long
Dim l00D0 As Integer
Dim l00D2 As Long
  l00CE = gv0894.M1A5F + gv08D6.M1B12 + 1
  l00CC = gv0894.M1A5F + gv08D6.M1B1A + 1
  ReDim gv084E(gv08D6.M1AE6)
  For l00D2 = 1 To gv08D6.M1AE6
    Get gv088E, l00CE, l00D0: l00CE = l00CE + 2
    Get gv088E, l00CC + l00D0, gv0A30
    gv084E(l00D2) = Space$(Asc(gv0A30))
    Get gv088E, , gv084E(l00D2)
    If Left$(gv084E(l00D2), 5) = "VBRUN" Then
      iVBVersion = Val(Mid$(gv084E(l00D2), 6, 1))
    ElseIf Left$(gv084E(l00D2), 5) = "VB400" Then
      iVBVersion = 4
    End If
  Next l00D2
End Sub

Sub sub00FE()
Dim l00D4 As Long
Dim l00D6 As String
Dim l00D8
Dim l00DA As Integer
  l00D4 = gv08D6.M1B22 + 1
  gv084C = 0
  Get gv088E, l00D4, gv0A30
  While Asc(gv0A30)
    gv084C = gv084C + 1
    l00D6 = Space$(Asc(gv0A30) + 2)
    Get gv088E, , l00D6
    Get gv088E, , gv0A30
  Wend
  ReDim gv081A(gv084C)
  Seek gv088E, l00D4
  For l00D8 = 1 To gv084C
    Get gv088E, , gv0A30
    l00D6 = Space$(Asc(gv0A30))
    Get gv088E, , l00D6
    l00DA = InStr(l00D6, Chr$(0))
    If l00DA Then l00D6 = Left$(l00D6, l00DA - 1)
    gv081A(l00D8).M1D68 = l00D6
    Get gv088E, , gv081A(l00D8).M1D70
  Next l00D8
End Sub

Sub sub0110(p00AA As Integer)
Dim l00AC
Dim l00B2 As T1C86
Dim l00B6 As T1CAD
  If p00AA = gv06BC Then Exit Sub
  sub0088 p00AA
  If gv0784(p00AA).M12D9 And gc07BE Then
    Get gv088E, gv06B4 + gv06B8, gv0758
    If gv0758 < 0 Then
      gv0758 = 0
    End If
    ReDim gv0726(gv0758)
    For l00AC = 1 To gv0758
      Get gv088E, , l00B2
      l00B6.M13E4 = l00B2.M13E4
      l00B6.M1CB7 = Asc(l00B2.M1C92)
      l00B6.M1C99 = Asc(l00B2.M1C99)
      l00B6.M1CA1 = l00B2.M1CA1
      l00B6.M1CA7 = l00B2.M1CA7
      gv0726(l00AC) = l00B6
    Next l00AC
  Else
    gv0758 = 0
  End If
End Sub

Sub sub011F(p0112 As Integer)
Dim l0114 As Integer
Dim l0118 As Integer
Dim l011A As Long
Dim l011E As T1CAD
Dim l0120 As Integer
Dim l0122 As Integer
Dim l0124 As String
Dim l0126() As Long
  gv06BC = 0: sub0110 p0112
  Erase gv0680: gv06B2 = 0
  ReDim gv0680(gv06BE, 255)
  ReDim l0126(255) As Long
  For l0114 = 1 To gv0758
    l011E = gv0726(l0114)
    If l011E.M1C99 <> 0 Or l011E.M1CB7 <> 2 Then
    Else
      l011A = fn0335(l011E.M13E4)
      gv06B2 = 1: l0124 = Chr$(1)
      Do
        gv06B2 = gv06B2 + 1
        Get gv088E, gv06B4 + l011A - 2, l0118: l011A = fn0335(l0118)
        l0126(gv06B2) = l011A
        For l0120 = 1 To Len(l0124)
          If l0126(Asc(Mid$(l0124, l0120, 1))) > l011A Then Exit For
        Next l0120
        l0124 = Left$(l0124, l0120 - 1) & Chr$(gv06B2) & Mid$(l0124, l0120)
      Get gv088E, , l0118
      l011A = fn0335(l0118)
      Loop Until l0118 = &HFFFF
      For l0120 = 1 To Len(l0124)
        l0122 = Asc(Mid$(l0124, l0120, 1))
        gv0680(l011E.M1CA1, l0120) = fn02CF(l0126(l0122))
      Next l0120
      gv0680(l011E.M1CA1, 0) = Len(l0124)
    End If
  Next l0114
End Sub

