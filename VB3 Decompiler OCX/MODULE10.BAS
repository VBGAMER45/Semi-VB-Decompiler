' Module10
Option Explicit
Const mc0050 = 7 ' &H7%
Const mc0052 = 6 ' &H6%
Const mc005E = 256 ' &H100%
Const mc008A = 1 ' &H1%
Const mc00B0 = 90 ' &H5A%
Const mc00BE = 86 ' &H56%

Private Sub sub0944 ()
Dim l00C0 As Integer
Dim l00C2 As Integer
Dim l00C4 As String
Dim l00C8 As T15B8
  l00C0 = gv0DE2
  Do
    l00C2 = gv0DE6(l00C0 + gc0E20)
    gv7A70(l00C2).M24AC = l00C0
    Get gv088E, fn0E71(l00C0), l00C8
    gv7A70(l00C2).M1374 = l00C8.M1374
    l00C4 = "Module" & Format$(l00C2)
    sub0441 l00C2, l00C4
    l00C0 = gv0DE6(l00C0 + gc0E22)
  Loop Until l00C0 <= 0
End Sub

Private Sub sub0951 (p006A As Integer)
Dim l006C As Integer
Dim l006E As Integer
Dim l0070 As Integer
Dim l0074 As T1CAD
Dim l0076 As Integer
Dim l0078  As Integer
Dim l007A As Long
Dim l007C As Integer
' Const mc008A = 1 ' &H1%
  sub0110 p006A
  l007C = (p006A - 3) * &H2000
  For l006C = 1 To gv0758
    l0074 = gv0726(l006C)
    If l0074.M1C99 <> 0 Or l0074.M1CB7 <> 2 Then
      sub01F8
    Else
      l007A = fn0335(l0074.M13E4)
      l006E = l007C + (l007A \ 8) - 4
      If gv0DE6(l006E) <> gc0E1A Then sub01F8
      l0076 = gv0DE6(l006E + gc0E26)
      l0070 = l0074.M1CA1
      If l0076 = 0 Or gv0DE6(l0076) <> gc0E18 Then
        sub01F8
      Else
        Do
          If gv0DE6(l006E) <> gc0E1A Then sub022F mc008A
          gv7ABC(gv0DE6(l006E + gc0E24)).M1C2B = l0070
          Get gv088E, gv06B4 + l007A, l0078
          l007A = fn0335(l0078)
          l006E = l007C + (l007A \ 8) - 4
        Loop Until l0078 = &HFFFF
      End If
    End If
  Next l006C
End Sub

Sub sub095F ()
Dim l008E As Integer
Dim l0090 As Long
  gv06BC = 0
  l0090 = 0
  sub0088 3
  If gv06BE <= 3 Then
    Exit Sub
  Else
    For l008E = 4 To 6
      If (gv0784(l008E).M12D9 And mc005E) = 0 Then Exit For
      sub0088 l008E
    Next
    l0090 = &H10000 * (l008E - 4) + gv06B8
  End If
  sub0450
  ReDim gv0DE6(l0090 \ 8)
  ReDim gv7876(0)
  gv7AA2 = 0
  gv7AEE = 0
  gv06BC = 0
  sub096A
  ReDim gv7A70(1 To gv7AA2)
  sub0944
  sub0979
  ReDim gv7ABC(gv7AEE)
  sub01E7
  If gv091A Then
    sub0988
  End If
  sub0435
  sub0995
  sub046C
End Sub

Sub sub096A ()
Dim l00A0 As Integer
Dim l00A2  As Integer
Dim l00A6 As T15B8
Dim l00A8 As Integer
' Const mc00B0 = 90 ' &H5A%
  sub0201 l00A8, gc03D6 & gc03E6
  Get gv088E, gv0DDA + mc00B0, l00A2
  gv0DE2 = fn0E69(l00A2)
  Get gv088E, fn0E71(gv0DE2), l00A6
  l00A2 = l00A6.M15E5
  gv0DE2 = fn0E69(l00A2)
  Do
    l00A0 = fn0E69(l00A2)
    Get gv088E, fn0E71(l00A0), l00A6
    gv7AA2 = gv7AA2 + 1
    gv0DE6(l00A0) = gc0E18
    gv0DE6(l00A0 + gc0E20) = gv7AA2
    l00A2 = l00A6.M15E5
    If l00A2 = &HFFFF Then Exit Do
    gv0DE6(l00A0 + gc0E22) = fn0E69(l00A2)
  Loop
  gv0DE6(l00A0 + gc0E22) = l00A2
  sub0220 l00A8
End Sub

Sub sub0979 ()
Dim l00B2 As Integer
Dim l00B4 As Integer
Dim l00B6 As Integer
Dim l00BC As T14DB
' Const mc00BE = 86 ' &H56%
  Get gv088E, gv0DDA + mc00BE, l00B4
  gv0DE4 = fn0E69(l00B4)
  Do While l00B4 <> &HFFFF
    l00B2 = fn0E69(l00B4)
    Get gv088E, fn0E71(l00B2), l00BC
    gv7AEE = gv7AEE + 1
    gv0DE6(l00B2 + gc0E24) = gv7AEE
    gv0DE6(l00B2 + gc0E2A) = (l00BC.M1511 \ &H100) And 7
    If l00BC.M152D And 1 Then
      gv0DE6(l00B2) = gc0E1A
      gv0DE6(l00B2 + gc0E2C) = l00BC.M1538
    Else
      gv0DE6(l00B2) = gc0E1C
      gv0DE6(l00B2 + gc0E2C) = -1
    End If
    l00B6 = fn0E69(l00BC.M12E2)
    If gv0DE6(l00B6) <> gc0E18 Then
      sub01F8
    End If
    gv0DE6(l00B2 + gc0E26) = l00B6
    l00B4 = l00BC.M1541
    If l00B4 = &HFFFF Then Exit Do
    gv0DE6(l00B2 + gc0E28) = fn0E69(l00B4)
  Loop
  gv0DE6(l00B2 + gc0E28) = l00B4
End Sub

Private Sub sub0988 ()
Dim l002C As Integer
Dim l002E As Integer
Dim l0030 As Integer
Dim l0032 As String
Dim l0034 As Integer
Dim l0036 As Integer
Dim l0038 As Integer
Dim l003C As T14DB
' Const mc0050 = 7 ' &H7%
' Const mc0052 = 6 ' &H6%
' Const mc005E = 256 ' &H100%
  If gv7AEE = 0 Then Exit Sub
  l0030 = fn0E7C(gv0DE4)
  Do While l0030 <> &HFFFF
    l002E = fn0E69(l0030)
    Get gv088E, fn0E71(l002E), l003C
    l002C = gv0DE6(l002E + gc0E24)
    gv7ABC(l002C).M24E0 = l002E
    gv7ABC(l002C).M1374 = l003C.M1374
    gv7ABC(l002C).M24AC = gv0DE6(l002E + gc0E26)
    If l003C.M1511 And 2 Then
      l0032 = "fn": l0038 = mc0050
      gv7ABC(l002C).M24F9 = gv0EE2((l003C.M1511 \ &H100) And 7)
    Else
      l0032 = "sub": l0038 = mc0052
    End If
    If l003C.M152D And 1 Then
      gv0DE6(l002E) = gc0E1A
      gv7ABC(l002C).M13E4 = l003C.M1538
      If l0030 = gv0F88 Then
        l0032 = "Main": l0038 = 0
      Else
        l0032 = l0032 & Hex$(l0030)
      End If
    Else
      gv0DE6(l002E) = gc0E1C
      gv7ABC(l002C).M13E4 = -1
      l0032 = "ext" & l0032 & Hex$(l0030)
      l0038 = l0038 + 2
    End If
    If l0038 Then l0032 = ""
    sub03A5 l002C, l0032, l0038
    l0030 = l003C.M1541
  Loop
  For l0034 = 3 To 6
    If (gv0784(l0034).M12D9 And mc005E) = 0 Then Exit For
    sub0951 l0034
  Next
  For l0034 = 1 To gv7AA2
    l002E = 0
    l0036 = gv7A70(l0034).M24AC
    For l002C = 1 To gv7AEE
      If gv7ABC(l002C).M24AC = l0036 Then
        If gv7ABC(l002C).M1C2B Then
          l002E = l002E + 1
          gv7ABC(l002C).M2503 = l002E
        End If
      End If
    Next
  Next
End Sub

