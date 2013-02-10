Attribute VB_Name = "Module14"
' Module14
Option Explicit
Const mc00E4 = 2048 ' &H800%
Const mc012E = 32 ' &H20%
Const mc0132 = 8192 ' &H2000%
Const mc0134 = 1024 ' &H400%
Const mc014E = 8 ' &H8%
Const mc0150 = 9 ' &H9%
Const mc0152 = 10 ' &HA%
Const mc0154 = 11 ' &HB%
Const mc0156 = 12 ' &HC%
Const mc0158 = 14 ' &HE%
Const mc015A = 10 ' &HA%
Const mc015C = 12 ' &HC%
Const mc015E = 512 ' &H200%
Const mc0160 = 256 ' &H100%
Const mc0162 = 8 ' &H8%
Const mc0164 = 9 ' &H9%
Const mc0166 = 13 ' &HD%
Const mc0168 = 11 ' &HB%
Const mc01F4 = 2 ' &H2%
Const mc01F6 = 1 ' &H1%
Const mc026A = 4 ' &H4%
Const mc026C = 8 ' &H8%
Const mc0284 = 2 ' &H2%
Const mc0288 = 1 ' &H1%
Const mc028E = 1 ' &H1%
Const mc0290 = 2 ' &H2%
Const mc02A2 = 3 ' &H3%
Const mc02A4 = 1 ' &H1%
Const mc02B6 = 2 ' &H2%
Const mc02CC = 1 ' &H1%

Function fn0CBF() As String
End Function

Function fn0CDA(p0224 As Integer) As String
Dim l0226 As Integer
Dim l0228 As String
  l0226 = fn0CE6(p0224)
  If l0226 > 0 Then
    fn0CDA = gv7876(l0226)
    Exit Function
  Else
    l0228 = gv793C(-l0226)
  End If
  If Left$(l0228, 1) = "g" Then
    fn0CDA = l0228 & fn02B5(fn033D(p0224, 2))
  Else
    fn0CDA = l0228 & fn02B5(gv098C(p0224))
  End If
End Function

Function fn0CE6(p0230 As Integer) As Integer
Dim l0232 As Integer
  If p0230 <= 0 Or p0230 > gv09C2 Then
    Exit Function
  End If
  l0232 = gv78A8(p0230)
  If l0232 Then fn0CE6 = l0232: Exit Function
  l0232 = Asc(Mid$(gv78DA, p0230, 1))
  If l0232 = 0 Or l0232 = gc7A18 Then
    If gv098C(p0230) Then
      l0232 = gc79E4
    Else
      l0232 = gc79E0
    End If
  End If
  Select Case l0232 And gc79C0
  Case gc79E0
    fn0CE6 = -gc78E2
  Case gc79E4, gc79E6
    fn0CE6 = -gc78E4
  Case Else
  End Select
End Function

Function fn0CF1(p0236 As Integer) As Integer
Dim l0238 As Integer
Dim l023A As Integer
On Error GoTo L18A24
  l023A = fn032D(gv09C4(p0236), 2)
  If l023A = 0 Or l023A > gv09C2 Then Exit Function
  l0238 = Asc(Mid$(gv7952, p0236, 1))
  Select Case l0238 And gc79C0
  Case gc79E2
    fn0CF1 = l023A
  Case gv79C2, gc79D6
    l0238 = Asc(Mid$(gv78DA, l023A, 1))
    If l0238 = 0 Then
      If Asc(Mid$(gv78DA, l023A - 1, 1)) = 0 Then
        Exit Function
      End If
      fn0CF1 = l023A
    Else
      l0238 = l0238 And gc79C0
      If l0238 = gc79E4 Or l0238 = gc79E0 Then
        fn0CF1 = l023A
      End If
    End If
  End Select
L18A24:
Exit Function
End Function

Function fn0CFD(ByVal pv023E As Integer) As Integer
  pv023E = fn032D(pv023E, 2)
  If pv023E < gv09C2 Then fn0CFD = Asc(Mid$(gv78DA, pv023E, 1)) = gc7A0A
End Function

Function fn0D0A(p0242 As Integer) As Integer
Dim l0244 As Integer
Dim l0246 As Integer
Dim l0248 As Integer
Dim l024A As Integer
Dim l024C As Integer
  l024C = gv7988(p0242)
  If l024C Then fn0D0A = l024C: Exit Function
  l0248 = fn033D(p0242, 2)
  l024A = gv09C4(p0242)
  l0246 = Asc(Mid$(gv7952, p0242, 1))
  Select Case l0246
  Case 0, gc7A18
    fn0D0A = -gc78F6
    Exit Function
  Case gc7A1A
    fn0D0A = -gc78F4
    l0244 = fn0E69(l024A)
    If l0244 > UBound(gv0DE6) Then
      Exit Function
    End If
    If gv0DE6(l0244) < mc01F4 Then
      Exit Function
    End If
    fn0D0A = -gc78FC
    Exit Function
  Case gc7A14
    fn0D0A = -gc7906
    Exit Function
  Case gc7A16
    fn0D0A = -gc78FA
    Exit Function
  End Select
  Select Case l0246 And gc79C0
  Case gc79DC
    If l024A Then sub01F8
    fn0D0A = -gc78F8
  Case gc79E2
    If l024A = 0 Then sub01F8
    fn0D0A = -gc78FA
  Case gc79D8
    fn0D0A = -gc7900
  Case gc79DE
    If l024A <= 0 Or (l024A And 1) > 0 Then
      fn0D0A = -gc78FE
    Else
      fn0D0A = -gc7902
    End If
  Case gc79E4
    fn0D0A = -gc7904
  Case gc79D6
    fn0D0A = -gc78F6
  Case Else
    fn0D0A = -gc78F4
  End Select
End Function

Function fn0D15(p0250 As Integer, p0252 As Integer, p0254 As Integer) As Integer
  Select Case p0254
  Case 0
  Case gc79D0
    sub01F8
  Case gc79D2
    fn0D15 = fn0C30(gv09C4(), p0250)
    Exit Function
  Case Else
    sub01F8
    fn0D15 = 2
    Exit Function
  End Select
  Select Case p0252 And gc79C0
  Case gv79C2: sub01F8
  Case gc79D6
  Case gc79E2
    fn0D15 = 1
  Case gc79DC, gc79E4
    fn0D15 = fn0D66(gv09C4(), p0250, p0252 And &H1F, p0254)
  Case gc79D8
    Select Case p0252 And &HF
    Case 6
      fn0D15 = 2
    Case Else
      fn0D15 = 1
    End Select
  Case gc79DE
    Select Case p0252 And &HF
    Case 6
      fn0D15 = 2
    Case 0 To gc79C8
      fn0D15 = 1
    Case gc79CA
      fn0D15 = 1
    Case Else
      sub01F8
      fn0D15 = 1
    End Select
  Case Else
    sub01F8
  End Select
End Function

Function fn0D20(p0258 As Integer) As String
On Error Resume Next
  fn0D20 = fn0D2B(gv7A70(p0258).M1374)
End Function

Function fn0D2B(p025C As Integer) As String
Dim l025E As Integer
Dim l0260 As Integer
  l025E = fn032D(p025C, 4)
  l0260 = gv790A(l025E)
  If l0260 = 0 Then
    l0260 = -gc78E0
    gv790A(l025E) = l0260
  End If
  If l0260 > 0 Then
    fn0D2B = gv7876(l0260)
  ElseIf l0260 < 0 Then
    fn0D2B = gv793C(-l0260) & fn02B5(p025C)
  Else
    gv790A(l025E) = -gc78E0
  End If
End Function

Function fn0D66(p02A8() As Integer, p02AE As Integer, p02B0 As Integer, p02B2 As Integer) As Integer
Dim l02B4 As Integer
' Const mc02B6 = 2 ' &H2%
  Select Case p02B2
  Case gc79D2
    fn0D66 = fn0C30(p02A8(), p02AE)
    Exit Function
  End Select
  Select Case p02B0
  Case 0
    fn0D66 = 0
  Case 1
    fn0D66 = 1
  Case 2, 3
    fn0D66 = 2
  Case 4
    fn0D66 = 4
  Case 5
    fn0D66 = 4
  Case 6
    fn0D66 = 8
  Case 7
    fn0D66 = 2
  Case gc79C6
    fn0D66 = fn032D(p02A8(p02AE - 1) + 1, 2)
  Case gc79CA
    l02B4 = p02A8(p02AE - 1)
    If Not fn0CFD(l02B4) Then
      sub022F mc00E4
      Exit Function
    End If
    l02B4 = fn032D(l02B4, 2)
    fn0D66 = fn032D(gv098C(l02B4 + mc02B6) + 1, 2)
  Case gc79C8
    fn0D66 = 2
  Case Else
    sub01F8
  End Select
End Function

Function fn0D70(p02BA As Integer) As String
  fn0D70 = fn0D2B(gv7ABC(p02BA).M1374)
End Function

Function fn0C4A(p020C As Integer, p020E As Integer) As String
Dim l0210 As Integer
  l0210 = gv7A70(p020C).M12C5
  l0210 = gv0FBC(l0210).M25D0
  If l0210 = 0 Then sub01F8
  fn0C4A = gv0FF6(l0210 + p020E).M1D68
End Function

Sub sub03A5(p0192 As Integer, p0194 As String, p0196 As Integer)
Dim l0198 As Integer
Dim l019A As Integer
  l0198 = gv7ABC(p0192).M1374
  If Len(p0194) Then
    If p0194 = fn0D70(p0192) Then Exit Sub
    sub0C83 l0198, p0194
  Else
    l019A = fn032D(l0198, 4)
    If l019A > UBound(gv790A) Then ReDim Preserve gv790A(l019A)
    gv790A(l019A) = -p0196
    p0194 = fn0D2B(l0198)
    gv091A = True
  End If
End Sub

Sub sub0441(p017C As Integer, p017E)
Dim l0180 As Integer
  l0180 = gv7A70(p017C).M1374
  sub0C83 l0180, p017E
End Sub

Sub sub04AA(p01B6 As Integer)
Dim l01B8 As Integer
Dim l01BA As Integer
Dim l01BC As Integer
Dim l01BE As Integer
Dim l01C0 As Integer
Dim l01C2 As Integer
Dim l01C4 As Integer
Dim l01C8 As T1369
Dim l01CC As T13C2
Dim l01D0 As T20E0
  l01BA = fn032D(p01B6, 2)   ' scaning types
  If Asc(Mid$(gv78DA, l01BA, 1)) Then
    p01B6 = 0
    Exit Sub
  End If
  Get gv088E, gv0B96 + fn0335(p01B6), l01C8
  sub04B5 l01BA, 4, gc7A0A, 0, 0
  sub0C96 gc78E6, l01BA
  l01B8 = l01C8.M1397
  If l01C8.M1380 And 1 Then
    l01BC = fn032D(l01B8, 2)
    l01B8 = gv098C(l01BC - 1)
  End If
  Do While l01B8
    l01BC = fn032D(l01B8, 2)
    Get gv088E, gv0B96 + fn0335(l01B8), l01CC
    l01C4 = l01BC + 4
    l01BE = l01CC.M13DD
    If l01BE > 9 Or l01BE < 0 Then
      l01BE = gc7A0C Or gc79CC
    Else
      l01BE = gc7A0C Or l01CC.M13DD
    End If
    If l01CC.M13CF And 1 Then
      l01BE = l01BE Or gc79CE
      l01C0 = 6 + (gv098C(l01BC + 6) And &HFF) * 2
      sub04B5 l01C4, l01C0, 0, gc79D2, 0
      l01C4 = l01C4 + l01C0
    End If
    If Asc(Mid$(gv78DA, l01BC - 1, 1)) = 0 Then
      l01BE = l01BE Or gc79C4
      If l01CC.M13DD = 8 Then
        l01C2 = gc79C6
      ElseIf l01CC.M13DD = 0 Then
        sub01F8
      Else
        sub01F8
      End If
    Else
      l01C2 = 0
    End If
    sub04B5 l01BC, 4, l01BE, l01C2, 0
    sub0C96 gc78E8, l01BC
    l01B8 = l01CC.M13CF And -2
  Loop
  If l01C8.M1380 And 1 Then
    l01B8 = l01C8.M1397
    l01BC = fn032D(l01B8, 2) - 1
    If l01BC <> l01C4 Then sub01F8
    Get gv088E, gv0B96 + fn0335(l01B8), l01D0
    l01C0 = (6 + l01D0.M1EF7 * 3) \ 2
    sub04B5 l01BC, l01C0, 0, gc79D0, 0
    l01C4 = l01BC + l01C0
  End If
  Mid$(gv78DA, l01C4 - 1, 1) = Chr$(gc7A0E)  'E2
  p01B6 = l01C8.M1380 And -2
End Sub

Sub sub04B5(p010A As Integer, p010C As Integer, p010E As Integer, p0110 As Integer, p0112 As Integer)
Dim l0114 As Integer
Dim l0116 As Integer
Dim l0118  As Integer
Dim l011A  As Integer
  If p010A < 3 Or p010A > gv09C2 Then sub01F8: Exit Sub
  l0114 = p010E Or p0110 Or (p0112 And gc79CE)
  l0116 = Asc(Mid$(gv78DA, p010A, 1))
  If l0116 >= gc79E6 Then
    If (l0114 Xor l0116) And &H1F Then
      p0110 = l0116 And &HF
      If l0116 And gc79CE Then p0112 = gc79D2 Else p0112 = 0
    End If
    Exit Sub
  End If
  If p010E = gc79E0 And p0112 = 0 Then
    Select Case p0110
    Case 1 To 6
      If p010C = 0 Then p010C = fn0D66(gv098C(), p010A, p0110, p0112)
      If p010C > 0 Then
        For l0116 = p010A To p010A + p010C - 1
          If gv098C(l0116) Then
            p010E = gc79E4
            Exit For
          End If
        Next
      End If
    End Select
  End If
  sub0C76 l0118, l011A, p010E%, p0110%, p010A%, True
  If p010E Then
    If p010C = 0 Then p010C = fn0D66(gv098C(), p010A, p0110, p0112)
    If p0110 And gc79C4 Then
      If gv2382 > 0 Then
        If Asc(Mid$(gv78DA, p010A - 1, 1)) Then
          If Asc(Mid$(gv78DA, p010A - 1, 1)) <> p0110 Then sub01F8
        End If
      End If
      Mid$(gv78DA, p010A - 1, 1) = Chr$(p0110)
    End If
    Mid$(gv78DA, p010A, 1) = Chr$(p010E Or p0110 Or (p0112 And gc79CE))
    If p010E > gc79D6 And p0110 Then
      If gv78A8(p010A) = 0 Then
        sub0C63 p010A
      End If
    End If
  Else
    Mid$(gv78DA, p010A, 1) = Chr$(p0110)
  End If
  If p010C > 1 Then
    Mid$(gv78DA, p010A + 1, p010C - 1) = String$(p010C - 1, gc79D0)
    If p0112 Then Mid$(gv78DA, p010A + 1, 1) = Chr$(p0112)
  End If
  gv091A = True
End Sub

Sub sub04C3(p0102 As String, p0104 As Integer)
Dim l0106 As Integer
  If p0104 > gv09C2 Then
    p0102 = "???"
    Exit Sub
  End If
  l0106 = gv78A8(p0104)
  If Len(p0102) = 0 Then
    sub0C63 p0104
    p0102 = fn0CDA(p0104)
    Exit Sub
  End If
  If l0106 < 0 Then sub0C21 p0102, gv78A8(p0104)
End Sub

Sub sub07BC(p0278() As Integer, p027E As Integer)
Dim l0280 As Integer
Dim l0282 As String
' Const mc0284 = 2 ' &H2%
Static m0286 As Integer
' Const mc0288 = 1 ' &H1%
  Select Case p0278(p027E - 1) And &HFFC0
  Case &HC000, &HC500, &HC600
    l0280 = 0
  Case &HC100, &HC200
    l0280 = p0278(p027E - 2)
    If l0280 And &H4000 Then
      l0280 = l0280 And &HFF
    Else
      l0280 = 0
    End If
  Case &H4600
    l0280 = 0
  Case Else
    l0280 = p0278(p027E + mc0284)
    If (l0280 And &H7FC0) = &H5000 Then
      l0280 = l0280 And &HFF
    Else
      sub022F mc012E
      Exit Sub
    End If
  End Select
  l0282 = ")"
  p027E = p027E + 6
  While l0280 > 0
    l0282 = Format$(p0278(p027E + m0286) + p0278(p027E + mc0288) - 1) & l0282
    If p0278(p027E + mc0288) Then
      l0282 = Format$(p0278(p027E + mc0288)) & " To " & l0282
    End If
    l0280 = l0280 - 1
    If l0280 Then l0282 = ", " & l0282
    p027E = p027E + 2
  Wend
  sub05FF "(" & l0282
End Sub

Sub sub07C9(p0292 As Integer, p0294 As Integer, p0296 As Integer)
  Select Case p0292 And &HF
  Case 0
  Case gc79CA
    sub05FF " As " & fn0CDA(fn032D(p0294, 2))
  Case gc79C6
    sub05FF " As String *" & Str$(fn0335(p0294))
  Case Is < 8
    sub05FF " As " & gv1594(p0292 And &HF)
  Case gc79C8:
    If gv0FBC(p0294 And &HFF).M25C6 < 0 Then sub05FF "()"
    sub05FF " As "
    If p0296 And p0294 And &H8000 Then sub05FF "New "
    sub05FF gv0FBC(p0294 And &HFF).M1D68
  Case Else
    If fn0CFD(p0292) Then
      sub05FF " As " & fn0CDA(fn032D(p0292, 2))
    Else
      sub01F8
      sub05FF " As " & Hex$(p0292)
    End If
  End Select
End Sub

Sub sub07EB(p013A As Integer, p013C As Integer, p013E As Integer)
Dim l0140 As Integer
Dim l0142 As Integer
Dim l0144 As Integer
Dim l0146 As Integer
Dim l0148 As Integer
Dim l014A  As Integer
' Const mc014E = 8 ' &H8%
' Const mc0150 = 9 ' &H9%
' Const mc0152 = 10 ' &HA%
' Const mc0154 = 11 ' &HB%
' Const mc0156 = 12 ' &HC%
' Const mc0158 = 14 ' &HE%
' Const mc015A = 10 ' &HA%
' Const mc015C = 12 ' &HC%
' Const mc015E = 512 ' &H200%
' Const mc0160 = 256 ' &H100%
' Const mc0162 = 8 ' &H8%
' Const mc0164 = 9 ' &H9%
' Const mc0166 = 13 ' &HD%
' Const mc0168 = 11 ' &HB%
  If p013C < 8 Then Exit Sub
  l0144 = fn032D(p013A, 2)
  If l0144 > gv09FA Then
    Exit Sub
  End If
  l0140 = Asc(Mid$(gv7952, l0144, 1))
  l0142 = l0140 And gc79C0
  l0140 = l0140 Xor l0142
    If gv7988(l0144) Then
      sub0C56 l0144, l0142
      Exit Sub
    End If
  If l0140 And gc79CE Then Exit Sub
  If l0140 > gc79E6 Then Exit Sub
  If l0142 <= gc79D6 Then
    Select Case p013C
    Case mc014E
      l0146 = gc79DE
    Case mc0150
      l0146 = gc79DC
    Case mc0152
      l0146 = gc79E2
    Case mc0154
      l0146 = gc79D8
    Case mc0156
      If l0142 > gc79D6 Then Exit Sub
      l0146 = gc79D6
    Case mc0158
      If p013E = mc015A Then
        sub0A10 l0144
        Exit Sub
      ElseIf p013E = mc015C Then
        l0146 = gc79E6
      End If
      If gv09C4(l0144) = 0 Then
        sub022F mc015E
        l0146 = gc79DC
      ElseIf gv09C4(l0144) And &H8001 Then
        l0146 = gc79DE
      Else
        l0146 = gc79D8
      End If
    Case Else
      sub01F8
      If l0142 Then Exit Sub
      l0146 = gc79D6
    End Select
    sub0C56 l0144, l0146
  Else
    l0146 = l0142
  End If
  Select Case p013E
  Case 0
    sub022F mc0160
  Case mc0162
    l0148 = 0
  Case mc0164
    l0148 = gc79C8
  Case mc015A
    sub0A10 l0144
    Exit Sub
  Case mc0166: l0148 = gc79C6
  Case Is < 8
    l0148 = p013E
  Case mc0168
    l0148 = gc79CA
  Case mc015C
    sub09E5 l0144, 1, gc79E6, gc79EC, 0
    Exit Sub
  Case Else
    sub01F8
    Exit Sub
  End Select
  sub0A05 l0144, l0146, l0148, l014A
End Sub

Sub sub09D4(p00DA As Integer)
Dim l00DC As Integer
' Const mc00E4 = 2048 ' &H800%
  l00DC = gv098C(p00DA + 2)
  If (l00DC And &HFFF0) <> &HC100 Then
    sub022F mc00E4
    Exit Sub
  End If
  l00DC = l00DC And &HF
  If l00DC = 8 Then
  ElseIf l00DC = 0 Then
    l00DC = gc79CA
    If Not fn0CFD(gv098C(p00DA - 1)) Then
      sub022F mc00E4
      Exit Sub
    End If
  ElseIf l00DC <= 9 Then
  Else
    sub01F8
    Exit Sub
  End If
  sub0CCF p00DA, gc79E0, l00DC, gc79D2
End Sub

Sub sub09E5(p016A As Integer, p016C As Integer, p016E As Integer, p0170 As Integer, p0172 As Integer)
Dim l0174  As Integer
Dim l0176  As Integer
Dim l0178   As Integer
  If p016E = gc79DC And p0172 = 0 Then
    Select Case p0170
    Case 1 To 6
      If p016C = 0 Then p016C = fn0D15(p016A, p016E Or p0170, p0172)
      If p016C > 0 Then
        For l0174 = p016A To p016A + p016C - 1
          If gv09C4(l0174) Then
            p016E = gc79E4
            Exit For
          End If
        Next
      End If
    End Select
  End If
  sub0C76 l0176, l0178, p016E, p0170, p016A, False
  If p016E Then
    If p016C = 0 Then p016C = fn0D15(p016A, p016E Or p0170, p0172)
    If p0170 And gc79C4 Then
      If Asc(Mid$(gv7952, p016A - 1, 1)) Then
        If Asc(Mid$(gv7952, p016A - 1, 1)) <> p0170 Then
          If Asc(Mid$(gv7952, p016A - 1, 1)) <> gc79D0 Then
            If gv2382 > 0 Then
              sub01F8
              Exit Sub
            End If
          End If
        End If
      End If
      Mid$(gv7952, p016A - 1, 1) = Chr$(p0170)
    End If
    If p016E = gc79E2 Then sub04B5 fn032D(gv09C4(p016A), 2), 0, gc79E0, p0170, p0172
    Mid$(gv7952, p016A, 1) = Chr$(p016E Or p0170 Or (p0172 And gc79CE))
  Else
    Mid$(gv7952, p016A, 1) = Chr$(p0170)
  End If
  gv7988(p016A) = fn0D0A(p016A)
  If p016C > 1 Then
    Mid$(gv7952, p016A + 1, p016C - 1) = String$(p016C - 1, gc79D0)
    If p0172 Then Mid$(gv7952, p016A + 1, 1) = Chr$(p0172)
  End If
  gv091C = True
End Sub

Sub sub09F3(p01E4 As String, p01E6 As Integer)
Dim l01E8 As Integer
Dim l01EA As Integer
' Const mc01F4 = 2 ' &H2%
' Const mc01F6 = 1 ' &H1%
  l01EA = Asc(Mid$(gv7952, p01E6, 1))
  Select Case l01EA And gc79C0
  Case gv79C2, gc79E6
    Select Case l01EA
    Case gc7A1A
      l01E8 = fn0E69(gv09C4(p01E6))
      If gv0DE6(l01E8) = mc01F4 Then
        If p01E4 = fn0D70(gv0DE6(l01E8 + mc01F6)) Then Exit Sub
        Stop
        sub03A5 gv0DE6(l01E8 + mc01F6), p01E4, 0
        Exit Sub
      End If
    Case gc7A14
      If gv2382 <> gc79BC Then
        sub01F8
      Else
        GoTo L1813A
      End If
    Case Else
      sub01F8
    End Select
  Case gc79E2
    If gv09C4(p01E6) Then
      sub04C3 p01E4, fn032D(gv09C4(p01E6), 2)
    Else
      sub01F8
    End If
  Case gc79E4, gc79DC, gc79DE, gc79D8
L1813A:
    l01EA = UBound(gv7956) + 1
    ReDim Preserve gv7956(l01EA)
    gv7956(l01EA) = p01E4
    gv7988(p01E6) = l01EA
  Case Else
    sub01F8
  End Select
End Sub

Sub sub0A05(p0120 As Integer, p0122 As Integer, p0124 As Integer, p0126 As Integer)
Dim l0128 As Integer
Dim l012A As Integer
Dim l012C As Integer
' Const mc012E = 32 ' &H20%
' Const mc0132 = 8192 ' &H2000%
' Const mc0134 = 1024 ' &H400%
  If (p0122 And &H1F) Then
    sub022F mc012E
    p0124 = (p0124 Or p0122) And &HF
    p0122 = p0122 And gc79C0
  End If
  If gv7A2A Then
    Select Case p0122
    Case gv79C2, gc79E6
      sub022F mc0132
      Exit Sub
    Case gc79D6
    Case gc79DC
      If gv09C4(p0120) Then
        p0122 = gc79E4
      End If
    Case gc79DE, gc79D8, gc79E4
      If gv09C4(p0120) = 0 Then
        sub022F mc0134
        p0122 = gc79DC
      End If
    Case gc79E2
      If fn0335(gv09C4(p0120)) > gv09BE Or (gv09C4(p0120) And 1) Then
      Else
        sub09E5 p0120, 0, gc79E2, p0124, p0126
      End If
      Exit Sub
    Case Else
      sub022F mc012E
      Exit Sub
    End Select
  End If
  l0128 = Asc(Mid$(gv7952, p0120, 1))
  l012A = p0124 Or (p0126 And gc79CE)
  If (p0122 Or l012A) = l0128 Then
    Exit Sub
  End If
  l012C = l0128 And gc79C0
  l0128 = l0128 Xor l012C
  If p0122 < l012C Then p0122 = l012C
  If p0122 = gc79E0 Then
    l0128 = Asc(Mid$(gv78DA, fn032D(gv09C4(p0120), 2), 1))
    l012C = l0128 And gc79C0
    l0128 = l0128 Xor l012C
    If p0122 < l012C Then p0122 = l012C
  End If
  If l012A = 0 Then
    l012A = l0128
  ElseIf l0128 Then
    If l012A <> l0128 Then
      sub022F mc00E4
    End If
    Exit Sub
  End If
  If (l012A Or p0122) = (l0128 Or l012C) Then
    Exit Sub
  End If
  sub09E5 p0120, 0, p0122, l012A, p0126
End Sub

Sub sub0A10(p00E6 As Integer)
Dim l00E8 As Integer
Dim l00EA As Integer
  l00E8 = gv09C4(p00E6 + 2) And &HF
  If l00E8 = 0 Then l00E8 = gc79CA
  Select Case gv09C4(p00E6 + 2) And &HF00
  Case 0: l00EA = gc79E0
    If gv09C4(p00E6) = 0 Then sub01F8
    sub09D4 fn032D(gv09C4(p00E6), 2)
  Case &H100: l00EA = gc79DC
    If gv09C4(p00E6) Then sub01F8
  Case &H200: l00EA = gc79DC
    If gv09C4(p00E6) Then sub01F8
  Case &H500: l00EA = gc79D8
    If gv09C4(p00E6) < 6 Then sub01F8
  Case &H600: l00EA = gc79DE
    If gv09C4(p00E6) >= 0 Then sub01F8
  Case Else
    sub01F8
  End Select
  sub0C56 p00E6, l00EA
  sub09E5 p00E6, 0, l00EA, l00E8, gc79D2
End Sub

Function fn0C30(p02C2() As Integer, p02C8 As Integer) As Integer
Dim l02CA As Integer
' Const mc02CC = 1 ' &H1%
  l02CA = p02C2(p02C8 + 2)
  Select Case l02CA And &HFFF0
  Case &HC000, &HC500, &HC600
    fn0C30 = 3
  Case &HC100, &HC200
    l02CA = p02C2(p02C8 + 1)
    If l02CA And &H4000 Then
      l02CA = l02CA And &HF
    Else
      l02CA = 8
    End If
    fn0C30 = l02CA * 2 + 9
  Case &H4600
    fn0C30 = 3
  Case Else
    sub022F mc02CC
    Exit Function
  End Select
End Function

Sub sub0C21(ByVal p02BC As String, p02BE As Integer)
  If p02BE <= 0 Then
    p02BE = UBound(gv7876) + 1
    ReDim Preserve gv7876(p02BE)
  Else
    If p02BC = gv7876(p02BE) Then Exit Sub
  End If
  gv7876(p02BE) = p02BC
  gv091A = True
End Sub

Sub sub0C3D(p01FC As Integer, p0200 As ComboBox, p0204 As ComboBox, p0208 As ComboBox)
  p0200.AddItem "<special>"
  p0200.AddItem "<refd>"
  If p01FC Then
    p0200.AddItem "?"
    p0200.AddItem "?"
    p0200.AddItem "Member"
  Else
    p0200.AddItem "Param"
    p0200.AddItem "Modul"
    p0200.AddItem "Local"
  End If
  p0200.AddItem "Global"
  p0200.AddItem "Const"
  p0200.AddItem "<fix>"
  p0204.AddItem ""
  p0204.AddItem "%"
  p0204.AddItem "&"
  p0204.AddItem "!"
  p0204.AddItem "#"
  p0204.AddItem "@"
  p0204.AddItem "Variant"
  p0204.AddItem "$"
  p0204.AddItem "-$*"
  p0204.AddItem "-Object"
  p0204.AddItem "-Type"
  p0204.AddItem "+MType"
  p0208.AddItem ""
  p0208.AddItem "+"
  p0208.AddItem "+()"
End Sub

Sub sub0C63(p01D6 As Integer)
Dim l01D8 As Integer
  l01D8 = gv78A8(p01D6)
  If l01D8 Then Exit Sub
  gv78A8(p01D6) = fn0CE6(p01D6)
End Sub

Private Sub sub0C76(p01A4 As Integer, p01A6 As Integer, p01A8 As Integer, p01AA As Integer, p01AC As Integer, p01AE As Integer)
Dim l01B0 As Integer
Dim l01B2 As Integer
  If gv2382 = 0 Then
    If gv0BC6 Then sub01F8
    Exit Sub
  ElseIf gv2382 = gc79BA Then
    If p01AE Then p01A4 = gv2382
    p01A6 = gv2382
    Exit Sub
  End If
  If p01AE Then
    l01B0 = Asc(Mid$(gv78DA, p01AC, 1))
  Else
    l01B0 = Asc(Mid$(gv7952, p01AC, 1))
  End If
  l01B2 = l01B0 And &H1F
  l01B0 = l01B0 Xor l01B2
  l01B2 = l01B2 And &HF
  Select Case l01B0
  Case gv79C2
    p01A4 = gv2382
  Case gc79D6
    If p01A8 > gc79D6 Then p01A4 = gv2382
  Case gc79E4
    If p01A8 <> l01B0 Then
      If p01A4 <> gc79BC And p01A8 <> gc79DC And p01A8 <> gc79E0 Then
        GoSub L17A90
      End If
      p01A8 = l01B0
    End If
  Case gc79E6
      p01A8 = l01B0
      Exit Sub
    If p01A8 <> l01B0 Then
      GoSub L17A90
      p01A8 = l01B0
      Exit Sub
    End If
  Case Else
    If p01A8 <> l01B0 Then
      If p01A8 <> gc79E4 Then
        GoSub L17A90
        p01A8 = l01B0
      End If
    End If
  End Select
  If p01AA Then
    If l01B2 = 0 Then
      p01A6 = gv2382
    ElseIf p01AA <> l01B2 Then
      GoSub L17A90
    End If
  Else
    p01AA = l01B2
  End If
Exit Sub

L17A90:
  sub022F mc00E4
Return
End Sub

Sub sub0C83(p018A As Integer, ByVal p018C As String)
Dim l018E As Integer
Dim l0190 As Integer
  l018E = fn032D(p018A, 4)
  If l018E > UBound(gv790A) Then ReDim Preserve gv790A(l018E)
  l0190 = gv790A(l018E)
  If l0190 <= 0 Then
    sub0C21 p018C, gv790A(l018E)
  ElseIf gv7876(l0190) <> p018C Then
    gv091A = True
    gv7876(l0190) = p018C
  End If
End Sub

Sub sub0C96(p01DA As Integer, p01DC As Integer)
Dim l01DE As Integer
  l01DE = gv78A8(p01DC)
  If l01DE > 0 Then
    Exit Sub
  ElseIf l01DE < 0 Then
    If l01DE = -p01DA Then Exit Sub
  End If
  gv78A8(p01DC) = -p01DA
  gv091A = True
End Sub

Sub sub0CAC(p01E0 As Integer, p01E2 As Integer)
  gv7988(p01E0) = p01E2
End Sub

Sub sub0CCF(p00F4 As Integer, p00F6 As Integer, p00F8 As Integer, p00FA As Integer)
Dim l00FC As Integer
Dim l00FE As Integer
Dim l0100 As Integer
  If p00F6 And &H1F Then
    sub01F8
    p00F8 = p00F6 Or p00F8
    p00F6 = p00F6 And gc79C0
    p00F8 = p00F8 Xor p00F6
  End If
  If p00F6 = 0 Then
    sub01F8
    Exit Sub
  End If
  l00FC = Asc(Mid$(gv78DA, p00F4, 1))
  If l00FC = 0 And gv098C(p00F4) Then
    p00F6 = gc79E4
  End If
  l00FE = p00F8 Or (p00FA And gc79CE)
  If (p00F6 Or l00FE) = l00FC Then
    Exit Sub
  End If
  l0100 = l00FC And gc79C0
  l00FC = l00FC Xor l0100
  If p00F6 < l0100 Then p00F6 = l0100
  If l00FE = 0 Then
    l00FE = l00FC
  ElseIf l00FC And l00FC <> l00FE Then
    Exit Sub
  End If
  If (l00FE Or p00F6) = (l00FC Or l0100) Then
    Exit Sub
  End If
  sub04B5 p00F4, 0, p00F6, p00F8, p00FA
  If p00F8 And gv78A8(p00F4) = 0 Then
    sub0C63 p00F4
  End If
End Sub

Sub sub0D3A()
Dim l0262 As Integer
' Const mc026A = 4 ' &H4%
' Const mc026C = 8 ' &H8%
  If gv06BE <= 3 Then
    sub022F gc038E
    gv7B30 = gv7B30 Or mc026A
    gv7B32 = mc026C
    Exit Sub
  End If
  sub049B
  gv0BA6 = ""
  If iVBVersion = 3 Then
    sub040D &H8003
    If gv0DD4 Then
      Get gv088E, , l0262
      gv0BA6 = Space$(l0262)
      Get gv088E, , gv0BA6
    End If
  ElseIf iVBVersion <> 2 Then
    sub022F mc026A
    Exit Sub
  End If
  gv7A2A = True
  sub095F
  sub015B
  gv7A2A = False
  gv7AB6 = 0
End Sub

Sub sub0D47(p028A As Long)
Dim l028C As Integer
' Const mc028E = 1 ' &H1%
' Const mc0290 = 2 ' &H2%
  l028C = p028A \ 2
  sub05FF "  " & fn0CDA(l028C)
  If gv098C(l028C + mc028E) And 1 Then
    sub07BC gv098C(), l028C + 4
  End If
  If gv098C(l028C + mc0290) >= gc79C8 Or gv098C(l028C + mc0290) < 0 Then
    sub07C9 gc79CA, gv098C(l028C + mc0290), False
  Else
    sub07C9 gv098C(l028C + mc0290), gv098C(l028C - 1), False
  End If
  sub06EA
  p028A = gv098C(l028C + mc028E) And &HFFFE&
End Sub

Sub sub0D56(p029E As Integer)
Dim l02A0 As Long
' Const mc02A2 = 3 ' &H3%
' Const mc02A4 = 1 ' &H1%
  sub05FF "Type " & fn0CDA(p029E): sub06EA
  l02A0 = fn0335(gv098C(p029E + mc02A2))
  If gv098C(p029E + mc02A4) And 1 Then
    l02A0 = fn0335(gv098C((l02A0& \ 2) - 1))
  End If
  While l02A0
    sub0D47 l02A0
  Wend
  sub05FF "End Type": sub06EA: sub06EA
End Sub

