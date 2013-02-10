' Module16
Option Explicit
Const mc000C = -1' &HFFFF%
Dim m000E() As Integer
Dim m0040 As Integer
Dim m0042 As Integer
Dim m0044 As Integer
Const mc0076 = 8 ' &H8%
Const mc0078 = 9 ' &H9%
Const mc007A = 64 ' &H40%
Const mc0086 = -2' &HFFFE%
Const mc008C = 5 ' &H5%
Const mc009E = 227 ' &HE3%
Const mc00BA = 15 ' &HF%
Const mc00BE = 1 ' &H1%
Const mc00C2 = 21 ' &H15%
Const mc00CA = 14 ' &HE%
Const mc0108 = 2048 ' &H800%
Const mc010C = 2 ' &H2%
Const mc0116 = 15 ' &HF%
Const mc0118 = 13 ' &HD%
Const mc011A = 48 ' &H30%
Const mc012C = 8 ' &H8%
Const mc012E = 32 ' &H20%
Const mc0130 = 8192 ' &H2000%
Const mc0132 = 128 ' &H80%
Const mc013C = 224 ' &HE0%
Const mc013E = 96 ' &H60%
Const mc0140 = 224 ' &HE0%
Const mc0142 = 192 ' &HC0%
Const mc0144 = 160 ' &HA0%
Const mc0146 = 32 ' &H20%
Const mc014E = 2 ' &H2%
Const mc0150 = 3 ' &H3%
Const mc015C = 229 ' &HE5%
Const m0160 = 0 ' &H0%
Const mc0162 = 225 ' &HE1%
Const mc0164 = 226 ' &HE2%
Const mc0166 = 233 ' &HE9%
Const mc016E = 231 ' &HE7%
Const mc0184 = 160 ' &HA0%
Const mc01B0 = 12 ' &HC%
Const mc01B2 = 32 ' &H20%
Const mc01B4 = 2 ' &H2%
Const mc01B6 = 6 ' &H6%
Const mc01BA = 7 ' &H7%
Const mc01BC = 8 ' &H8%
Const mc01BE = 9 ' &H9%
Const mc01C0 = 11 ' &HB%
Const mc01DC = 16 ' &H10%

Sub sub03B4 (p0046 As Integer, p0048 As Integer, p004A As Integer)
  gv7ABC(p0046).M24E7 = p0048
  gv7ABC(p0046).M24F0 = p004A
End Sub

Sub sub0A21 (p0176 As Integer)
Dim l0178  As Integer
Dim l017A As Integer
Dim l017C As Integer
' Const mc0184 = 160 ' &HA0%
  If p0176 <> gv7AB6 Then sub01F8: sub01A8 p0176
  sub0DFB gv7AB8
  If gv7A2A Then
    m0044 = gv09FA
    m0042 = m0044
    Exit Sub
  End If
  If m0040 Then
    m0044 = gv7ABC(m000E(1)).M250C
  Else
    m0044 = gv09FA
  End If
  gv7874 = True
  If mc000C Then sub05FF "Option Explicit": sub06EA
  For l0178 = 1 To gv7AEE
    If gv7ABC(l0178).M24AC = gv7AB8 Then
      If gv7ABC(l0178).M1C2B = 0 Then
        sub07A3 l0178
      End If
    End If
  Next
  If gv7AA6.M12C5 Then
    l0178 = &H1A \ 2
  Else
    l0178 = 3
  End If
  Do While l0178 < gv09FA
    l017A = gv09FC(l0178)
    l017C = Asc(Mid$(gv7952, l0178, 1))
    Select Case l017C And mc013C
    Case m0160, mc0140
      Select Case l017C
      Case 0
        While Asc(Mid$(gv7952, l0178, 1)) = 0
          l0178 = l0178 + 1
        Wend
      Case mc015C
        Exit Do
      Case mc009E
        l0178 = l0178 + 1
      Case 8 To &H1F
        l0178 = l0178 + 1
      Case mc0166
        l0178 = l0178 + 2
      Case mc016E
        l0178 = l0178 + 1
      Case Else
        sub01F8
        l0178 = l0178 + 1
      End Select
    Case mc0146
      sub01F8
      If l017A = 0 Then sub05FF "Dim " & fn0DCA(l0178) & " ' As ?": sub06EA
      l0178 = l0178 + 1
    Case mc0142
      sub0795 l0178, gv7952, False: sub06EA
      l0178 = l0178 + 1
    Case mc0184
      l0178 = l0178 + 1
    Case mc013E
    If l017A = 0 Then
      sub05FF "Dim "
      sub07B0 l0178, gv09C4(), gv7952, False
      sub06EA
    End If
      l0178 = l0178 + 1
    Case mc0132, mc007A
      l0178 = l0178 + 1
    Case Else
      sub01F8
      l0178 = l0178 + 1
    End Select
  Loop
End Sub

Sub sub0A33 ()
Dim l011E As Integer
Dim l0122  As Integer
Dim l0124 As Integer
Dim l0126 As Integer
Dim l0128 As Integer
Dim l012A  As Integer
' Const mc012C = 8 ' &H8%
' Const mc012E = 32 ' &H20%
' Const mc0130 = 8192 ' &H2000%
' Const mc0132 = 128 ' &H80%
' Const mc013C = 224 ' &HE0%
' Const mc013E = 96 ' &H60%
' Const mc0140 = 224 ' &HE0%
' Const mc0142 = 192 ' &HC0%
' Const mc0144 = 160 ' &HA0%
' Const mc0146 = 32 ' &H20%
  gv2382 = mc0086
  l011E = m0044 - 1
  If Asc(Mid$(gv7952, m0044, 1)) And mc012C Then l011E = l011E - 1
  For l0122 = 1 To l011E: gv09FC(l0122) = 0: Next
  For l0122 = 1 To m0040
    l011E = 0
    For l012A = gv09FA To 1 Step -1
      If gv09FC(l012A) = l0122 Then l011E = l012A: Exit For
    Next
    If l011E Then
      Do While l011E > 3
        l011E = l011E - 1
        l0124 = gv09FC(l011E)
        If l0124 And l0124 <> l0122 Then
          For l012A = 1 To l011E: gv09FC(l012A) = 0: Next
          If l011E >= m0044 Then
            m0044 = l011E + 1
          End If
          Exit For
        End If
      Loop
      Exit For
    End If
  Next
  m0042 = gv09FA
  For l0122 = m0040 To 1 Step -1
    l0126 = m000E(l0122)
    If gv7ABC(l0126).M13E4 <> -1 Then
      l0128 = 0
      For l012A = 1 To gv09FA
        If gv09FC(l012A) = l0122 Then l0128 = l012A: Exit For
      Next
      If l0128 = 0 Then
        gv7ABC(l0126).M250C = m0042
      Else
        m0042 = l0128
        gv7ABC(l0126).M250C = l0128
      End If
      l0128 = gv7ABC(l0126).M24F9
      If l0128 Then
        If gv09FC(m0042) > l0122 Then
          m0042 = m0042 - 1
        End If
        l0124 = gv09C4(m0042)
        Do While l0124 <> l0128
          GoSub L1AFC0
          If l0124 = l0122 Then
            If gv0DE6(gv7ABC(l0126).M24E0 + mc008C) = 6 Then
              m0042 = m0042 - 1
              GoSub L1AFC0
            End If
          End If
          If l0124 = l0122 Then
            gv7ABC(l0126).M250C = m0042
            gv09FC(m0042) = l0122
            l0124 = l0128
          ElseIf l0124 Then
            sub022F mc012E
            l0124 = 0
            Exit Do
          Else
            m0042 = m0042 - 1
            l0124 = gv09C4(m0042)
            If m0042 < m0044 Then
              sub022F mc0130
              l0124 = 0
              Exit Do
            End If
          End If
        Loop
        If l0124 = l0128 Then
          l0124 = gv7ABC(l0126).M24E0
          sub09E5 m0042, 0, mc0132, gv0DE6(l0124 + mc008C), 0
          If gv7988(m0042) <> -mc00C2 Then
            gv7988(m0042) = -mc00C2
            gv091C = True
          End If
        Else
          sub022F mc0130
        End If
      End If
    End If
  Next
  If m0042 <> m0044 Then
    m0044 = m0042
  End If
  l0128 = 0
  For l011E = m0044 To gv09FA
    l0124 = gv09FC(l011E)
    If l0124 > m0040 Then
      l0124 = 0
      gv09FC(l011E) = l0124
    End If
    l0122 = Asc(Mid$(gv7952, l011E, 1))
    If l0124 Then
      l0126 = m000E(l0124)
      If l0124 < l0128 Then
        sub01F8
        For l012A = 1 To l011E - 1: gv09FC(l012A) = 0: Next
        l0128 = l0124
      ElseIf l0124 > l0128 Then
        m0042 = l011E
        Do
          l0128 = l0128 + 1
          gv7ABC(m000E(l0128)).M250C = m0042
        Loop While l0128 < l0124
      End If
      Select Case l0122 And mc013C
      Case mc007A
        If gv7ABC(l0126).M24E7 Then
          sub0DD6 l011E, l0126
        End If
      Case mc0132
      Case mc013E
      Case mc0140
      Case mc0142
      Case mc0144
      Case Else
        sub022F mc0130
      End Select
    Else
      If l0122 >= mc0146 Then gv09FC(l011E) = l0128
    End If
  Next
Exit Sub

L1AFC0:
  l0124 = gv09C4(m0042)
  If l0124 And l0124 <> l0128 Then
    l0124 = 0
  Else
    l0124 = gv09FC(m0042)
    If l0124 = 0 Then
    ElseIf l0124 <> l0122 Then
      sub022F mc012E
      Return
    End If
    l011E = Asc(Mid$(gv7952, m0042, 1))
    If l011E And (l011E And mc013C) <> mc0132 Then
      If l0124 = l0122 Then l0124 = 0
    Else
      If l0124 = 0 Then l0124 = l0122
    End If
  End If
Return
End Sub

Sub sub0A44 ()
  m0042 = gv7ABC(gv7B2A).M250C
  If m0042 = 0 Then
    m0042 = gv09FA
  End If
  gv1140 = 0
End Sub

Sub sub0A56 ()
End Sub

Sub sub0A78 (p00F6 As T265D, p00F8 As Integer, p00FA As String)
Dim l00FC As String
Dim l0100 As Integer
Dim l0102 As Integer
Dim l0104 As Integer
Dim l0106 As Integer
' Const mc0108 = 2048 ' &H800%
' Const mc010C = 2 ' &H2%
' Const mc0116 = 15 ' &HF%
' Const mc0118 = 13 ' &HD%
' Const mc011A = 48 ' &H30%
  l0104 = p00F6.M269C
  l0106 = p00F6.M2684
  l00FC = gv7ABC(l0106).M1CB7
  p00F8 = p00F8 + 1
  If Len(l00FC) <> p00F8 Then
    If Len(l00FC) Then
      sub022F mc0108
    Else
      l00FC = String$(p00F8, 0)
    End If
  End If
  While p00F8 > 1
    If l0104 < 0 Then
      If gv7A2A And mc010C Then
        p00FA = ",?" + p00FA
      End If
      p00F8 = p00F8 - 1
    Else
      l0100 = gv110E(l0104).M2690
      If l0100 Then
        l0102 = Asc(Mid$(l00FC, p00F8, 1))
        If l0102 = 0 Then
          Mid$(l00FC, p00F8, 1) = Chr$(l0100)
        ElseIf (l0100 Xor l0102) And &HF Then
          If gv7ABC(l0106).M13E4 = -1 Then
            If (l0102 And mc0116) <> mc0118 Then
              l0102 = mc0118 Or (l0102 And mc011A)
              Mid$(l00FC, p00F8, 1) = Chr$(l0102)
            End If
          ElseIf (l0102 And &HF) > (l0100 And &HF) Then
            Mid$(l00FC, p00F8, 1) = Chr$(l0100)
          End If
        ElseIf (l0100 Or l0102) <> l0102 Then
          Mid$(l00FC, p00F8, 1) = Chr$(l0100 Or l0102)
        End If
        If gv7A2A And mc010C Then
          p00FA = "," + gv110E(l0104).M2696 + p00FA
        End If
        p00F8 = p00F8 - 1
      End If
      l0104 = gv110E(l0104).M269C
    End If
  Wend
  p00F6.M269C = l0104
  If l00FC <> gv7ABC(l0106).M1CB7 Then
    gv7ABC(l0106).M1CB7 = l00FC
    gv091A = True
  End If
End Sub

Sub sub0A86 (ByVal pv0148 As Integer, p014A As Integer)
Dim l014C As Integer
' Const mc014E = 2 ' &H2%
' Const mc0150 = 3 ' &H3%
  pv0148 = fn0E69(pv0148)
  Select Case gv0DE6(pv0148)
  Case mc014E, mc0150
    l014C = gv0DE6(pv0148 + mc00BE)
    If l014C Then
    If Len(gv7ABC(l014C).M1CB7) <> p014A + 1 Then
      If Len(gv7ABC(l014C).M1CB7) = 0 Then
        gv7ABC(l014C).M1CB7 = Chr$(gv0DE6(pv0148 + mc008C)) & String$(p014A, 0)
      Else
        sub01F8
      End If
    End If
    Else
      sub01F8
    End If
  Case Else
    sub01F8
  End Select
End Sub

Sub sub0A96 (p01C8 As Integer)
Dim l01CA As Integer
Dim l01CC As Integer
Dim l01CE As Integer
  l01CE = gv7ABC(gv7B2A).M2503
  If m0042 = gv09FA Then
    Exit Sub
  End If
    l01CC = m0042
    Do
      l01CA = gv09C4(l01CC)
      If l01CA <= p01C8 Then
        If l01CA < p01C8 Then Exit Do
        If Asc(Mid$(gv7952, l01CC - 1, 1)) = mc0078 Then Exit Sub
        sub09E5 l01CC, 1, mc0132, mc0078, 0
        Exit Sub
      End If
      l01CA = l01CC + 1
      l01CC = 0
      For l01CA = l01CA To gv09FA
        If gv09FC(l01CA) = l01CE Then l01CC = l01CA: Exit For
      Next
    Loop While l01CC
End Sub

Sub sub0C56 (p01D0 As Integer, p01D2 As Integer)
Dim l01D4 As Integer
Dim l01D6 As Integer
Dim l01D8 As Integer
Dim l01DA As Integer
' Const mc01DC = 16 ' &H10%
  l01D6 = Asc(Mid$(gv7952, p01D0, 1)) And mc013C
  If l01D6 > p01D2 Then Exit Sub
  l01DA = gv09FC(p01D0)
  If gv7B2A Then
    l01D8 = gv7ABC(gv7B2A).M2503
    If l01DA Then
      If l01DA <> l01D8 Then
        If p01D2 = mc013E Then
          l01D6 = Asc(Mid$(gv7952, p01D0, 1)) And Not mc013C
          If l01D6 And mc01DC Then
            Exit Sub
          End If
          p01D2 = mc0142
          Mid$(gv7952, p01D0, 1) = Chr$(l01D6 Or p01D2)
          gv091C = True
          Exit Sub
        ElseIf p01D2 <> mc0142 Then
          If gv0BC6 Then sub01F8
        End If
      End If
    End If
  End If
  l01D4 = gv09C4(p01D0)
  If p01D2 <= mc0146 Then
    If l01D4 = 0 Then
      p01D2 = mc013E
    ElseIf l01D4 < 0 Then
    ElseIf l01D4 And 1 Then
      sub01F8
    ElseIf l01D4 And &H8001 Then
      sub01F8
      If -l01D4 > gv7AF2.M13DD Then p01D2 = mc0132
    ElseIf l01D4 >= 6 And l01D4 < gv7AF2.M1519 \ &H200 Then
      p01D2 = mc007A
    Else
      p01D2 = mc0146
    End If
  End If
  If l01DA Or gv7B2A = 0 Then
    Exit Sub
  End If
  Select Case p01D2
  Case mc007A, mc0132
    If p01D0 < m0044 Then m0044 = p01D0
    GoSub L1CEF0
  Case mc0144
  Case mc013E
    If l01D4 Then
      p01D2 = mc0142
    Else
      GoSub L1CEF0
    End If
  Case mc0142
  Case mc0146
    GoSub L1CEF0
  Case mc0140
  Case Else
    sub01F8
  End Select
Exit Sub

L1CEF0:
  If p01D0 < m0044 Then Return
  If l01D8 > m0040 Then
  Else
    gv09FC(p01D0) = l01D8
    If p01D0 < m0042 Then
      gv7ABC(gv7B2A).M250C = p01D0
      m0042 = p01D0
    End If
    gv091C = True
  End If
Return
End Sub

Function fn0DCA (p00A4 As Integer) As String
Dim l00A6 As Integer
Dim l00A8 As Integer
Dim l00AA As Integer
Dim l00AC As Integer
' Const mc00BA = 15 ' &HF%
' Const mc00BE = 1 ' &H1%
' Const mc00C2 = 21 ' &H15%
' Const mc00CA = 14 ' &HE%
  l00AC = fn0D0A(p00A4)
  If l00AC > 0 Then fn0DCA = gv7956(l00AC): Exit Function
  l00A6 = fn033D(p00A4, 2)
  l00AA = gv09C4(p00A4)
  Select Case -l00AC
  Case mc00BA
    l00A8 = fn0E69(l00AA)
    l00A8 = gv0DE6(l00A8 + mc00BE)
    fn0DCA = fn0D70(l00A8)
  Case mc00C2
    l00A6 = gv09FC(p00A4)
    fn0DCA = fn0D70(m000E(l00A6))
  Case mc00CA
    fn0DCA = fn0CDA(fn032D(l00AA, 2))
  Case Else
    fn0DCA = gv793C(-l00AC) & fn02B5(l00A6)
  End Select
End Function

Function fn0E08 (p00DA As Integer) As Integer
On Error Resume Next
  fn0E08 = gv09FC(p00DA)
End Function

Function fn0E26 (p01C6 As Integer) As String
  fn0E26 = fn0DCA(p01C6)
End Function

Sub sub0607 (p0196 As Integer)
Dim l0198 As Integer
Dim l019A As Integer
Dim l019C As Integer
Dim l019E As Integer
Dim l01A0 As Integer
Dim l01A2 As Integer
Dim l01A4 As String
Dim l01A6 As Integer
Dim l01A8 As String
Dim l01AA As Integer
Dim l01AE As Integer
' Const mc01B0 = 12 ' &HC%
' Const mc01B2 = 32 ' &H20%
' Const mc01B4 = 2 ' &H2%
' Const mc01B6 = 6 ' &H6%
' Const mc01BA = 7 ' &H7%
' Const mc01BC = 8 ' &H8%
' Const mc01BE = 9 ' &H9%
' Const mc01C0 = 11 ' &HB%
  If gv7A2A Then Exit Sub
  sub05FF " ("
  l0198 = gv7ABC(p0196).M24E7
  If l0198 Then
    If gv7ABC(p0196).M24F0 Then
      sub05FF ("Index%")
      If Len(gv14DA(l0198).M2C0C) Then sub05FF ", "
    End If
    sub05FF gv14DA(l0198).M2C0C
  ElseIf gv7ABC(p0196).M13E4 = -1 Then
    l01A4 = gv7ABC(p0196).M1CB7
    l01A8 = l01A4
    GoSub L1C774
  Else
    l01A2 = gv7ABC(p0196).M2503
    For l0198 = 1 To gv09FA
      If gv09FC(l0198) = l01A2 Then Exit For
    Next
    If l0198 > gv09FA Then l0198 = 0
    l01A4 = gv7ABC(p0196).M1CB7
    l01A6 = gv7ABC(p0196).M250C
    l01A8 = gv7ABC(p0196).M2517
    If l0198 <> l01A6 Then
      If l01A6 = 0 Or l01A6 < l0198 Then
        l01A6 = l0198
        gv7ABC(p0196).M250C = l0198
        gv091A = True
      ElseIf l0198 = 0 Then
        l0198 = l01A6
      Else
        l01A6 = l0198
      End If
    End If
    l01A0 = gv7ABC(p0196).M24F9
    If l01A0 Then
      For l0198 = l0198 + 1 To gv09FA
        If gv09FC(l0198) = l01A2 Then Exit For
      Next
      If l0198 > gv09FA Then l0198 = 0
      l01A6 = l0198
    End If
    If Len(l01A4) > 1 And Len(l01A8) <> Len(l01A4) Then
      l01A8 = Chr$(0)
      If l01A6 Then
        Do While l0198
          l019C = Asc(Mid$(gv7952, l0198, 1))
          Select Case l019C And mc013C
          Case mc0132
            If (gv09C4(l0198) And &H8001) = 0 Then
              l01A8 = l01A8 + Chr$(l0198 - l01A6)
            Else
              Exit Do
            End If
          Case mc007A
            l01A8 = l01A8 + Chr$(l0198 - l01A6)
          End Select
          For l0198 = l0198 + 1 To gv09FA
            If gv09FC(l0198) = l01A2 Then Exit For
          Next
          If l0198 > gv09FA Then l0198 = 0
        Loop
      Else
      End If
    End If
    If Len(l01A8) <> Len(l01A4) Then l01A4 = ""
    l0198 = l01A6
    Do While l0198
      l019E = gv09FC(l0198)
      If l019E <> 0 And l019E <> l01A2 Then Exit Do
      l019C = Asc(Mid$(gv7952, l0198, 1))
      Select Case l019C And mc013C
      Case mc0132
        If (gv09C4(l0198) And &H8001) = 0 Then
          If l019A Then sub05FF ", "
          l019A = l019A + 1
          sub05FF "ByVal "
          sub07B0 l0198, gv09C4(), gv7952, False
        Else
          Exit Do
        End If
      Case mc007A
        If l019A Then sub05FF ", "
        l019A = l019A + 1
        If Len(l01A4) > l019A Then
          l01AE = Asc(Mid$(l01A4, l019A + 1, 1))
          If l01AE Then
            If (l01AE And mc0116) = mc01B0 Then
              If l01AE And mc01B2 Then
              Else
                sub05FF "ByVal "
              End If
            ElseIf l01AE And &H38 Then
            Else
              sub05FF "ByVal "
            End If
          End If
        End If
        sub07B0 l0198, gv09C4(), gv7952, False
      End Select
      For l0198 = l0198 + 1 To gv09FA
        If gv09FC(l0198) = l01A2 Then Exit For
      Next
      If l0198 > gv09FA Then
        l0198 = 0
      ElseIf gv09FC(l0198) <> l01A2 Then
        l0198 = 0
      End If
    Loop
    l01A4 = gv7ABC(p0196).M1CB7
  End If
  sub05FF ")"
  l0198 = gv7ABC(p0196).M24E0
  l0198 = gv0DE6(l0198 + mc008C)
  If l0198 Then sub07C9 l0198, 0, False
  sub06EA
  If Len(l01A4) > 1 And Len(l01A4) <> Len(l01A8) Then
    sub05FF "' called ("
    GoSub L1C774
    sub05FF ")"
    sub06EA
  End If
Exit Sub

L1C774:
  l019A = 0
  For l0198 = 2 To Len(l01A4)
    If l019A Then sub05FF ", "
    l019A = l019A + 1
    l01AE = Asc(Mid$(l01A4, l019A + 1, 1))
    l01AA = l01AE And mc0116
    If l01AE Then
      If l01AA = mc01B0 Then
        l01AE = l01AE And mc01B2
      ElseIf l01AE = mc0118 Then
        l01AE = l01AE And mc011A
      Else
        l01AE = l01AE And &H38
      End If
      If l01AE = 0 Then sub05FF "ByVal "
    End If
    sub05FF "p" & Format$(l019A)
    Select Case l01AA
    Case 0
    Case mc01B4 To mc01B6, mc01B0
      sub05FF Mid$(gc117C, l01AA, 1)
    Case mc01BA: sub05FF " As Variant"
    Case mc01BC: sub05FF " As object"
    Case mc01BE: sub05FF " As type"
    Case mc01C0: sub05FF " As array"
    Case Else
      sub05FF " As Any"
    End Select
  Next
Return
End Sub

Sub sub0616 ()
Dim l0186 As Integer
Dim l0188 As Integer
Dim l018A As Integer
Dim l018C As Integer
Dim l018E As Integer
  l018C = gv7ABC(gv7B2A).M2503
  m0042 = gv7ABC(gv7B2A).M250C
  If gv7A2A Then
    If m0044 = 0 Then
    End If
    If m0042 = 0 Then
      m0042 = m0044
    End If
    Exit Sub
  End If
  If gv7AF2.M13DD <= &H16 Then
    If gv7AF2.M13DD <> &H16 Then sub01F8: Exit Sub
  End If
  If m0042 = 0 Then
    For l0186 = 1 To gv09FA
      If gv09FC(l0186) = l018C Then m0042 = l0186: Exit For
    Next
    If m0042 = 0 Then Exit Sub
    gv7ABC(gv7B2A).M250C = m0042
  End If
  l0186 = m0042
  If gv09FC(l0186) > l018C Then Exit Sub
  l018E = gv7ABC(gv7B2A).M24F9
  If l018E Then
    If gv09C4(l0186) = l018E Or gv09C4(l0186) = 0 Then
      If gv7988(l0186) <> -mc00C2 Then
        gv7988(l0186) = -mc00C2
        gv091C = True
      End If
      For l0186 = l0186 + 1 To gv09FA
        If gv09FC(l0186) = l018C Then Exit For
      Next
      If l0186 > gv09FA Then l0186 = 0
    End If
  End If
  Do While l0186
    If gv09FC(l0186) > l018C Then Exit Do
    l0188 = Asc(Mid$(gv7952, l0186, 1))
    Select Case l0188 And mc013C
    Case mc013E
      sub05FF "Static "
      sub07B0 l0186, gv09C4(), gv7952, False
      sub06EA
    Case mc0132
      If gv09C4(l0186) <> l018E Then
        If gv09C4(l0186) And &H8001 Then
            sub05FF "Dim "
            sub07B0 l0186, gv09C4(), gv7952, False
            sub06EA
        End If
      End If
    Case mc0144 'A0
      l018A = Asc(Mid$(gv78DA, fn032D(gv09C4(l0186), 2), 1))
      If (l018A And mc013C) = mc0142 Then
        sub05FF "' "
        sub0795 fn032D(gv09C4(l0186), 2), gv78DA, True
        sub06EA
      End If
    Case mc0142  'C0
      sub05FF "' "
      sub0795 l0186, gv7952, False: sub06EA
    End Select
    If l0186 >= gv09FA Then
      l0186 = 0
    Else
      For l0186 = l0186 + 1 To gv09FA
        If gv09FC(l0186) = l018C Then Exit For
      Next
      If l0186 > gv09FA Then Exit Do
    End If
  Loop
End Sub

Sub sub0DD6 (p0054 As Integer, p0056 As Integer)
Dim l0058 As Integer
Dim l005A As String
Dim l005C As String
Dim l005E As Integer
Dim l0060 As Integer
Dim l0062 As Integer
Dim l0064   As Integer
' Const mc0076 = 8 ' &H8%
' Const mc0078 = 9 ' &H9%
' Const mc007A = 64 ' &H40%
  l0058 = gv7ABC(p0056).M24E7
  l005A = gv14DA(l0058).M2C0C
  l005C = gv14DA(l0058).M1CB7
  If gv7ABC(p0056).M24F0 Then
    If Len(l005A) Then
      l005A = "Index%, " & l005A
    Else
      l005A = "Index%"
    End If
    l005C = Chr$(1) & l005C
  End If
  l0060 = gv09C4(p0054) - 6
  If l0060 Mod 4 Then sub01F8
  l0060 = l0060 \ 4
  If l0060 >= Len(l005C) Then
    sub01F8
    Exit Sub
  End If
  l0060 = Len(l005C) - 1 - l0060
  l0062 = Asc(Mid$(l005C, l0060 + 1, 1))
  l005E = 0
  While l0060 > 0
    l005E = InStr(l005E + 1, l005A, ",")
    If l005E = 0 Then sub01F8
    l0060 = l0060 - 1
  Wend
  l005A = Trim$(Mid$(l005A, l005E + 1))
  l005E = InStr(l005A, ",")
  If l005E Then l005A = Left$(l005A, l005E - 1)
  l005E = InStr(l005A, " ")
  If l005E Then
    l005A = RTrim$(Left$(l005A, l005E))
  Else
    l005E = InStr(gc1576, Right$(l005A, 1))
    If l005E Then l005A = Left$(l005A, Len(l005A) - 1)
  End If
  Select Case l0062
  Case 1 To 5
  Case 6: l0062 = 7
  Case 7: l0062 = mc0076
  Case 8: l0062 = mc0078
  Case Else: sub01F8
  End Select
  sub0A05 p0054, mc007A, l0062, l0064
  sub09F3 l005A, p0054
End Sub

Sub sub0DE9 ()
Dim l0152 As Integer
Dim l0156 As Integer
' Const mc015C = 229 ' &HE5%
' Const m0160 = 0 ' &H0%
' Const mc0162 = 225 ' &HE1%
' Const mc0164 = 226 ' &HE2%
' Const mc0166 = 233 ' &HE9%
' Const mc016E = 231 ' &HE7%
  gv7874 = True
  sub05FF "' main.txt - global definitions": sub06EA
  'Mid$(gv78DA, 745, 1) = Chr$(&HE1)
  l0152 = 3
  If Asc(Right$(gv78DA, 1)) <> mc015C Then
    sub01F8
    Mid$(gv78DA, gv09C2 + 1, 1) = Chr$(mc015C)
  End If
  Do
    l0156 = Asc(Mid$(gv78DA, l0152, 1))
    Select Case l0156 And mc013C  ' E0
    Case m0160, mc0140 'E0, 0
      Select Case l0156
      Case 0
        While Asc(Mid$(gv78DA, l0152, 1)) = 0
          l0152 = l0152 + 1
        Wend
      Case mc0162
        sub0D56 l0152
        l0152 = InStr(l0152, gv78DA, Chr$(mc0164))
        If l0152 = 0 Then
          sub022F mc012E
          Exit Do
        Else
          l0152 = l0152 + 1
        End If
      Case mc015C: Exit Do
      Case 8 To &H1F
        l0152 = l0152 + 1
      Case mc009E: sub01F8: l0152 = l0152 + 1
      Case mc0166: l0152 = l0152 + 1
      Case mc0140
        If gv78A8(l0152) Then sub01F8
        l0152 = l0152 + 1
      Case mc016E
        sub0795 l0152, gv78DA, True
        sub06EA
        l0152 = l0152 + 1
      Case Else
        l0152 = l0152 + 1
        sub022F mc012E
      End Select
    Case mc0146
      sub022F mc012E
      sub05FF "Global " & fn0CDA(l0152) & " ' As ?"
      sub06EA
      l0152 = l0152 + 1
    Case mc0142
      sub0795 l0152, gv78DA, True
      sub06EA
      l0152 = l0152 + 1
    Case mc0144
      sub07B0 l0152, gv098C(), gv78DA, True
      sub06EA
      l0152 = l0152 + 1
    Case Else
      sub01F8
      l0152 = l0152 + 1
    End Select
  Loop
End Sub

Private Sub sub0DFB (p007C As Integer)
Dim l007E  As Integer
Dim l0080 As Integer
Dim l0082 As Integer
Dim l0084 As Integer
' Const mc0086 =-2 ' &HFFFE%
' Const mc008C = 5 ' &H5%
' Const mc009E = 227 ' &HE3%
  gv2382 = mc0086
  ReDim m000E%(gv7AEE)
  m0040 = 0
  For l007E = 1 To gv7AEE
    If gv7ABC(l007E).M24AC = p007C Then
      l0080 = gv7ABC(l007E).M2503
      l0082 = gv7ABC(l007E).M24E0
      If l0080 Then
        m000E(l0080) = l007E
        m0040 = m0040 + 1
      End If
      If gv0DE6(l0082 + mc008C) Then l0084 = l0084 + 1
    End If
  Next
  If l0084 Then
    If gv7A70(gv7AB6).M12C5 Then l0082 = &H1A \ 2 Else l0082 = 6 \ 2
    Mid$(gv7952, l0082, l0084) = String$(l0084, mc009E)
  End If
End Sub

Private Sub sub0E17 (ByVal pv00DC As Integer)
Dim l00DE As Integer
Dim l00E0 As Integer
Dim l00E8 As T14DB
Dim l00EA As Integer
  If pv00DC Then
    ReDim gv09FC(gv09FA)
    Exit Sub
  End If
  If gv7AEE = 0 Then sub01F8: Exit Sub
  For pv00DC = 1 To gv7AA2
    l00DE = 0
    l00EA = gv7A70(pv00DC).M24AC
    For l00E0 = 1 To gv7AEE
      If gv7ABC(l00E0).M24AC = l00EA Then
        Get gv088E, fn0E71(gv7ABC(l00E0).M24E0), l00E8
        If l00E8.M15AE And &H1000000 Then
          l00DE = l00DE + 1
          gv7ABC(l00E0).M2503 = l00DE
        ElseIf gv7ABC(l00E0).M1C2B Then
          l00DE = l00DE + 1
          gv7ABC(l00E0).M2503 = l00DE
        End If
      End If
    Next
  Next
End Sub

