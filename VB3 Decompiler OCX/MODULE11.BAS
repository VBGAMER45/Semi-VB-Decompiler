Attribute VB_Name = "Module11"
' Module11
Option Explicit
Const mc00E0 = 64 ' &H40%
Const mc010C = 4130815  ' &H3F07FF&
Const mc0110 = 4130752  ' &H3F07C0&
Const mc0114 = 4130754  ' &H3F07C2&
Const mc0118 = 1026 ' &H402%
Const mc0192 = 48 ' &H30%
Const mc01CC = 256 ' &H100%
Const mc01CE = 5 ' &H5%
Const mc01D0 = 32 ' &H20%
Const mc01D2 = 4 ' &H4%
Const mc01D4 = 8 ' &H8%
Const mc01D8 = 64 ' &H40%
Const mc01EA = 2 ' &H2%
Const mc01EC = 3 ' &H3%
Const mc01EE = 32 ' &H20%
Const mc01F0 = 5 ' &H5%
Const mc01F2 = 1 ' &H1%
Const mc01F8 = 2048 ' &H800%
Const mc01FA = 2 ' &H2%
Const mc01FC = 7 ' &H7%
Const mc01FE = 9 ' &H9%
Const mc0200 = 10 ' &HA%
Const mc0202 = 224 ' &HE0%
Const mc0204 = 32 ' &H20%
Const mc0206 = 64 ' &H40%
Const mc0208 = 128 ' &H80%
Const mc020A = 192 ' &HC0%
Const mc020C = 16 ' &H10%
Const mc020E = 2 ' &H2%
Const mc0216 = 9 ' &H9%
Const mc0218 = 1 ' &H1%
Const mc021C = 512 ' &H200%
Const mc021E = 12 ' &HC%
Const mc0220 = 9 ' &H9%
Const mc0222 = 10 ' &HA%
Const mc0224 = 8192 ' &H2000%
Const mc023C = -2 ' &HFFFE%
Const mc023E = 96 ' &H60%
Const mc0240 = 224 ' &HE0%
Const mc024C = 10 ' &HA%
Const mc024E = 16 ' &H10%
Const mc0256 = 4 ' &H4%
Const mc0282 = 224 ' &HE0%
Const mc0284 = 231 ' &HE7%

Function fn0A66() As Integer
Dim l01A4 As Integer
Dim l01A6 As Integer
Dim l01AA As Integer
Dim l01AC As Integer
Dim l01AE As Integer
Dim l01B0 As Integer
Dim l01B2 As Integer
Dim l01b4 As Integer
Dim l01B6 As String
Dim l01B8 As String
Dim l01BA As Integer
Dim l01BE As T265D
Dim l01C2 As Integer
Dim l01C4 As Long
Dim l01C6 As Long
Dim l01C8 As Long
' Const mc01CC = 256 ' &H100%
' Const mc01CE = 5 ' &H5%
' Const mc01D0 = 32 ' &H20%
' Const mc01D2 = 4 ' &H4%
' Const mc01D4 = 8 ' &H8%
' Const mc01D8 = 64 ' &H40%
' Const mc01EA = 2 ' &H2%
' Const mc01EC = 3 ' &H3%
' Const mc01EE = 32 ' &H20%
' Const mc01F0 = 5 ' &H5%
' Const mc01F2 = 1 ' &H1%
' Const mc01F8 = 2048 ' &H800%
' Const mc01FA = 2 ' &H2%
' Const mc01FC = 7 ' &H7%
' Const mc01FE = 9 ' &H9%
' Const mc0200 = 10 ' &HA%
' Const mc0202 = 224 ' &HE0%
' Const mc0204 = 32 ' &H20%
' Const mc0206 = 64 ' &H40%
' Const mc0208 = 128 ' &H80%
' Const mc020A = 192 ' &HC0%
' Const mc020C = 16 ' &H10%
' Const mc020E = 2 ' &H2%
' Const mc0216 = 9 ' &H9%
' Const mc0218 = 1 ' &H1%
' Const mc021C = 512 ' &H200%
' Const mc021E = 12 ' &HC%
' Const mc0220 = 9 ' &H9%
' Const mc0222 = 10 ' &HA%
' Const mc0224 = 8192 ' &H2000%
  sub09C7       'load analyze operand
  l01BE.M1E3E = gv0924
  l01BE.M266C = gv2382
  l01BE.M269C = gv1140
  l01A4 = gv237A.M2EF9
  gv091E = gv15E0.M2CC3 * 2
  If gv091E < 0 Then
    Get gv088E, , gv091E: gv0920 = gv0924 + 4 + gv091E
  Else
    gv0920 = gv0924 + 2 + gv091E
  End If
  If l01A4 = 0 Then
    sub022F mc01CC
    fn0A66 = mc0192
    gv1140 = gv1140 + 1
    gv110E(gv1140) = l01BE
    Exit Function
  End If
  Select Case gv15E0.M2CAB And &HF
  Case mc01CE
    fn0A66 = mc01D0
    gv1140 = -1
    l01BE.M269C = -1
    ReDim gv110E(64)
  Case mc01D2
    fn0A66 = mc0192
  Case mc01D4
    l01BE.M2690 = gc11A2
    If gv7A2A > 0 Then
      If gv7A2A And gc1146 Then
        l01BE.M2696 = "$"
        Get gv088E, gv0924 + 2, l01A6
        Get gv088E, , l01A4
        If l01A6 > 2 Then
          Get gv088E, , l01A4
          If l01A4 Then
            l01BE.M268A = Space$(l01A4): Get gv088E, , l01BE.M268A
          End If
        End If
        l01BE.M268A = """" + l01BE.M268A + """"
      End If
    End If
  Case Else
    If gv7A2A > 0 Then
      l01AA = Asc(Mid$(gv1142, gv2382 + 1, 1))
      l01b4 = Asc(Mid$(gv1142, gv2382 + 2, 1))
      If l01b4 And gc11D6 Then
        l01b4 = l01b4 Xor gc11D6
      End If
      If l01AA Then
        l01AA = gv2382 - l01AA
        l01AC = Asc(Mid$(gv1142, l01AA, 1))
        l01B2 = Asc(Mid$(gv1142, l01AA + 1, 1))
        If l01AC > gc114E Then
          Get gv088E, gv0920 - 2, l01A6: l01A6 = fn032D(l01A6, 2)
          If gv15E0.M2CB3 And mc01D8 Then
            l01BE.M2672 = l01A6
          End If
          Select Case l01b4
          Case 0
          Case gc1172
            Select Case l01AC
            Case gc1156
              l01A6 = fn033D(l01A6, 2)
              l01A6 = fn0E69(l01A6)
            Case gc1152
              l01A6 = fn0E69(gv09C4(l01A6))
            Case Else
              sub022F mc01CC
            End Select
            l01B2 = gv0DE6(l01A6)
            If l01B2 <> mc01EA And l01B2 <> mc01EC Then
              sub022F mc01EE
            End If
            l01B2 = gv0DE6(l01A6 + mc01F0)
            l01B2 = gv11DC(l01B2)
            l01A6 = gv0DE6(l01A6 + mc01F2)
            l01AC = gc115A
            l01BE.M2684 = l01A6
            If l01B2 = gc1192 Then
              l01B0 = gc1194
            End If
          Case gc1168 '1
            l01B2 = Asc(Mid$(gv7952, l01A6, 1))
            Select Case l01AC
            Case gc1154, gc1152, gc1150 '4,3,2,5
            Case gc1156
              l01C2 = Asc(Mid$(gv78DA, fn032D(gv09C4(l01A6), 2), 1))
              If l01C2 < l01B2 Then
                Mid$(gv78DA, fn032D(gv09C4(l01A6), 2), 1) = Chr$(l01B2)
              ElseIf l01C2 > l01B2 Then
                l01B2 = l01C2
                Mid$(gv7952, l01A6, 1) = Chr$(l01B2)
              End If
            Case Else
              sub022F mc01EE
              l01B2 = 0
            End Select
            If (l01B2 And &HF) Then
              l01B2 = gv11DC(l01B2 And &HF)
            Else
              l01B2 = gc119C
            End If
            If l01B2 = gc11A2 Then l01AE = gc11C8
          Case gc1170
            If l01B2 <> gc119A Then sub022F mc01F8
            If gv09C4(l01A6) < 0 Then
              l01AC = gc1154
            Else
              l01AC = gc1150
            End If
          Case gc1178
            Get gv088E, gv0920, l01A4
            If iVBVersion = 3 And l01A4 = &H6A63 Then
              l01B2 = gc11A2
            ElseIf iVBVersion = 2 And l01A4 = &H5D63 Then
              l01B2 = gc11A2
            Else
              l01A4 = Asc(Mid$(gv7952, l01A6, 1))
              Select Case l01A4 And &H1F
              Case 0
                  l01B2 = gc11A6
              Case mc01FA
                l01B2 = gc1190
              Case mc01FC
                l01B2 = gc11A2
              Case mc01FE
                l01B2 = gv11DC(l01A4 And &H1F)
              Case mc0200
                sub01F8
                l01B2 = gv11DC(l01A4 And &H1F)
              Case Else
                l01B0 = gc11A6
              End Select
            End If
          Case gc1176
            l01A4 = Asc(Mid$(gv7952, l01A6, 1))
            Select Case l01A4 And mc0202
            Case Is <= mc0204
              If gv09C4(l01A6) < 0 Then
                l01AC = gc1154
              End If
            Case mc0206
              l01AC = gc1150
            Case mc0208
              l01AC = gc1154
            Case mc020A
              l01AC = gc1158
            Case Else
              sub022F mc01EE
              l01A4 = 0
            End Select
            If l01A4 And &HF Then
              l01B2 = gv11DC(l01A4 And &H1F)
              If l01B2 = gc118C Then
                l01B2 = gc11A4
              ElseIf l01B2 And 8 Then
                l01AE = gc11C8
              ElseIf l01AC = gc1150 Then
                l01AE = gc11CA
              End If
            Else
              l01B2 = gc11A4
            End If
          Case gc116E
            l01AC = gv09C4(l01A6 + 2)
            l01B2 = l01AC And &HFF
            Select Case l01AC And &HF00
            Case 0
              l01AC = gc1156
            Case &H100, &H200
              l01AC = gc1152
            Case &H500
              l01AC = gc1150
            Case &H600
              l01AC = gc1154
            Case Else
              sub022F mc01F8
            End Select
            Get gv088E, gv0924 + 2, l01A4
            If l01A4 < 0 Then
              l01B2 = gc11A0
            Else
              If l01B2 = 0 Then
                l01B2 = gc119C
              Else
                l01B2 = gv11DC(l01B2)
                If l01B2 And 8 Then
                  l01AE = gc11C8
                Else
                  l01AE = gc11CA
                End If
              End If
              l01B2 = l01B2 Or mc020C
            End If
          Case gc116A, gc116C
            If l01AC <> gc115E Then sub022F mc01F8
            l01A4 = gv098C(l01A6 + mc020E)
            Select Case l01A4
            Case Is < 0, Is >= mc0216
              l01A4 = mc0200
            Case mc01FC
              l01AE = gc11C8
            Case Is < mc01FC
            End Select
            l01B2 = gv11DC(l01A4)
            If gv098C(l01A6 + mc0218) And 1 Then
              l01B2 = l01B2 Or mc020C
              If l01b4 <> gc116C Then sub022F mc01F8
            ElseIf l01b4 = gc116C Then
              sub022F mc01F8
            End If
          End Select
          l01BE.M2678 = l01AC
          l01BE.M267E = l01B2
          If gv7A2A And gc1146 Then
            l01BE.M268A = Mid$(gc114A, l01AC, 1)
            If l01B2 = gc11A0 Then
              l01BE.M268A = l01BE.M268A + "a" + fn02B5(l01BE.M2672 * 2)
            Else
              If l01B2 And mc020C Then
                l01BE.M268A = l01BE.M268A + "a"
                l01B2 = l01B2 And &HF
              ElseIf l01AC <> gc115A Then
                l01BE.M268A = l01BE.M268A + "v"
              End If
              l01BE.M268A = l01BE.M268A + fn02B5(fn0335(l01BE.M2672) * 2)
              If l01B2 > 1 Then
                l01BE.M268A = l01BE.M268A + Mid$(gc117C, l01B2, 1)
              End If
            End If
          Else
            l01B2 = l01B2 And &HF
          End If
          If l01B0 = 0 Then l01B0 = l01B2 And &H3F
          If l01B0 >= gc11A8 Then sub022F mc01F8
          l01B0 = l01B0 Or l01AE
          l01A4 = Asc(Mid$(gv1142, l01AA + 2, 1))
          l01B6 = Mid$(gv1142, l01AA + 3, l01A4)
        Else
          l01B0 = gc118C
          l01A4 = 0
          Select Case l01b4
          Case 0
            l01A4 = Asc(Mid$(gv1142, l01AA + 2, 1))
            l01B6 = Mid$(gv1142, l01AA + 3, l01A4)
          Case gc1174
            Get gv088E, gv0920 - 4, l01AA
            l01AA = gv1148 + (l01AA And &HFF) * &H18
            l01AA = l01AA + Asc(Mid$(gv1142, l01AA, 1)) + 1
            Get gv088E, , l01A4
            l01A4 = l01A4 + 1
            l01B6 = Mid$(gv1142, l01AA, l01A4)
          Case gc117A
            l01A4 = Asc(Mid$(gv1142, l01AA + 2, 1))
            l01B6 = Mid$(gv1142, l01AA + 3, l01A4)
          Case Else
            sub022F mc01CC
          End Select
        End If
        If l01A4 Then
          l01BE.M2690 = l01B0 Or l01AE
          l01B0 = l01B0 And gc11D2
          For l01AA = Len(l01B6) To 2 Step -1
            l01A4 = Asc(Mid$(l01B6, l01AA, 1))
            If l01A4 And gc11C4 Then
              Select Case l01A4 And &H1F
              Case gc11A8
                Get gv088E, gv0924 + 2, l01A4
                If l01BE.M2684 Then
                  sub0A78 l01BE, l01A4, l01B8
                Else
                  If l01A4 > 0 Then GoSub L1311A
                End If
              Case gc11AA
                l01AE = 0
                l01A4 = gc1190: l01B2 = l01A4
                Get gv088E, gv0924 + 2, l01A6
                If l01A6 >= 0 Then GoSub L13288
                GoSub L13288
              Case Else
                sub022F mc021C
              End Select
            ElseIf l01A4 Then
              If (l01A4 And &HF) = gc118C Then
                l01AE = l01A4 And gc11CA
                l01B2 = l01B0 Or l01AE
                If l01B2 > &H17 Then
                  sub022F mc01F8
                  l01B2 = l01B2 Xor gc11CA
                End If
              Else
                l01B2 = l01A4 And gc11D2
              End If
              GoSub L13288
            End If
          Next
          l01A4 = Asc(Left$(l01B6, 1))
          If l01A4 And gc11C4 Then
            sub022F mc01F8
          ElseIf l01A4 Then
            l01B2 = l01A4 And gc11D0
            Select Case l01B2
            Case gc11A4, gc11A6
              If l01B0 > gc118C Then
                l01A4 = (l01A4 Xor l01B2) Or l01B0
              End If
            Case Is >= gc11A8
              sub022F mc01F8
            End Select
          End If
          GoSub L1362E
          If gv7A2A And gc1146 Then
            l01BE.M268A = l01BE.M268A + "(" + Mid$(l01B8, 2) + ")"
          End If
        End If
      ElseIf l01b4 Then
        sub022F mc01EE
      End If
    ElseIf gv7A2A = True Then
      If gv15E0.M2CB3 And mc01D8 Then
        Get gv088E, gv0920 - 2, l01A6
        If gv237A.M2F09 = mc021E Then
          Get gv088E, gv0924 + 2, l01A4
          If gv237A.M2F00 = mc0220 Then
            sub07EB l01A6, gv237A.M2F00, gv237A.M2F09
            l01A6 = gv09C4(fn032D(l01A6, 2))
          ElseIf gv237A.M2F00 <> mc0222 Then
          End If
          sub0A86 l01A6, l01A4
        Else
          sub07EB l01A6, gv237A.M2F00, gv237A.M2F09
        End If
      ElseIf iVBVersion = 3 Then
        Select Case gv2382
        Case &H4FA6
          Get gv088E, gv0920 - 2, l01A4
          sub0A96 l01A4
        Case &H6F79
        End Select
      ElseIf iVBVersion = 2 Then
        Select Case gv2382
        Case &H4D89
          Get gv088E, gv0920 - 2, l01A4
          sub0A96 l01A4
        Case &H6231
        End Select
      End If
    End If
  End Select
  If gv7A2A > 0 Then
    gv1140 = gv1140 + 1
    If gv1140 > UBound(gv110E) Then ReDim Preserve gv110E(gv1140 + 10)
    gv110E(gv1140) = l01BE
  End If
Exit Function

L1311A:
  If l01A4 <= 0 Then
    If l01A4 = &H8000 Then
      If l01BE.M2690 <> gc11A0 Then
        l01BE.M2690 = gc11A0
      End If
    End If
    Return
  End If
  l01AC = l01BE.M269C
  While l01A4 > 0
    If l01AC < 0 Then
      If gv7A2A And gc1146 Then
        l01B8 = ",?" + l01B8
      End If
      l01A4 = l01A4 - 1
    Else
      l01BA = gv110E(l01AC).M2690
      If l01BA Then
        l01B6 = l01B6 + Chr$(l01BA)
        If gv7A2A And gc1146 Then
          l01B8 = "," + gv110E(l01AC).M2696 + l01B8
        End If
        l01A4 = l01A4 - 1
      End If
      l01AC = gv110E(l01AC).M269C
    End If
  Wend
  l01BE.M269C = l01AC
Return

L13288:
  If l01B2 And gc11C6 Then sub022F mc01F8: l01B2 = l01B2 Xor gc11C6
  l01AC = l01BE.M269C
  Do While l01AC > 0
    l01C2 = gv110E(l01AC).M2690 And gc11D2
    If l01C2 Then
      If l01C2 <> l01B2 Then
        l01C4 = gv11AC(l01B2)
        l01C6 = gv11AC(l01C2)
        l01C8 = l01C4 And l01C6
        If l01C4 = 0 Or l01C6 = 0 Then
          sub022F mc01F8
        ElseIf l01C8 = 0 Then
        ElseIf l01C8 = l01C4 Then
          If gv110E(l01AC).M2684 Then
            sub022F mc01F8
          ElseIf gv110E(l01AC).M2672 > 0 Then
            sub0A05 gv110E(l01AC).M2672, (gv110E(l01AC).M2678 * &H20) And &HFF, (gv11F2(l01B2 And gc11D0)), 0
          Else
            sub022F mc01F8
          End If
          l01C2 = gv110E(l01AC).M2690 Xor l01C2
          gv110E(l01AC).M2690 = l01B2 Or l01C2
        ElseIf l01C8 = l01C6 Then
          l01B2 = l01C2
          If (l01A4 And &HF) = gc118C Then
            l01B0 = l01B2
            If l01BE.M2672 Then
              If l01BE.M2684 Then
                sub022F mc0224
              Else
                sub022F mc0224
              End If
            End If
          End If
        Else
          sub022F mc01F8
        End If
      End If
      If gv7A2A And gc1146 Then
        l01B8 = "," + gv110E(l01AC).M2696 + l01B8
      End If
      l01BE.M269C = gv110E(l01AC).M269C
      Return
    End If
    l01AC = gv110E(l01AC).M269C
  Loop
  sub022F mc01F8
  If gv7A2A And gc1146 Then
    l01B8 = l01B8 + ",?"
  End If
  l01BE.M269C = 0
Return

L1362E:
  If (l01A4 And gc11D0) = gc118C Then
    l01AE = l01A4 And Not gc11D0
    l01A4 = l01B0 Or l01AE
  End If
  If l01A4 And 8 Then l01A4 = l01A4 And Not gc11CA
  If l01A4 = gc118C Then sub022F mc01F8
  l01BE.M2690 = l01A4
  If l01A4 = 0 Then Return
  If l01BE.M2672 And l01BE.M2684 = 0 Then
    If (l01BE.M267E And gc11D0) > gc11A2 Then
      If l01BE.M267E <> (l01A4 And gc11D0) Then
        sub0A05 l01BE.M2672, l01BE.M2678 * &H20, (gv11F2(l01A4 And &HF)), 0
      End If
    End If
  End If
  If gv7A2A And gc1146 Then
    If l01A4 And &HF Then
      l01BE.M2696 = Mid$(gc1184, l01A4 And &HF, 1)
    Else
      l01BE.M2696 = "?"
    End If
    If l01A4 And gc11C8 Then
      l01BE.M2696 = "R" + l01BE.M2696
    ElseIf l01A4 And gc11CA Then
      l01BE.M2696 = "^" + l01BE.M2696
    End If
    If Len(l01BE.M268A) = 0 Then l01BE.M268A = l01BE.M2696
  End If
Return
End Function

Sub sub0210()
Dim l0226 As Integer
Dim l0228 As Integer
Dim l022A  As Integer
Dim l022C  As Integer
Dim l022E As Integer
Dim l0232 As T20E0
Dim l0234  As Integer
' Const mc023C =-2 ' &HFFFE%
' Const mc023E = 96 ' &H60%
' Const mc0240 = 224 ' &HE0%
' Const mc024C = 10 ' &HA%
' Const mc024E = 16 ' &H10%
' Const mc0256 = 4 ' &H4%
  sub0201 l0234, gc03D6 & gc03FE & gc03F6 'local fixups
  gv2382 = mc023C
  gv7B2A = 0
  Do
    Get gv088E, , l0226
    Get gv088E, , l022A
    Get gv088E, , l022C
    If l022C <> &H1E Then Exit Do
    If l022A = 0 Then
      l022A = (l0226 - 4) \ 3
      While l022A > 0
        Get gv088E, , l0232
        l0226 = fn032D(l0232.M1EF7, 2)
        Select Case Asc(l0232.M12B8)
        Case 7
          sub09E5 l0226, 2, mc023E, mc01FC, 0
        Case 9
          Select Case gv09C4(l0226) And &HF000
          Case &H8000
            If (gv09C4(l0226 - 1) And &HFF00) <> &H4000 Then sub01F8
            sub09E5 l0226, 2, mc0240, mc0216, 0
            sub09F3 fn0C4A(gv7AB6, gv09C4(l0226) And &HFF), l0226
          Case &HC000
            sub09E5 l0226, 2, mc0240, mc0216, 0
            sub064E gv7AA6.M12C5
            sub09F3 fn0B6F(gv09C4(l0226)), l0226
          Case 0
            If gv09C4(l0226 - 1) > gv0FD2 Then
              sub01F8
            Else
              sub0A05 l0226, mc023E, mc0216, 0
            End If
          Case Else
            sub01F8
          End Select
        Case &H20
          If Not fn0CFD(gv09C4(l0226 - 1)) Then sub01F8
          sub0A05 l0226, mc023E, mc024C, 0
        Case &H80
          sub0A10 l0226
        Case &HC0
          If gv09C4(l0226) Then sub01F8
          sub0A10 l0226
        Case Else
          sub022F mc024E
        End Select
        l022A = l022A - 1
      Wend
    Else
      If l0226 \ 2 <> l022A + 2 Then sub022F mc01EE
      While l022A
        Get gv088E, , l0228
        l0226 = fn032D(l0228, 2)
        l022E = Asc(Mid$(gv7952, l0226, 1))
        If l022E = 0 Then
          If l0228 And 1 Then
            If gv09C4(l0226 + 1) <> gv0B70 Then sub01F8
            sub09E5 l0226, 2, mc020A Or 7, 0, 0
          Else
            sub0A10 l0226
          End If
        End If
        l022A = l022A - 1
      Wend
    End If
  Loop
  sub047C l0234, gc03D6 & gc03F2 'scaning tokens
  sub0A21 gv7AB6
  For l0226 = 1 To gv7AEE
    If gv7ABC(l0226).M24AC = gv7AB8 Then
      If (gv7A1C And mc0256) = 0 Then
        frmMain.control9.Text = fn0D70(l0226)
      End If
      DoEvents
      If gv7ABC(l0226).M13E4 <> -1 Then
        sub01C8 l0226
      End If
    End If
  Next l0226
  sub047C l0234, gc03DA & gc03FE & gc03EA   'local scopes
  sub0A33
  sub0220 l0234
End Sub

Sub sub03FC()
  Erase gv0FBC
  ReDim gv0FF6(0)
  gv1028 = 0
  gv0FBC(1).M1D68 = "Form"
    gv0FBC(1).M25C6 = &HD + 1
    gv0FBC(1).M25DA = 5
  gv0FBC(2).M1D68 = "MDIForm"
    gv0FBC(2).M25C6 = &H14 + 1
  gv0FBC(4).M1D68 = "Control"
    gv0FBC(4).M25C6 = 0
  gv0FBC(5).M1D68 = "Controls"
    gv0FBC(5).M25DA = -4
  gv0FBC(8).M1D68 = "Forms"
    gv0FBC(8).M25DA = -1
    gv0FBC(&HB).M1D68 = "Database"
      gv0FBC(&HB).M25C6 = &H20 + 1
      gv0FBC(&HB).M25DA = &HC
    gv0FBC(&HC).M1D68 = "TableDefs"
      gv0FBC(&HC).M25C6 = &H1D + 1
      gv0FBC(&HC).M25DA = -&HD
    gv0FBC(&HD).M1D68 = "TableDef"
      gv0FBC(&HD).M25C6 = &H1A + 1
      gv0FBC(&HD).M25DA = &HF
    gv0FBC(&HE).M1D68 = "Dynaset"
      gv0FBC(&HE).M25C6 = &H1F + 1
      gv0FBC(&HE).M25DA = &HF
    gv0FBC(&HF).M1D68 = "Fields"
      gv0FBC(&HF).M25C6 = &H1C + 1
      gv0FBC(&HF).M25DA = -&H11
    gv0FBC(&H10).M1D68 = "Indexes"
      gv0FBC(&H10).M25C6 = &H1E + 1
      gv0FBC(&H10).M25DA = -&H12
    gv0FBC(&H11).M1D68 = "Field"
      gv0FBC(&H11).M25C6 = &H19 + 1
    gv0FBC(&H12).M1D68 = "Index"
      gv0FBC(&H12).M25C6 = &H1B + 1
  If iVBVersion = 3 Then
    gv0FBC(&H14).M1D68 = "Object"
      gv0FBC(&H14).M25C6 = &H24 + 1
    gv0FBC(&H15).M1D68 = "QueryDef"
      gv0FBC(&H15).M25C6 = &H21 + 1
    gv0FBC(&H16).M1D68 = "Table"
      gv0FBC(&H16).M25C6 = &H22 + 1
      gv0FBC(&H16).M25DA = &HF
    gv0FBC(&H17).M1D68 = "Snapshot"
      gv0FBC(&H17).M25C6 = &H23 + 1
      gv0FBC(&H17).M25DA = &HF
    gv0FBC(&H32).M1D68 = "Printer"
      gv0FBC(&H32).M25C6 = &HC + 1
    gv0FBC(&H33).M1D68 = "Screen"
      gv0FBC(&H33).M25C6 = &HE + 1
    gv0FBC(&H34).M1D68 = "Clipboard"
      gv0FBC(&H34).M25C6 = &HF + 1
    gv0FBC(&H3D).M1D68 = "App"
      gv0FBC(&H3D).M25C6 = &H15 + 1
  ElseIf iVBVersion = 2 Then
    gv0FBC(&H2C).M1D68 = "Printer"
      gv0FBC(&H2C).M25C6 = &HC + 1
    gv0FBC(&H2D).M1D68 = "Screen"
      gv0FBC(&H2D).M25C6 = &HE + 1
    gv0FBC(&H2E).M1D68 = "Clipboard"
      gv0FBC(&H2E).M25C6 = &HF + 1
    gv0FBC(&H37).M1D68 = "App"
      gv0FBC(&H37).M25C6 = &H15 + 1
  Else
    sub01F8
  End If
End Sub

Sub sub04D2()
Dim l0274 As Integer
Dim l0276 As Integer
Dim l0278 As Integer
Dim l027A As Integer
Dim l027C As Integer
' Const mc0282 = 224 ' &HE0%
' Const mc0284 = 231 ' &HE7%
  Get gv088E, gv0B96 + gv09BE, l0274
  Get gv088E, , l0278
  Get gv088E, , l027A
  If l027A <> &H1E Then sub01F8
  If l0274 \ 2 <> l0278 + 2 Then sub01F8
  While l0278 > 0
    Get gv088E, , l0276
    l0274 = fn032D(l0276, 2)
    l027C = Asc(Mid$(gv78DA, l0274, 1))
    If l027C = 0 Then
      If l0276 And 1 Then
        If gv098C(l0274 + 1) <> gv0B70 Then sub01F8
        sub04B5 l0274, 2, mc0282, mc01FC, 0
        If Asc(Mid$(gv78DA, l0274, 1)) <> mc0284 Then sub01F8
      Else
        sub09D4 l0274
      End If
    End If
    l0278 = l0278 - 1
  Wend
End Sub

Sub sub06D9(p026C As Integer)
  gv2382 = p026C
  If gv2382 <= 0 Then
    gv237A.M1C99 = 0
    gv237A.M2EF9 = 0
  Else
    gv237A.M1C99 = Asc(Mid$(gv7856, gv2382, 1))
    gv237A.M2F00 = Asc(Mid$(gv7856, gv2382 + 1, 1))
    gv237A.M2F09 = Asc(Mid$(gv7856, gv2382 + 2, 1))
    gv237A.M2EF9 = gv23AA.M267E(gv2382 \ 3)
  End If
  gv15EA = gv237A.M2EF9 And &H1FF
  LSet gv15E0 = gv1646.M2D10(gv15EA)
  gv15E0.M2CC3 = (gv15E0.M2CAB And &HF000) \ &H1000
  gv15E0.M2CCB = (gv15E0.M2CB3 And &HF000) \ &H1000
End Sub

Function fn0AA5(p0260 As T265D) As String
Dim l0262 As String
Dim l0264 As Integer
Dim l0266 As Integer
  If p0260.M2678 <= 1 Then Exit Function
  l0264 = p0260.M2678
  l0262 = Mid$(gc114A, l0264, 1)
  l0266 = p0260.M267E
  If l0266 = gc11A0 Then
    l0262 = l0262 + "a" + fn02B5(p0260.M2672 * 2)
  Else
    If l0266 And mc020C Then
      l0262 = l0262 + "a"
      l0266 = l0266 And &HF
    ElseIf l0264 <> gc115A Then
      l0262 = l0262 + "v"
    End If
    l0262 = l0262 + fn02B5(p0260.M2672 * 2)
    If l0266 > 1 Then
      l0262 = l0262 + Mid$(gc117C, l0266, 1)
    End If
  End If
  fn0AA5 = l0262
End Function

Function fn0AB0() As String
Dim l0154 As String
Dim l0156 As Integer
  l0156 = fn0A66()
  If gv237A.M2EF9 Then l0154 = fn0BCF(gv15EA)
  If Len(l0154) = 0 Then
    l0154 = "****"
  End If
  l0154 = Chr$(9) & l0154
  If gv1140 >= 0 Then l0154 = " " & gv110E(gv1140).M268A & l0154
  If gv15E0.M2CC3 < 0 Then
    Seek gv088E, gv0920
  Else
    gv0924 = gv0924 + 2
    Seek gv088E, gv0924
    While gv0924 < gv0920
      Get gv088E, , gv091E
      l0154 = l0154 & " +" & fn02B5(gv091E)
      gv0924 = gv0924 + 2
    Wend
  End If
  fn0AB0 = l0154
End Function

Sub sub01D7()
  sub019A gv0920, gv0928 'get offset i end for sub
  If gv0920 <= 1 Then
    sub01F8
    Exit Sub
  End If
  Seek gv088E, gv0920
If gv7A2A Then
    sub0A44
    Do While gv0920 < gv0928
      If fn0A66() = mc0192 Then
        Exit Do
      End If
    Loop
    sub0A56
Else
    gv0988 = 0
    ReDim gv0956(gv7AF2.M15A4 + 20)
    sub05F1
    gv0956(gv0988 + 1).M1E3E = gv0920
    sub06C9 gv0928 - gv0920, 1
End If
End Sub

Sub sub0995()
  sub04E2 False
End Sub

Sub sub09C7()
  gv0924 = gv0920 'load operand
  Get gv088E, gv0920, gv2382
  sub06D9 gv2382  'analyze operand
End Sub

Sub sub0AC3()
Dim l0168 As Integer
Dim l016A As Integer
Dim l016C As Integer
Dim l016E As Integer
' Const mc0192 = 48 ' &H30%
  gv7A2A = 1
  If gv0BC4 Then gv7A2A = gv7A2A Or gc1146
  frmMain.control7.Value = True
  sub0201 l016C, gc0406
  For l016A = 1 To gv7AA2
    sub047C l016C, gc0406 & " " & fn0D20(l016A)
    gv7B2A = 0
    sub01A8 l016A
    For l0168 = 1 To gv7AEE
      If gv7ABC(l0168).M24AC = gv7AB8 Then
        If gv7ABC(l0168).M13E4 <> -1 Then
          sub0201 l016E, fn0D70(l0168) 'add text
          sub01C8 l0168                'scan sub
          sub0220 l016E                'remove text
        End If
      End If
    Next
  Next
  gv7A2A = 0
  frmMain.control6.Value = True
  sub0220 l016C
  gv7AB6 = 0
End Sub

Sub sub0AD1()
Dim l0268 As Integer
Dim l026A As Long
  l026A = 1
  For l0268 = gc118E To gc11A2
    gv11AC(l0268) = l026A
    l026A = l026A * 2
  Next
  l026A = &H10000
  For l0268 = gc118E To gc1198
    gv11AC(l0268 + 16) = l026A
    l026A = l026A * 2
  Next
  l026A = mc00E0
  For l0268 = gc119A To gc11A2
    gv11AC(l0268 + 16) = l026A
    l026A = l026A * 2
  Next
  gv11AC(gc118C) = mc010C
  gv11AC(gc11A4) = mc0114
  gv11AC(gc11A6) = mc0118
  gv11AC(gc11DA) = mc0110
End Sub

Sub sub0AE0()
  LSet gv15E0 = gv1646.M2D10(gv15EA)
  gv15E0.M2CC3 = (gv15E0.M2CAB And &HF000) \ &H1000
  gv15E0.M2CCB = (gv15E0.M2CB3 And &HF000) \ &H1000
End Sub

