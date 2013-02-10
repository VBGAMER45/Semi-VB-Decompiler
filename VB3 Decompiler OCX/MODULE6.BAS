Attribute VB_Name = "Module6"
' Module6
Option Explicit
Dim m001C As Integer
Dim m001E As Integer
Dim m0026 As Integer
Dim m0032 As Long
Dim m0036 As Single
Dim m003A  As Double
Dim m0042  As Currency
Dim m0062 As Integer
Dim m0064 As Integer
Dim m0066   As Integer
Dim m0068    As Integer
Dim m006A    As Integer
Dim m006C As Integer
Dim m006E   As Integer
Dim m0070 As String
Dim m0074 As String
Dim m0078 As Integer
Dim m007A As Integer
Dim m007C(255) As Integer
Dim m0092() As String
Dim m00C4 As Integer
Const mc00CE = 1000 ' &H3E8%
Const mc00E0 = 32 ' &H20%
Const mc00F4 = 10 ' &HA%
Const mc00FA = 8 ' &H8%
Const mc0104 = 9 ' &H9%
Const mc010E = 8192 ' &H2000%
Const mc0132 = 1 ' &H1%
Const mc0134 = 2 ' &H2%
Const mc0136 = 4 ' &H4%
Const mc0138 = 5 ' &H5%
Const mc013C = 8 ' &H8%
Const mc013E = 10 ' &HA%
Const mc0140 = 11 ' &HB%
Const mc0142 = 12 ' &HC%
Const mc0144 = 14 ' &HE%
Const mc0146 = 13 ' &HD%
Const mc014C = 256 ' &H100%
Const mc014E = 1 ' &H1%
Const mc0150 = 3 ' &H3%
Const mc0152 = 1024 ' &H400%
Const mc0156 = 9 ' &H9%
Const mc0158 = 4 ' &H4%
Const mc015A = 5 ' &H5%
Const mc015C = 6 ' &H6%
Const mc015E = 7 ' &H7%
Const mc0160 = 8 ' &H8%
Const mc0162 = 10 ' &HA%
Const mc0164 = 2 ' &H2%
Const mc0166 = 512 ' &H200%
Const mc016A = 256 ' &H100%
Const mc0170 = 32 ' &H20%
Const mc01C8 = 9 ' &H9%
Const mc022E = 224 ' &HE0%
Const mc0230 = 192 ' &HC0%
Const mc0232 = 16 ' &H10%
Const mc0234 = 17 ' &H11%
Const mc0248 = 32 ' &H20%
Const mc0264 = 12 ' &HC%
Const mc0266 = 10 ' &HA%
Const mc026C = 1 ' &H1%

Private Function fn0752(p01A8 As Integer) As String
    Select Case p01A8
    Case -6 To -1
        fn0752 = "L" & Hex$(gv0BDA) & ":"
    Case Else
      fn0752 = Format$(fn0335(p01A8))
    End Select
End Function

Private Function fn075C(p01AC As Integer) As String
Dim l01AE As Long
Dim l01B0 As Integer
    l01AE = Seek(m0026)
    Get m0026, gv06B4 + fn0335(p01AC) + 4, l01B0
    Select Case l01B0
    Case -1
        fn075C = "L" & Hex$(fn0335(p01AC) + gv06B4 - 1)
    Case Else
      fn075C = Format$(fn0335(l01B0))
    End Select
    Seek m0026, l01AE
End Function

Function fn0766(p01B6 As String) As String
  If m0074 = "" Then
    fn0766 = m0070 & p01B6
  Else
    fn0766 = m0070 & m0074
    m0074 = ""
  End If
  m0070 = ""
End Function

Function fn0772() As String
  If m00C4 <= 0 Then
    fn0772 = "<missed arg>"
  Else
    fn0772 = m0092(m00C4)
    m0092(m00C4) = ""
    m00C4 = m00C4 - 1
  End If
End Function

Private Function fn0779(p01C0 As Integer, p01C2 As Integer, p01C4 As Integer) As String
Dim l01C6 As String
' Const mc01C8 = 9 ' &H9%
  If gv0BC8 Then
    If gv237A.M2F09 = mc01C8 Then
      l01C6 = fn07D5(p01C0)
      If p01C2 Then sub0671
    Else
      l01C6 = fn0CDA(fn032D(p01C0, 2))
    End If
  Else
    If p01C0 < 0 Then
      If p01C0 <> -2 Then sub01F8
      Exit Function
    End If
    l01C6 = fn07E1(p01C0)
  End If
  If Len(l01C6) Then
    If p01C4 Then
      fn0779 = "!" & l01C6
    Else
      fn0779 = "." & l01C6
    End If
  End If
End Function

Function fn0786(p01CC As Integer) As String
  If p01CC > 0 And Len(gv0BA6) Then
    fn0786 = "!" & fn03DF(p01CC)
    sub0671
  Else
    fn0786 = "!?" & fn02B5(p01CC)
  End If
End Function

Function fn07D5(p0238 As Integer) As String
  fn07D5 = fn0B6F(p0238)
End Function

Function fn07E1(p023E As Integer) As String
Dim l0240 As Long
Dim l0242 As Integer
Dim l0244 As Integer
' Const mc0248 = 32 ' &H20%
  l0240 = fn0335(p023E)
    If p023E = 0 Then Beep: Exit Function
    If l0240 > gv09F6 Then Beep: Exit Function
      l0244 = l0240 \ 2
      l0242 = Asc(Mid$(gv7952, l0244, 1))
      If (l0242 And &H1F) = 0 Or (l0242 And mc022E) = mc0248 Then
        sub07EB p023E, gv237A.M2F00, gv237A.M2F09
      End If
      fn07E1 = fn0DCA(l0244)
End Function

Sub sub05C5(p024A As Integer)
  If p024A Then sub05D1 p024A Else sub05E0 ""
  sub05E8 "(" & fn0772() & ")"
End Sub

Sub sub05D1(p00DA As Integer)
Dim l00DC As String
' Const mc00E0 = 32 ' &H20%
  If m001C Or (p00DA And 1) Then sub067F p00DA: Exit Sub
  While p00DA > 0
    l00DC = fn0772() & l00DC
    If m0092(m00C4) = gc157A Then
      m00C4 = m00C4 - 1
    Else
      l00DC = fn0772() & " To " & l00DC
    End If
    p00DA = p00DA - 2
    If p00DA Then l00DC = ", " & l00DC
  Wend
  sub05E0 l00DC
  p00DA = 1
End Sub

Sub sub05E0(p01D6 As String)
On Error GoTo 501
  m00C4 = m00C4 + 1
  m0092(m00C4) = p01D6
Exit Sub

501 If m00C4 < 0 Then
      m00C4 = 1
    Else
      ReDim Preserve m0092(1 To UBound(m0092) + 10)
    End If
    Resume
End Sub

Sub sub05E8(p01EA As String)
  If m001E Then Exit Sub
  gv0C5A = gv0C5A & p01EA
End Sub

Sub sub05F1()
Dim l024E As Integer
  l024E = gv0BC8
  gv0C5E = 0
  gv0BDA = 0
  gv0BC8 = True
  If gv7AF2.M1511 And 1 Then
    gv0BDE = "Sub"
  Else
    gv0BDE = "Function"
  End If
  If gv7A2A Then
  Else
    If gv7AF2.M1519 And &H40 Then sub05FF "Private "
    sub05FF gv0BDE & " " & fn0D70(gv7B2A)
  End If
  sub0607 gv7B2A
  If gv0BCE = 0 Then sub0616
  gv0BC8 = l024E
End Sub

Sub sub0625(p0258 As Integer, p025A As Integer)
Dim l025C As Integer
Dim l025E As String
Dim l0260 As Integer
Dim l0262 As Integer
' Const mc0264 = 12 ' &HC%
' Const mc0266 = 10 ' &HA%
' Const mc026C = 1 ' &H1%
  On (p0258 \ &H10) + 1 GoTo LC620, LCBCE, LD2C8, LD49E, LD61A, LD912, LDB28, LDC9A, LDE0C, LE60C, LE6C4, LE95E, LE9F2, LE9FE, LEA0A, LEC9C, LEE1C, LEE3C, LEF9A, LEFA6, LF0F6, LF20C

LC60A:
  Error mc00CE
Exit Sub

LC620: On p0258 And &HF GoTo LC66C, LC684, LC68A, LC696, LC6D2, LC70E, LC738, LC76E, LC774, LDD5A, LC782, LC898, LC8F0, LC95C, LCB52
    sub0632 p025A: p025A = 0
    Exit Sub
LC66C:
    sub0632 m0066
    Exit Sub
LC684:
LC68A:
    GoTo LC60A
LC696:
    sub05E8 fn0752(m0068) & " "
    Exit Sub
LC6D2:
    sub05E8 fn0752(m0068) & Space$(m006A)
    Exit Sub
LC70E:
    sub05E8 ": "
    Exit Sub
LC738:
    sub05E8 ": "
    sub063C m0066
    Exit Sub
LC76E:
LC774:
    GoTo LC60A

LC782:
    sub0642 m0066, p025A: p025A = 0
      If gv237A.M2F09 = mc01C8 Then
        l0260 = fn032D(m0066, 2)
        l0262 = gv09C4(l0260 - 1)
        Select Case l0262 And &HC000
        Case 0, &H8000
          sub064E l0262
        Case &H4000
          If l0262 = &H4006 Then
            sub064E l0262
            l025E = fn07D5(gv09C4(l0260))
          Else
            sub064E l0262
          End If
        Case Else
          sub01F8
        End Select
      End If
    Exit Sub

LC898:
    sub0659 m0066, p025A, 0
    p025A = 0
    sub05E8 " = " & fn0772()
  Exit Sub
LC8F0:
    sub0659 m0066, p025A, -1
    p025A = 0
      If gv237A.M2F09 = mc01C8 Then
        sub064E gv09C4(fn032D(m0066, 2) - 1)
      End If
  Exit Sub

LC95C:
    l025C = m0066
    If l025C >= 0 Then
        l0262 = fn032D(m0068, 2)
        gv0BFC.M13DD = l0262
        Select Case gv237A.M2F09
        Case mc0264
          sub0664 fn07E1(l0262 * 2), p025A, l025C
          Exit Sub
        Case mc01C8
          sub064E gv09C4(l0262 - 1)
          sub0671
        Case mc0266
          Select Case m006E
          Case &H7E0 To &H816
            If l025C Then sub05D1 l025C
          Case Else
  If l0262 <= Len(gv7952) Then
            If (Asc(Mid$(gv7952, l0262, 1)) And &HF) = 9 Then
             sub064E gv09C4(l0262 - 1)
            End If
  End If
          End Select
        Case Else
          sub01F8
        End Select
        If l025C > 1 Then sub067F l025C
      sub0664 fn07E1(m0068), p025A, l025C
    Else
      sub0642 m0068, p025A
      p025A = 0
    End If
    Exit Sub
LCB52:
    sub0664 fn07E1(m0068), p025A, m0066
    sub0689 fn0772()
    sub05E8 " = " & fn0772()
    Exit Sub
LCBCE: On p0258 And &HF GoTo LCCC4, LCEB0, LCFB6, LD068, LD0EE, LD15A, LD1E8, LD1F4, LD202, LD210, LD264, LD270, LD2B0, LD2B6, LD2BC
    sub0659 m0068, p025A, 0
    sub05C5 m0066
    If p025A = 0 Then
      If m007A Then
        sub0659 0, 0, -1
      Else
        sub0659 0, 0, 0
      End If
    Else
      p025A = 0
    End If
    Exit Sub
LCC8A:
    sub0659 m0068, p025A, 0
    sub05C5 m0066
    m0078 = -1
    Exit Sub
LCCC4:
    Select Case p025A
    Case 0 To 7
      If gv237A.M2F09 = mc01C8 Then
        sub0693 "." & fn07D5(m0066)
      Else
        sub0693 fn0779(m0066, False, 0)
      End If
      If p025A Then
        sub0693 Mid$(gc1576, p025A, 1)
        p025A = 0
      End If
    Case 9
      If m0066 < 0 Then
        sub0693 "!" & fn07D5(m0066)
      ElseIf m0066 And 1 Then
        sub0693 fn0779(m0066, False, p025A)
      ElseIf iVBVersion < 3 Then
        sub0693 "!" & fn07D5(m0066)
      Else
        sub0693 fn0786(m0066)
      End If
      p025A = 0
    Case Else
      sub01F8
    End Select
    Exit Sub
LCEB0:
    If p025A = 9 Then
      sub0689 fn0772() & fn0779(m0066, False, p025A)
      p025A = 0
    Else
      sub0689 fn0772() & fn0779(m0066, False, 0)
      If p025A Then
        sub05E8 Mid$(gc1576, p025A, 1)
        p025A = 0
      End If
    End If
    sub05E8 " = " & fn0772()
    Exit Sub
LCFB6:
    l025E = fn0772() & fn0779(m0068, True, p025A = 9)
    If m0066 >= 0 Then
      sub0664 l025E, p025A, m0066
    Else
      If p025A Then
        l025E = l025E & Mid$(gc1576, p025A, 1)
        p025A = 0
      End If
      sub05E0 l025E
    End If
    Exit Sub
LD068:
    l025E = fn0772() & fn0779(m0068, True, p025A = 9)
    sub0664 l025E, p025A, m0066
    sub069E fn0772() & " = ", 1
    Exit Sub
LD0EE:
    gv0BFC.M13DD = m0066
    gv0BFC.M220A = m0068
    If m007A Then sub01F8
    m007A = p0258
    If m0078 Then
      sub0659 m0078, 0, -1
    End If
    Exit Sub
LD15A:
    If m007A Then sub01F8
    gv0BFC.M13DD = m0066
    If gv0BC8 And gv0BFC.M13DD > 7 Then gv0BFC.M13DD = 0
    gv0BFC.M220A = m0068
    m007A = p0258
    If m0078 Then
      sub0659 m0078, 0, -1
    End If
    Exit Sub
LD1E8:
  GoTo LC60A
LD1F4:
  GoTo LC60A
LD202:
  GoTo LC60A
LD210:
LD216:
    sub069E fn0BDF(), 1
    If m007A Then
      sub0659 0, 0, -1
    End If
    Exit Sub
LD264:
    GoTo LC60A
LD270:
    sub05E8 "Type " & fn07E1(m0068)
    Exit Sub
LD2B0:
LD2B6:
LD2BC:
    GoTo LC60A
LD2C8:
  On p0258 And &HF GoTo LD2FC, LD302, LD30E, LD334, LD3C8, LD3CE, LD3DA, LD428, LD42E, LD434, LD43A, LD440, LD446, LD64E, LD492
LD2FC:
LD302:
  GoTo LC60A
LD30E:  sub05E8 "Call "
LD334:
      l025E = fn0D70(gv0DE6((fn0E69(m0068)) + mc026C))
    If m0066 And p0258 = &H23 Then
      sub0664 l025E, 0, m0066
      sub05E8 fn0772()
    Else
      sub069E l025E, m0066
    End If
    Exit Sub
LD3C8:
LD3CE:
    GoTo LC60A
LD3DA:
    sub06AA "To"
    sub069E "Case", 1
    Exit Sub
LD428:
LD42E:
LD434:
LD43A:
LD440:
LD446:
    sub06B3 fn0BDF()
    sub069E "Case", 1
    Exit Sub

LD492:
    GoTo LC60A
LD49E: On p0258 And &HF GoTo LD4D2, LD4D8, LD4DE, LD4E4, LD4EA, LD6FA, LD4F2, LD4F8, LD504, LD53C, LD542, LD548, LD554, LD5B2, LD58C

LD4D2:
LD4D8:
LD4DE:
LD4E4:
LD4EA:

LD4F2:
LD4F8:
  GoTo LC60A
LD504:
  sub05E8 "End " & gv0BDE
  Exit Sub
LD53C:
LD542:
LD548:
    GoTo LC60A
LD554:
  sub0689 "Exit " & gv0BDE
  Exit Sub
LD58C:
    sub06AA "Step"
LD5B2:
    sub06AA "To"
    sub06AA "="
    sub069E "For ", 1
    Exit Sub
LD61A: On p0258 And &HF GoTo LD666, LD6A8, LD6BA, LD6AE, LD700, LD75E, LD764, LD770, LD706, LD7D0, LD7D6, LD7DC, LD7E8, LD826, LD832

LD64E:
LD656:
    Exit Sub

LD666:
  sub0689 "GoSub " & fn075C(m0066)
  Exit Sub
LD6A8:
LD6AE:
  GoTo LC60A
LD6BA:
    sub0689 "GoTo " & fn075C(m0066)
    Exit Sub
LD6FA:
LD700:
LD706:
    sub0689 fn0BDF() & " " & fn0772() & " Then "
    Exit Sub
LD75E:
LD764:
  GoTo LC60A
LD770:
    sub0689 "If " & fn0772() & " GoTo " & fn075C(m0066)
    Exit Sub

LD7D0:
LD7D6:
LD7DC:
  GoTo LC60A
LD7E8:
    sub069E fn0BDF(), 0
    If gv0BC8 Then m00C4 = m00C4 - 1
    Exit Sub
LD826:
  GoTo LC60A
LD832:
    sub0689 fn0BDF()
    Select Case m0066
    Case -1
      sub05E8 "GoTo 0"
    Case -2
      sub05E8 "Resume Next"
    Case Else
      sub05E8 "GoTo " & fn075C(m0066)
    End Select
    Exit Sub
LD912: On p0258 And &HF GoTo LD944, LDA08, LDA14, LDA56, LDA5C, LDA62, LDA68, LD656, LDA70, LDA76, LDA82, LDAAE, LDAB4, LDABA, LDAC0
LD944:
    sub0689 "On " & fn0772() & fn0BDF()
    sub05E8 fn075C(m007C(1))
    For l0260 = 2 To m007C(0)
      sub05E8 ", " & fn075C(m007C(l0260))
    Next l0260
    Exit Sub
LDA08:
    GoTo LC60A
LDA14:
    sub0689 "Resume " & fn075C(m0066)
    Exit Sub
LDA56:
LDA5C:
LDA62:
LDA68:

LDA70:
LDA76:
  GoTo LC60A
LDA82:
    sub0693 ", "
    Exit Sub
LDAAE:
LDAB4:
LDABA:
LDAC0:
    l025E = fn0772()
    While m006C > gv15EA
      l025E = ", " & l025E
      m006C = m006C - 1
    Wend
    sub0693 l025E
    Exit Sub
LDB28: On p0258 And &HF GoTo LDB5A, LDB60, LDB66, LDBD0, LDBF4, LDC22, LDC3A, LDC40, LDC6A, LDC70, LDC76, LDC7C, LDC82, LDC88, LDC8E
LDB5A:
LDB60:
LDB66:
    If p0258 And 1 Then l025E = "Step"
    If p0258 And 2 Then l025E = "-" & l025E
    sub0664 l025E, 0, 2
    Exit Sub
LDBD0:
    m0070 = "Debug."
    Exit Sub
LDBF4:
    m0070 = fn0772() & "."
    Exit Sub
LDC22:
    m006C = &H60
    Exit Sub
LDC3A:
LDC40:
    sub069E fn0BDF(), m00C4
    Exit Sub
LDC6A:
LDC70:
LDC76:
LDC7C:
LDC82:
LDC88:
LDC8E:
  GoTo LC60A
LDC9A: On p0258 And &HF GoTo LDCCC, LDCE4, LDCD2, LDCD8, LDCEA, LDD60, LDD96, LDDA4, LDDB2, LDDB8, LDDC4, LDDEE, LDDFA, LDE00

LDCCC:
LDCD2:
LDCD8:
  GoTo LC60A
LDCE4:
LDCEA:
    While m00C4 > 1
      sub0693 fn0772()
    Wend
    sub069E fn0766("Print"), m00C4
    Exit Sub
LDD5A:
LDD60:
    m001E = m001C
    sub06BD m0066, fn0772()
    Exit Sub
LDD96:
    Exit Sub
LDDA4:
    Exit Sub
LDDB2:
LDDB8:
    GoTo LC60A
LDDC4:
    sub069E fn0BDF(), m0066
    Exit Sub
LDDEE:
LDDF4:
LDDFA:
LDE00:
    GoTo LC60A
LDE0C: On p0258 And &HF GoTo LDE86, LDE8C, LDE92, LDEB2, LDE9E, LDFBC, LDFC8, LDFF4, LE0D2, LE13C, LE142, LE24A, LE256, LE2D8, LE29E
LDE3E:
    sub06B3 ", "
    sub069E fn0BDF(), 2
    Exit Sub
LDE86:
LDE8C:
LDE92:
    GoTo LC60A
LDE9E:
    l025E = fn0772()
LDEB2:
    If m00C4 > 1 Then sub0693 fn0772()
    Select Case m0066
    Case 0
    Case 1: l025E = l025E & ", B"
    Case 2: l025E = l025E & ", BF"
    Case Else: sub01F8
    End Select
    If Len(l025E) Then sub05E0 l025E
    sub069E fn0766(fn0BDF()), m00C4
    Exit Sub
LDFBC:
    GoTo LC60A
LDFC8:
    sub069E fn0BDF(), m00C4
    Exit Sub
LDFF4:
LDFFA:
    Select Case m0066 And &HFFFE
    Case &H8002
    Case &H4002
      l025E = "To " & fn0772()
      m0092(m00C4) = l025E
    Case &H2
      sub06AA "To"
    Case Else: sub01F8
    End Select
    sub069E fn0BDF(), 2
    Exit Sub
LE0D2:
LE0D8:
    sub0689 fn0BDF() & fn0772() & " = "
    sub05E8 fn0772()
    Exit Sub

LE13C:
LE142:
    If p025A Then
      sub0689 "Mid(" & fn0772() & ", "
      p025A = 0
    Else
      sub0689 "Mid$(" & fn0772() & ", "
    End If
    l025E = fn0772()
    If p0258 And 1 Then sub067F 2
    sub05E8 fn0772() & ") = " & l025E
    Exit Sub
LE24A:
    GoTo LC60A
LE256:
    sub06AA "As"
    sub069E fn0BDF(), 1
    Exit Sub
LE29E:
    sub0693 " Len = " & fn0772()
LE2D8:
    l025E = "As " & fn0772()
    l0262 = m0066
    If l0262 And &H8CD0 Then sub01F8
    Select Case l0262 And &H7000
    Case &H1000: l025E = "Lock Read Write " & l025E
    Case &H2000: l025E = "Lock Write " & l025E
    Case &H3000: l025E = "Lock Read " & l025E
    Case &H4000: l025E = "Shared " & l025E
    End Select
    If l0262 And &H300 Then
      If l0262 And &H200 Then l025E = "Write " & l025E
      If l0262 And &H100 Then l025E = "Read " & l025E
      l025E = "Access " & l025E
    End If
    Select Case l0262 And &H3F
    Case 1: l025E = "Input " & l025E
    Case 2: l025E = "Output " & l025E
    Case 4: l025E = "Random " & l025E
    Case 8: l025E = "Append " & l025E
    Case &H20: l025E = "Binary " & l025E
    End Select
    sub0689 fn0BDF() & fn0772() & " For " & l025E
    Exit Sub
LE60C: On p0258 And &HF GoTo LE63E, LE644, LE64A, LDE3E, LE652, LE658, LE65E, LE66A, LE69C, LE6A2, LE0D8, LE6AA, LE6B0, LDFFA, LE6B8
LE63E:
LE644:
LE64A:

LE652:
LE658:
LE65E:
    GoTo LC60A
LE66A:
    m001E = m001C
    sub0689 fn0772()
    Exit Sub
LE69C:
LE6A2:

LE6AA:
LE6B0:

LE6B8:
    GoTo LC60A
LE6C4: On p0258 And &HF GoTo LE7E4, LE7F0, LE80C, LE812, LE818, LE81E, LE824, LE82A, LE830, LE836, LE842, LE940, LE946, LE94C, LE952
    l025E = fn0772() & ")"
    l025E = fn0772() & ", " & l025E
    l025E = fn0772() & ")-(" & l025E
    l025E = " (" & fn0772() & ", " & l025E
    sub069E fn0766(fn0BDF()) & l025E, 0
    Exit Sub
LE7E4:
    GoTo LC60A
LE7F0:
    m0074 = fn0BDF()
    Exit Sub
LE80C:
LE812:
LE818:
LE81E:
LE824:
LE82A:
LE830:
LE836:
  GoTo LC60A
LE842:
      If m00C4 > 0 Then
      If p025A And &H20 Then
        p025A = InStr(m0092(m00C4), "$")
        m0092(m00C4) = Left$(m0092(m00C4), p025A - 1) & Mid$(m0092(m00C4), p025A + 1)
      ElseIf p025A Then
        m0092(m00C4) = "C" & gv157E(p025A) & "(" & m0092(m00C4) & ")"
      End If
      End If
      p025A = 0
      Exit Sub
LE940:
LE946:
LE94C:
LE952:
  GoTo LC60A
LE95E: On p0258 And &HF GoTo LE992, LE998, LE99E, LE9A4, LE9AA, LE9B0, LE9B6, LE9BC, LE9C2, LE9C8, LE9CE, LE9D4, LE9DA, LE9E0, LE9E6

LE992:
LE998:
LE99E:
LE9A4:
LE9AA:
LE9B0:
LE9B6:
LE9BC:
LE9C2:
LE9C8:
LE9CE:
LE9D4:
LE9DA:
LE9E0:
LE9E6:
    GoTo LC60A
LE9F2: GoTo LC60A

LE9FE: GoTo LC60A

LEA0A: On p0258 And &HF GoTo LEA3C, LEA42, LEA48, LEA54, LEA82, LEAA8, LEAE0, LEB18, LEB5E, LEB96, LEBDC, LEC14, LEC84, LEC90

LEA3C:
LEA42:
LEA48:
    GoTo LC60A
LEA54:
    sub05E0 Format$(p025A): p025A = 0
    Exit Sub
LEA82:
    sub05E0 Format$(m0066)
    Exit Sub
LEAA8:
    sub05E0 Format$(m0032) & "&"
    Exit Sub
LEAE0:
    sub05E0 "&H" & Hex$(m0066)
    Exit Sub
LEB18:
    sub05E0 "&H" & Hex$(m0032) & "&"
    Exit Sub
LEB5E:
    sub05E0 "&O" & Oct$(m0066)
    Exit Sub
LEB96:
    sub05E0 "&O" & Oct$(m0032) & "&"
    Exit Sub
LEBDC:
    sub05E0 Trim$(Str$(m0036)) & "!"
    Exit Sub
LEC14:
    sub05E0 Trim$(Str$(m003A)) & "#"
    Exit Sub
LEC4C:
    sub05E0 Trim$(Str$(m0042)) & "@"
    Exit Sub
LEC84:
LEC8A:
LEC90:
  GoTo LC60A
LEC9C: On p0258 And &HF GoTo LECD0, LECD6, LECDC, LECE2, LECE8, LECEE, LECF4, LED00, LED96, LED9C, LEDA8, LEDFE, LEE04, LEE0A, LEE10

LECD0:
LECD6:
LECDC:
LECE2:
LECE8:
LECEE:
LECF4:
    GoTo LC60A
LED00:
    sub06B3 "TypeOf "
    sub0693 " Is "
      l025E = gv0FBC(m0066).M1D68
      If Len(l025E) = 0 Then l025E = Hex$(m0066): sub01F8
    sub0693 l025E
    Exit Sub
LED96:
LED9C:
    GoTo LC60A
LEDA8:
    gv0BFC.M13DD = -m0068
    gv0BFC.M220A = m006A
    m007A = True
    If m0078 Then sub0659 m0078, 0, -1
    Exit Sub
LEDFE:
LEE04:
LEE0A:
LEE10:
    GoTo LC60A
LEE1C: If p0258 = &H107 Then GoTo LEC4C
  GoTo LC60A

LEE3C: On p0258 And &HF GoTo LEE70, LEE76, LEE7C, LEE82, LEE88, LEE8E, LEE94, LEE9A, LEEA0, LEEAC, LEEB2, LEF38, LEF44, LEF88, LEF8E

LEE70:
LEE76:
LEE7C:
LEE82:
LEE88:
LEE8E:
LEE94:
LEE9A:
LEEA0:
    GoTo LC60A
LEEAC:
LEEB2:
    l0260 = m0066
    l025E = fn0BB2(l0260)
    If p0258 And 1 Then l025E = fn0766(l025E)
    If l0260 >= 0 Then
      sub069E l025E, m0068
    Else
      sub0664 l025E, 0, m0068
    End If
    Exit Sub
LEF38:
    GoTo LC60A
LEF44:
    If m0066 <> &HFC Then GoTo LC60A
    sub069E "Dim Shared ", 0
    Exit Sub
LEF88:
LEF8E:
  GoTo LC60A
LEF9A:
  GoTo LC60A
LEFA6: On p0258 And &HF GoTo LEFD8, LEFDE, LEFEA, LF044, LCC8A, LF052, LF0BA, LF0C0, LF0C6, LF0CC, LF0D2, LF0D8, LF0E4, LF0EA

LEFD8:
LEFDE:
    GoTo LC60A
LEFEA:
    sub069E "Set", 1
    sub05E8 " = " & fn0772()
    Exit Sub
LF044:
    GoTo LC60A

LF052:
    If m0066 Then l025E = "Text" Else l025E = "Binary"
    sub069E fn0BDF() & l025E, 0
    Exit Sub
LF0BA:
LF0C0:
LF0C6:
LF0CC:
LF0D2:
LF0D8:
LF0DE:
LF0E4:
LF0EA:
  GoTo LC60A
LF0F6: On p0258 And &HF GoTo LF12A, LF130, LF136, LD216, LF13E, LF14A, LF1A2, LF1D4, LF1DA, LF1E0, LF1E6, LF1EC, LF1F2, LF1F8, LF1FE

LF12A:
LF130:
LF136:

LF13E:
    GoTo LC60A
LF14A:
  sub064E m0066
  l025E = gv0FBC(m0066).M1D68
  sub05E0 "New " & l025E
  Exit Sub
LF1A2:
    sub05E0 fn0BDF()
    sub064E &H4006
    Exit Sub
LF1D4:
LF1DA:
LF1E0:
LF1E6:
LF1EC:
LF1F2:
LF1F8:
LF1FE:
  GoTo LC60A

LF20C: On p0258 And &HF GoTo LF23E, LF244, LF24A, LF250, LF256, LF25C, LF262, LF268, LF26E, LF274, LF27A, LF280, LF286, LF292, LF2DE
LF23E:
LF244:
LF24A:
LF250:
LF256:
LF25C:
LF262:
LF268:
LF26E:
LF274:
LF27A:
LF280:
LF286:
  GoTo LC60A
LF292:
  l025E = fn0772() & "." & fn03DF(m0068)
  sub069E l025E, m0066
Exit Sub
LF2DE:
  GoTo LC60A
End Sub

Private Sub sub0642(p01E0 As Integer, p01E2 As Integer)
  sub05E0 fn07E1(p01E0)
  If p01E2 Then m0092(m00C4) = m0092(m00C4) & Mid$(gc1576, p01E2, 1)
End Sub

Private Sub sub0659(ByVal pv00E2 As Integer, p00E4 As Integer, p00E6 As Integer)
Dim l00E8 As Integer
' Const mc00F4 = 10 ' &HA%
' Const mc00FA = 8 ' &H8%
' Const mc0104 = 9 ' &H9%
' Const mc010E = 8192 ' &H2000%
  If p00E4 = 0 And p00E6 = -1 And m007A Then
    p00E4 = gv0BFC.M13DD
    p00E6 = gv0BFC.M220A
    l00E8 = True
    m007A = False
  End If
  If p00E6 = -1 Then
    p00E6 = 0
  ElseIf p00E6 < -1 Then
    l00E8 = True
  End If
  If p00E4 Then m0078 = 0
  If pv00E2 Then
    If m0062 Then sub05E8 ", " Else m0062 = True
    sub05E8 fn07E1(pv00E2)
    If Not l00E8 And p00E4 > 0 And p00E4 < 8 And p00E4 <> 6 Then
      sub05E8 Mid$(gc1576, p00E4, 1)
      Exit Sub
    End If
  End If
  If p00E4 = 0 Then Exit Sub
  If p00E6 < 0 Then
    sub06BD (p00E6 And &H7F) - 1, " As New "
  Else
    sub06BD p00E6 - 1, " As "
  End If
  If p00E4 < 0 Then
    sub05E8 "String *" & Str$(-p00E4)
    Exit Sub
  End If
    pv00E2 = gv09C4(p00E4 - 1)
    p00E4 = Asc(Mid$(gv7952, p00E4, 1))
    Select Case p00E4 And &HF
    Case 0
      sub01F8
    Case mc00F4
        sub05E8 fn0CDA(fn032D(pv00E2, 2))
    Case mc00FA
      sub05E8 "String *" & Str$(fn0335(pv00E2))
    Case Is < 8
      sub05E8 gv1594(p00E4 And &HF)
    Case mc0104
      If gv0FBC(pv00E2 And &HFF).M25C6 < 0 Then
        sub022F mc010E
        sub05E8 "()"
      End If
      sub05E8 gv0FBC(pv00E2 And &HFF).M1D68
    Case Else
      sub01F8
      sub05E8 "?" & Hex$(p00E4)
    End Select
End Sub

Sub sub0664(p01D8 As String, p01DA As Integer, ByVal pv01DC As Integer)
Dim l01DE As String
  If p01DA Then
    p01D8 = p01D8 & Mid$(gc1576, p01DA, 1)
    p01DA = 0
  End If
  If pv01DC Then
    sub067F pv01DC
    l01DE = fn0772()
  End If
  sub05E0 p01D8 & "(" & l01DE & ")"
End Sub

Sub sub067F(p00D2 As Integer)
Dim l00D4 As String
Dim l00D6 As String
  If (p00D2 And &H7FFF) = 0 Then Exit Sub
  If Abs(p00D2) > m00C4 Then
    sub05E0 "<concat?>"
    Exit Sub
  End If
  If p00D2 < 0 Then
    p00D2 = -p00D2
    Do While p00D2
      l00D4 = fn0772(): p00D2 = p00D2 - 1
      If Left$(l00D4, 8) <> "<default" Then Exit Do
    Loop
    While p00D2
      l00D6 = fn0772(): p00D2 = p00D2 - 1
      If Left$(l00D6, 8) = "<default" Then
        l00D4 = ", " & l00D4
      Else
        l00D4 = l00D6 & ", " & l00D4
      End If
    Wend
  Else
    l00D4 = fn0772(): p00D2 = p00D2 - 1
    While p00D2 > 0
      l00D4 = fn0772() & ", " & l00D4
      p00D2 = p00D2 - 1
    Wend
  End If
  sub05E0 l00D4
  p00D2 = 1
End Sub

Sub sub0689(ByVal p01E8 As String)
  If m001E Then Exit Sub
  If Len(gv0C5A) Then
    If Right$(gv0C5A, 1) <> " " Then
      gv0C5A = gv0C5A & " " & p01E8
      Exit Sub
    End If
  End If
  gv0C5A = gv0C5A & p01E8
End Sub

Sub sub0693(p01BA As String)
  If m00C4 <= 0 Then sub05E0 "<missed arg>"
  m0092(m00C4) = m0092(m00C4) & p01BA
End Sub

Sub sub069E(p01E4, ByVal pv01E6 As Integer)
  If m0064 Then
    sub05E8 ", "
  Else
    m0064 = Len(gv0C5A) + 1
    sub0689 p01E4
  End If
  If pv01E6 Then
    sub067F pv01E6
    sub0689 fn0772()
  End If
End Sub

Sub sub06AA(ByVal p01A2 As String)
Dim l01A4 As String
  l01A4 = fn0772()
  If m00C4 < 1 Then
    m00C4 = 1
    m0092(m00C4) = "<missed arg>"
  End If
  m0092(m00C4) = m0092(m00C4) & " " & p01A2 & " " & l01A4
End Sub

Sub sub06B3(ByVal p01BC As String)
  If m00C4 > 0 Then m0092(m00C4) = p01BC & m0092(m00C4)
End Sub

Sub sub06BD(p01EC As Integer, ByVal p01EE As String)
  If m001E Then Exit Sub
  sub063C p01EC
  sub05E8 p01EE
End Sub

Sub sub06C9(p0110 As Long, p0112 As Integer)
Dim l0114 As String
Dim l0116 As Long
Dim l0118  As Integer
Dim l011A  As Integer
Dim l011C As Integer
Dim l011E  As Integer
Dim l0122 As Integer
' Const mc0132 = 1 ' &H1%
' Const mc0134 = 2 ' &H2%
' Const mc0136 = 4 ' &H4%
' Const mc0138 = 5 ' &H5%
' Const mc013C = 8 ' &H8%
' Const mc013E = 10 ' &HA%
' Const mc0140 = 11 ' &HB%
' Const mc0142 = 12 ' &HC%
' Const mc0144 = 14 ' &HE%
' Const mc0146 = 13 ' &HD%
' Const mc014C = 256 ' &H100%
' Const mc014E = 1 ' &H1%
' Const mc0150 = 3 ' &H3%
' Const mc0152 = 1024 ' &H400%
' Const mc0156 = 9 ' &H9%
' Const mc0158 = 4 ' &H4%
' Const mc015A = 5 ' &H5%
' Const mc015C = 6 ' &H6%
' Const mc015E = 7 ' &H7%
' Const mc0160 = 8 ' &H8%
' Const mc0162 = 10 ' &HA%
' Const mc0164 = 2 ' &H2%
' Const mc0166 = 512 ' &H200%
Static m0168 As Integer
' Const mc016A = 256 ' &H100%
' Const mc0170 = 32 ' &H20%
  gv0BC8 = p0112 And gc0BCC
  m001C = gv0BCE And gc0BD2
  m001E = False
  m00C4 = 0
  ReDim m0092(1 To 50)
On Error GoTo LACEE
    m0026 = gv088E
  gv0C5E = -1
  l0116 = Seek(m0026) + p0110
  Get m0026, , m006E
  While Seek(m0026) < l0116

LA5A2:
    l011C = m006E
      gv0924 = Seek(m0026) - 2
      sub06D9 l011C
        l011C = gv237A.M2EF9
        If gv237A.M1C99 = 0 Then
        End If
    l0122 = (l011C And &HFC00) \ &H400
    Select Case gv15E0.M2CAB And &HF
    Case mc0132, mc0134

LA640:
      Select Case gv15E0.M2CC3
      Case 0
      Case 1
        Get m0026, , m0066
      Case 2
        Get m0026, , m0066
        Get m0026, , m0068
      Case 3
        Get m0026, , m0066
        Get m0026, , m0068
        Get m0026, , m006A
      Case Else
        sub01F8
      End Select
    Case mc0136
      sub06EA
      m0026 = 0
      Exit Sub
    Case mc0138
      m001E = False
      If m00C4 And gv0BD8 = 0 Then
        sub05E8 " ?"
        gv0BD8 = -2
      End If
      sub06EA
      If m00C4 Then sub06F5
      m0062 = 0
      gv0BDA = Loc(m0026) - 2
      m0078 = 0: m007A = False
      If gv15EA Then GoTo LA640
    Case mc013C
      Get m0026, , l011E
      Get m0026, , l011A
      If l011E > 2 Then
        Get m0026, , l011A
      Else
        l011A = 0
      End If
      l0114 = Space$(l011A)
      If l011A Then Get m0026, , l0114
      sub0704 l0114
      sub05E0 l0114
      l011A = l011E - l011A - 4
      If l011A > 0 Then
        Seek m0026, Seek(m0026) + l011A
      End If
    Case mc013E: Get m0026, , m0032
    Case mc0140: Get m0026, , m0036
    Case mc0142:    Get m0026, , m003A
    Case mc0144: Get m0026, , m0042
    Case mc0146
      Get m0026, , l011A: l011A = l011A \ 2
      m007C(0) = l011A
      For l0118 = 1 To l011A
        Get m0026, , m007C(l0118)
      Next l0118
    Case Else
      gv7B34 = gv0988
      sub022F mc014C
      gv0BD8 = gv15EA
      If gv15E0.M2CC3 Then
        Seek m0026, Seek(m0026) + gv15E0.M2CC3 * 2
      End If
    End Select
    Get m0026, , m006E
    Select Case gv15E0.M2CB3 And &HF
    Case mc014E: sub0625 gv15EA, l0122
    Case mc0150
      If gv15E0.M2CB3 And mc0152 Then m0064 = 0
      sub069E fn0BDF(), gv15E0.M2CCB
    Case mc0156
      sub069E fn0766(fn0BDF()), m00C4
    Case mc0158: sub05E0 fn0BDF()
    Case mc015A:  sub06B3 fn0BDF()
    Case mc015C: sub0693 fn0BDF()
    Case mc015E:  sub06AA fn0BDF()
    Case mc0160: sub0664 fn0BDF(), l0122, gv15E0.M2CCB
    Case mc0162
      sub0664 fn0766(fn0BDF()), l0122, gv15E0.M2CCB
    Case mc0164
      If gv15E0.M2CB3 And mc0166 Then l0122 = 0
    Case m0168:
      gv0BD8 = gv15EA
    Case Else
      sub01F8
    End Select
    If gv15E0.M2CB3 And mc016A Then
      If gv0BD8 = 0 Then gv0BD8 = gv15EA
    End If
    If gv15E0.M2CB3 And mc0152 Then
      m0064 = 0: m0062 = 0: m001E = False
      sub064E 0
    End If
  Wend
  m0026 = 0
Exit Sub

LACEE:
  sub01F8
  sub05E8 " ?"
  sub06EA
  If gv0C52 Then Print #gv0C52, "' --- unknown "; Hex$(gv15EA); " @ "; Hex$(gv0BDA); ":"; Hex$(Loc(m0026) - 2); " ---"
    While (fn0A66() And mc0170) = 0: Wend
    m006E = gv2382
    gv1140 = 0
  On Error GoTo LACEE
  If Seek(m0026) < l0116 Then Resume LA5A2
  m0026 = 0
  Exit Sub
End Sub

Sub sub06EA()
  If gv0C5E < 0 Then gv0C5E = 0: gv0C5A = "": Exit Sub
  gv0C5A = RTrim$(gv0C5A)
  While m00C4 > 0
    gv0C5A = gv0C5A & " ?+" & fn0772() & "?"
  Wend
  If m001C And Len(gv0C5A) = 0 Then Exit Sub
  If gv0C52 Then
    Print #gv0C52, gv0C5A
  Else
    On Error Resume Next
    gv0C56.AddItem gv0C5A
    If Err Then
      If Len(gv0C5A) = 0 Then Exit Sub
      gv0988 = gv0988 + 1
      If gv0988 >= UBound(gv0956) Then ReDim Preserve gv0956(gv0988 + 1)
      gv0956(gv0988).M1E6C = gv0C5A
      gv0956(gv0988).M1E3E = gv0BDA
      If m0026 Then gv0956(gv0988 + 1).M1E3E = Seek(m0026)
    End If
  End If
  gv0C5A = ""
End Sub

Sub sub06F5()
  If m00C4 = 0 Then Exit Sub
  If gv0C52 Then
    Print #gv0C52, "--- unused expressions (top down) ---"
    While m00C4 > 0
      Print #gv0C52, , fn0772()
    Wend
  End If
End Sub

Sub sub0704(p0194 As String)
Dim l0196 As Integer
  If Len(p0194) Then
    If InStr(p0194, Chr$(&HD)) Then sub01F8
    l0196 = -1
    Do
      l0196 = InStr(l0196 + 2, p0194, """")
      If l0196 = 0 Then Exit Do
      p0194 = Left$(p0194, l0196) & Mid$(p0194, l0196)
    Loop
  End If
  p0194 = """" & p0194 & """"
End Sub

Sub sub0715()
End Sub

Sub sub0723()
End Sub

Sub sub0730()
End Sub

Sub sub073D(p0198 As Integer)
Dim l019A As Integer
Dim l019C As String
  l019C = Space$(p0198)
  If p0198 > 0 Then
    Get m0026, , l019C
    If p0198 And 1 Then Get m0026, , gv0A30
    l019A = InStr(l019C, Chr$(&HD))
    While l019A
      l019C = Left$(l019C, l019A - 1) & String$(Asc(Mid$(l019C, l019A + 1, 1)), Mid$(l019C, l019A + 2, 1)) & Mid$(l019C, l019A + 3)
      l019A = InStr(l019C, Chr$(&HD))
    Wend
  End If
  sub05E0 fn0BDF() & l019C
End Sub

Sub sub0745()
  ReDim gv0956(10), m0092(50)
  gv0BC8 = True
  gv0988 = 0
End Sub

Sub sub0795(ByVal pv01F0 As Integer, p01F2 As String, p01F4 As Integer)
Dim l01F6 As Long
Dim l01F8 As Long
Dim l01FA As Integer
Dim l01FC As String
  l01FA = Asc(Mid$(p01F2, pv01F0, 1)) And &HF
  If l01FA >= mc00FA Then Exit Sub
  If p01F4 Then
    l01F6 = gv0B96 + 2 * CLng(pv01F0)
    sub05FF "Global Const " & fn0CDA(pv01F0) & " ="
  Else
    l01F6 = gv0BA2 + 2 * CLng(pv01F0)
    sub05FF "Const " & fn0DCA(pv01F0) & " ="
  End If
  l01F8 = Seek(gv088E)
  Seek gv088E, l01F6
  Select Case l01FA
  Case 0:
    Get gv088E, , l01FA
    sub05FF Str$(l01FA) & " ' ? &H" & fn02AC(l01FA)
    Do
      pv01F0 = pv01F0 + 1
      If Asc(Mid$(p01F2, pv01F0, 1)) Then Exit Do
      Get gv088E, , l01FA
      sub05FF fn02AC(l01FA)
    Loop
  Case 1: Get gv088E, , l01FA: sub05FF Str$(l01FA) & " ' &H" & Hex$(l01FA) & "%"
  Case 2: Get gv088E, , m0032: sub05FF Str$(m0032) & "& ' &H" & Hex$(m0032) & "&"
  Case 3: Get gv088E, , m0036: sub05FF Str$(m0036) & "!"
  Case 4:
    Get gv088E, , m003A: l01FC = Str$(m003A)
    If InStr(l01FC, ".") = 0 Then l01FC = l01FC & "#"
    sub05FF l01FC
  Case 5: Get gv088E, , m0042: sub05FF Str$(m0042) & "@"
  Case 6: sub022F mc00E0
  Case 7
    Get gv088E, , l01FA
    Get gv088E, gv0B92 + l01FA, l01FA
    Get gv088E, gv0B92 + l01FA, l01FA
    l01FC = Space$(l01FA)
    If l01FA > 0 Then Get gv088E, , l01FC
    sub0704 l01FC
    sub05FF " " & l01FC
  Case mc00FA: sub01F8
  Case mc0104
    Get gv088E, l01F6 - 2, l01FA
    sub05FF " As " & gv0FBC(l01FA And &HFF).M1D68
  Case mc00F4: sub01F8
  Case Else
    sub01F8
    Do
      Get gv088E, , l01FA: sub05FF fn02AC(l01FA)
      pv01F0 = pv01F0 + 1
    Loop Until Asc(Mid$(p01F2, pv01F0, 1))
  End Select
  Seek gv088E, l01F8
End Sub

Sub sub07A3(p0208 As Integer)
Dim l020C As T14DB
Dim l020E As String
  Get gv088E, fn0E71(gv7ABC(p0208).M24E0), l020C
  If l020C.M1511 And 1 Then
    gv0BDE = "Sub"
  Else
    gv0BDE = "Function"
  End If
  If l020C.M1519 <> &HC Then sub01F8
  l020E = fn0D70(p0208)
  gv0C4A.M223F = fn03D3(l020C.M1577)
  gv0C4A.M2246 = fn03D3(l020C.M158D)
  sub05FF "Declare " & gv0BDE & " " & l020E
  If Right$(gv0C4A.M223F, 1) = "." Then
    gv0C4A.M223F = Left$(gv0C4A.M223F, Len(gv0C4A.M223F) - 1)
  End If
  sub05FF " Lib """ & gv0C4A.M223F & """"
  If gv0C4A.M2246 <> l020E Then
    sub05FF " Alias """ & gv0C4A.M2246 & """"
  End If
  sub0607 p0208
End Sub

Sub sub07B0(ByVal pv021E As Integer, p0220() As Integer, p0226 As String, p0228 As Integer)
Dim l022A As Integer
Dim l022C As Integer
' Const mc022E = 224 ' &HE0%
' Const mc0230 = 192 ' &HC0%
' Const mc0232 = 16 ' &H10%
' Const mc0234 = 17 ' &H11%
  l022C = Asc(Mid$(p0226, pv021E, 1))
  If (l022C And mc022E) = mc0230 Then sub0795 pv021E, p0226, p0228: Exit Sub
  If p0228 Then
    sub05FF "Global " & fn0CDA(pv021E)
  Else
    sub05FF fn0DCA(pv021E)
  End If
  If l022C And mc0232 Then
    l022A = Asc(Mid$(p0226, pv021E + 1, 1))
    Select Case l022A
    Case mc0234
      sub07BC p0220(), pv021E + 3
    Case Else
      sub01F8
    End Select
  End If
  sub07C9 l022C, p0220(pv021E - 1), True
End Sub

Sub sub07F9()
  sub022F mc00E0
End Sub

