Attribute VB_Name = "MODULE4"
' Module4
Option Explicit
Dim m000E As T148E
Dim m0018 As String
Dim m001E As T14B9
Dim m0026 As Integer
Const mc004E = 64 ' &H40%
Const mc0054 = 32 ' &H20%
Const mc0066 = 4 ' &H4%
Const mc007A = 70 ' &H46%
Const mc0086 = 128 ' &H80%
Const mc00FE = 67 ' &H43%
Const mc0104 = 88 ' &H58%
Const mc0110 = 64 ' &H40%
Const mc0112 = 8 ' &H8%
Const mc012E = 1 ' &H1%
Const mc0150 = 36 ' &H24%
Const mc016E = 229 ' &HE5%
Const mc018A = 1 ' &H1%
Const mc0192 = -2 ' &HFFFE%
Const mc0196 = 224 ' &HE0%
Const mc0198 = 9 ' &H9%
Const mc01D8 = 1 ' &H1%
Const mc01DA = 10 ' &HA%

Function fn03D3(p00AE As Integer) As String
Dim l00B0 As String
Dim l00B4 As T221B
  Get gv088E, gv0B9A + fn0335(p00AE), l00B4
  l00B0 = Space$(Asc(l00B4.M222D))
  Get gv088E, , l00B0
  fn03D3 = l00B0
End Function

Function fn03DF(p00BA As Integer) As String
Dim l00bc As String
  l00bc = fn034F(gv0BA6, p00BA - 1)
  If InStr(l00bc, " ") Or InStr(gc1576, Right$(l00bc, 1)) Then
    l00bc = "[" & l00bc & "]"
  End If
  fn03DF = l00bc
End Function

Sub sub0385(p019A As Integer)
Dim l019C As String
Dim l019E As String
Dim l01A0 As Integer
Dim l01A2 As Integer
Dim l01A6 As String * 2
Dim l01A8 As Integer
Dim l01AA  As Integer
Dim l01AC As Integer
Dim l01AE As Integer
Dim l01B0 As String
Dim l01B2 As Integer
Dim l01B8 As Long
Dim l01BA As Long
Dim l01BC As Integer
Dim l01BE As String
Dim l01C0 As Integer
Dim l01C2 As String
Dim l01C4 As Long
Dim l01C6 As Long
' Const mc01D8 = 1 ' &H1%
' Const mc01DA = 10 ' &HA%
  Get gv088E, , m001E
  l01B8 = m001E.M14C8 And &H80000000
  l01BA = m001E.M14C8 Xor l01B8
  l01BA = l01BA - Len(m001E) - 1
  l01C6 = Seek(gv088E) + l01BA + 1
  If l01B8 Then
    Get gv088E, , l01A6
    l01BA = l01BA - 2
  End If
  Get gv088E, , gv0A30: l01BC = Asc(gv0A30)
  If l01BC Then
    sub01F8
    l01C2 = Space$(l01BC)
    Get gv088E, , l01C2
    l01BA = l01BA - l01BC
  Else
    l01A2 = Asc(m001E.M14D1)
    l01C2 = gv0FF6(gv0B72 + l01A2).M1D68
    m001E.M14C8 = m001E.M14C8 + Len(l01C2)
  End If
  l01BE = Chr$(Len(l01C2)) & l01C2
  l01C4 = Seek(gv088E)
  Get gv088E, , gv0A30
  l01C0 = Asc(gv0A30)
  If l01C0 = &HFF Then
    Get gv088E, , gv0A30: l01A0 = Asc(gv0A30)
    l019E = Space$(l01A0)
    Get gv088E, , l019E
    For l01C0 = 1 To gv141E
      If l019E = gv13EC(l01C0).M29AC Then Exit For
    Next l01C0
    If l01C0 > gv141E Then
      gv7B4A = l019E
      sub022F mc0110
      l01C0 = 0
    End If
  Else
    l01C0 = l01C0 + 1
    l019E = gv13EC(l01C0).M29AC
  End If
  If l01B8 Then l01BE = l01A6 & l01BE
  If p019A = 0 Then sub0394 gv0B72 + l01A2, l01C0, l01B8 <> 0
  Get gv088E, l01C4 + l01BA - 2, l01AA
  If l01AA <> &HFF Then
    l01AC = gv13EC(l01C0).M2BBD + 1
    Get gv088E, l01C4 + l01BA - 2 * l01AC - 2, l01AA
    If l01AA = (l01AC * &H100 Or &HFF) Then
      l01BA = l01BA - 2 * l01AC
    Else
      l01AC = &HFF
      Do
        l01BA = l01BA - 2
        l01AC = l01AC + &H100
        Get gv088E, l01C4 + l01BA - 2, l01AA
      Loop Until l01AA = l01AC
      l01AC = l01AA \ &H100
    End If
  End If
  m001E.M14C8 = m001E.M14C8 - 2 * l01AC - 1
  If p019A Then
    Put m0026, , m001E
    Put m0026, , l01BE
    Seek gv088E, l01C4
    l019C = Space$(&H4000)
    While l01BA > Len(l019C)
      Get gv088E, , l019C: Put m0026, , l019C
      l01BA = l01BA - Len(l019C)
    Wend
    l019C = Space$(l01BA - 1)
    Get gv088E, , l019C: Put m0026, , l019C
    If Asc(Right$(l019C, 1)) <> &HFF Then sub01F8
    Get gv088E, , gv0A30
    If Asc(gv0A30) <> l01AC Then sub01F8
    If l01AC Then Seek gv088E, Seek(gv088E) + 2 * l01AC
    If Seek(gv088E) <> l01C6 Then sub01F8
  ElseIf l01AC Then
    For l01A8 = 0 To l01AC - 1
      Get gv088E, , l01AA
      If l01AA Then
        If l01C0 Then
          l01AE = gv13EC(l01C0).M2BB3 + l01A8
          l01AE = gv150E(l01AE)
          l01B0 = gv14DA(l01AE).M1D68
        Else
          l01B0 = "event" & fn02B5(l01A8)
        End If
        If l01A2 Then
          l01B0 = l01C2 & "_" & l01B0
        ElseIf l01C0 Then
          l01B0 = gv13EC(l01C0).M29AC & "_" & l01B0
        Else
          l01B0 = "ctrl?_" & l01B0
        End If
        If l01AA And 1 Then
          l01B2 = gv0DE6(fn0E69(l01AA) + mc01D8)
          sub03A5 l01B2, l01B0, mc01DA
          If l01C0 Then sub03B4 l01B2, l01AE, l01B8 <> 0
        Else
          sub01F8
        End If
      End If
    Next l01A8
  End If
  Seek gv088E, l01C6
End Sub

Sub sub03C1()
End Sub

Sub sub03EC()
Dim l00C4  As Integer
Dim l00C6 As Integer
Dim l00C8 As Integer
Dim l00CA As Integer
Dim l00CC As String
Dim l00CE As Integer
Dim l00D0 As Integer
Dim l00D2 As Long
Dim l00D6 As T1257
Dim l00DA As T129E
' Const mc00FE = 67 ' &H43%
Dim l0100 As Long
' Const mc0104 = 88 ' &H58%
' Const mc0110 = 64 ' &H40%
' Const mc0112 = 8 ' &H8%
Dim l0114() As Integer
  Get gv088E, gv0DDA, l00D6
  sub03FC
  gv0F72 = 0: gv0F70 = 0: gv0F74 = 0
  gv0F82 = 0
  If l00D6.M1277 = -1 Then
    gv0F88 = 0
  Else
    gv0F88 = l00D6.M1277 And -2
  End If
  gv0FD2 = 0
  gv0F76 = fn0357(l00D6.M127E)
  gv0F7A = fn08ED(l00D6.M1289)
  gv0F7E = fn08ED(l00D6.M1294)
  ReDim gv1042(0), gv1076(0)
  Do
    Get gv088E, , l00DA
    If (l00DA.M12AE) = 0 Then Exit Sub
    gv0F74 = gv0F74 + 1
   'string length (unicode 2 byte) changed invalid asignment
    LSet gv0F5A(gv0F74) = l00DA
    
      
    l00C6 = gv0F5A(gv0F74).M13DD And &HFF
    gv0F5A(gv0F74).M13DD = l00C6
    Select Case l00C6
    Case mc00FE, mc007A
      gv0F72 = gv0F72 + 1
      l00CC = fn08ED(l00DA.M12B8 + 0)
      If l00C6 = mc007A Then
        l00D2 = Seek(gv088E)
        sub040D l00DA.M12BD: l0100 = gv0DDA
        sub040D l00DA.M12BD + 1
        Seek gv088E, l00D2
        If gv0DD4 = 0 Or l0100 >= gv0DDA Then
          gv7AA2 = 0
        End If
        If Right$(l00CC, 4) <> ".FRM" Then
          gv7B4A = l00CC
          sub022F mc0086
        End If
        If gv0F82 = 0 Then
          gv0F82 = gv0F74
          gv0F86 = l00DA.M12E2
        End If
        If l00DA.M12C5 > gv0FD2 Then gv0FD2 = l00DA.M12C5
        If Asc(l00DA.M12D9) And 1 Then
          gv0FBC(l00DA.M12C5).M25C6 = &H14 + 1
        Else
          gv0FBC(l00DA.M12C5).M25C6 = &HD + 1
        End If
        gv0FBC(l00DA.M12C5).M25D0 = l00DA.M12BD
      Else
        gv0F70 = gv0F70 + 1
        sub0418
        sub0427 0, l00CC
      End If
      gv0F42(gv0F74) = l00CC
    Case mc0104
      l00CE = l00DA.M12B8
      If l00CE = &HFF Then
        Get gv088E, , gv0A30
        gv0F42(gv0F74) = fn08ED(Asc(gv0A30))
        For l00CE = gv141E To 1 Step -1
          If gv13EC(l00CE).M29AC = gv0F42(gv0F74) Then Exit For
        Next
        If l00CE = 0 Then
          gv7B4A = gv0F42(gv0F74)
          sub022F mc0110
          gv13EC(0).M29AC = "?"
        End If
      Else
        l00CE = l00CE + 1
        gv0F42(gv0F74) = gv13EC(l00CE).M29AC
      End If
      gv0F5A(gv0F74).M2535 = l00CE
      If gv0FBC(l00DA.M12C5).M25C6 Then
        If gv0FBC(l00DA.M12C5).M25C6 <> l00CE Then
          gv7B4A = gv0F42(gv0F74)
          sub022F mc0110
        End If
      Else
        gv0FBC(l00DA.M12C5).M25C6 = l00CE
        gv0FBC(l00DA.M12C5).M1D68 = gv13EC(l00CE).M29AC
      End If
      If l00DA.M12CF <> &H1FF Then
        gv0FBC(l00DA.M12CF).M25C6 = -l00CE
        gv0FBC(l00DA.M12CF).M1D68 = gv13EC(l00CE).M29AC
      End If
    Case Else
      sub022F mc0112
      Exit Do
    End Select
    gv0F5A(gv0F74).M13DD = l00C6
    Select Case l00C6
    Case mc007A, mc0104
      Get gv088E, , l00CA
      If l00CA > 0 Then
        ReDim l0114%(l00CA)
        For l00C8 = 1 To l00CA
          Get gv088E, , l0114(l00C8)
        Next
        For l00C8 = 1 To l00CA
          l00D0 = l0114(l00C8)
          Get gv088E, , l00C4
          If l00C6 = mc007A Then
            If l00D0 > UBound(gv1076) Then ReDim Preserve gv1076(l00D0)
            If gv1076(l00D0).M2635 = 0 Then
              gv1076(l00D0).M263C = l00C4
              gv1076(l00D0).M2635 = l00DA.M12C5
            End If
          Else
            If l00D0 > UBound(gv1042) Then ReDim Preserve gv1042(l00D0)
            If gv1042(l00D0) = 0 Then
              gv1042(l00D0) = gv1476(gv13EC(l00CE).M2B94 + l00C4)
            End If
          End If
        Next
      End If
    End Select
  Loop
End Sub

Sub sub0435()
Dim l0120 As Integer
Dim l0122 As Integer
Dim l0124 As Integer
Dim l0126 As Integer
Dim l0128 As Integer
' Const mc012E = 1 ' &H1%
  sub0201 l0128, gc03AE
  For l0120 = 1 To gv0F72
    If gv0F5A(l0120).M13DD <> mc00FE Then
      l0122 = gv0F5A(l0120).M12E2
      l0126 = fn0E69(l0122)
      l0124 = gv0DE6(l0126 + mc012E)
      sub0441 l0124, (gv0F42(l0120))
      gv7A70(l0124).M24B4 = gv0F5A(l0120).M13DD
      gv7A70(l0124).M24BE = Asc(gv0F5A(l0120).M12D9)
      gv7A70(l0124).M12BD = gv0F5A(l0120).M12BD
      gv7A70(l0124).M12C5 = gv0F5A(l0120).M12C5
      gv0F5A(l0120).M252C = l0124
    Else
      frmMain.control8.Text = gv0F42(l0120)
    End If
  Next l0120
  sub0220 l0128
End Sub

Sub sub0450()
Dim l0142 As Long
Dim l0144 As Integer
Dim l0146 As Integer
Dim l0148  As Integer
' Const mc0150 = 36 ' &H24%
' Const mc016E = 229 ' &HE5%
  sub0201 l0148, gc03D6 & gc0402 & gc03FA
  gv0DD4 = 0
  sub0460 gv0DC0 + 1
  gv0B8E = gv0DDA
  Get gv088E, gv0B8E + mc0150, gv0B70
  l0142 = &H60
  Get gv088E, gv0B8E + l0142, l0146: l0142 = l0142 + 2
  gv0B92 = gv0B8E + l0142
  l0142 = l0142 + fn0335(l0146)
  Get gv088E, gv0B8E + l0142, l0144: l0142 = l0142 + 2
  Get gv088E, gv0B8E + l0142, l0146: l0142 = l0142 + 2
  gv0B96 = gv0B8E + l0142
  gv09BE = fn0335(l0146)
  gv09C2 = (gv09BE \ 2) - 1
  ReDim gv098C(0 To gv09C2), gv78A8(0 To gv09C2)
  ReDim gv7876(0)
  ReDim gv790A(0)
  gv78DA = String$(gv09C2, 0) & Chr$(mc016E)
  For l0144 = 0 To gv09C2
    Get gv088E, , gv098C(l0144)
  Next
  l0142 = l0142 + gv09BE
  Get gv088E, gv0B8E + l0142, l0146: l0142 = l0142 + 2
  l0142 = l0142 + fn0335(l0146)
  Get gv088E, gv0B8E + l0142, l0146: l0142 = l0142 + 2
  gv0B9A = gv0B8E + l0142
  l0142 = l0142 + fn0335(l0146)
  gv0B9E = l0142
  sub0220 l0148
End Sub

Sub sub046C()
Dim l0172 As Long
Dim l0174 As Integer
Dim l0178 As Integer
Dim l017A  As Integer
Dim l017E  As Integer
Dim l00A2 As String
' Const mc018A = 1 ' &H1%
  sub0201 l017E, gc03D6 & gc03FE & gc03F6
  If gv0B9E = 0 Then sub01F8
  sub0460 gv0DC0 + 1
  If gv091A Then
    sub047C l017E, gc03D6 & gc03EE
    sub048D
    sub047C l017E, gc03D6 & gc03E2
  End If
  l0172 = gv0B9E
  Do
    Get gv088E, gv0B8E + l0172, l017A: l0172 = l0172 + 2
    If l017A = 0 Or l017A = -1 Then Exit Do
    l0178 = fn0E69(l017A)
    If gv0DE6(l0178) <> mc018A Then sub01F8
    l017A = gv0DE6(l0178 + mc012E)
    gv7A70(l017A).M24C8 = l0172
    If gv091A Then
    'decompiling inheritance when decompiling over exiting files
    l00A2 = fn0FBE("", Hex$(l017A))   'filename to var write
    On Error Resume Next
    If FileLen(l00A2) > 0 Then Kill (l00A2)
    ' end inheritance killing
    On Error GoTo 0
        sub01A8 l017A
        End If
    Get gv088E, gv0B8E + l0172, l0174: l0172 = l0172 + 2
    l0172 = l0172 + fn0335(l0174)
    If gv7A70(l017A).M24B4 = mc007A Then
      Get gv088E, gv0B8E + l0172, l0174: l0172 = l0172 + 2
      l0172 = l0172 + l0174
      Get gv088E, gv0B8E + l0172, l0174: l0172 = l0172 + 2
      l0172 = l0172 + l0174
    Else
      Get gv088E, gv0B8E + l0172, l0174: l0172 = l0172 + 2
      l0172 = l0172 + l0174
    End If
  Loop
  sub0220 l017E
End Sub

Sub sub048D()
Dim l018C As Integer
Dim l018E As Integer
' Const mc0192 =-2 ' &HFFFE%
' Const mc0196 = 224 ' &HE0%
' Const mc0198 = 9 ' &H9%
  gv2382 = mc0192   'scaning types
  If gv09BE < 6 Then sub01F8: Exit Sub
  l018C = gv098C(2)
  Do While l018C
    sub04AA l018C
  Loop
  For l018E = gv09C2 - 2 To 0 Step -3
    If gv098C(l018E + 1) Or gv098C(l018E + 2) Then Exit For
    l018C = gv098C(l018E) Xor &H8000
    If l018C <= 0 Or l018C > gv0FD2 Then Exit For
    sub04B5 l018E + 1, 2, mc0196, mc0198, 0
    sub04C3 gv0FBC(l018C).M1D68, l018E + 1
  Next l018E
  sub04D2
End Sub

Sub sub049B()
  gv0DD4 = 0
  sub0460 gv0DC0
  Get gv088E, gv0DDA, gv0A30
  Select Case iVBVersion
  Case 1 To 3
    sub03EC
  Case Else
    sub022F mc0112
    iVBVersion = 0
  End Select
End Sub

Sub sub04E2(p0056 As Integer)
Dim l0058 As String
Dim l005A As Integer
Dim l005C As Integer
Dim l005E As Integer
Dim l0060  As Integer
Dim l0062  As Integer
' Const mc0066 = 4 ' &H4%
' Const mc007A = 70 ' &H46%
' Const mc0086 = 128 ' &H80%
  If iVBVersion <> 2 And iVBVersion <> 3 Then
    sub022F mc0066
    Exit Sub
  End If
  If p0056 Then
    sub0201 l0060, gc03DA & "*.FRM"
    sub0201 l0062, gc03D2 & "empty.frm"
    m0026 = FreeFile
    Open gv0BBA & "empty.frm" For Binary Access Read As m0026
    Get m0026, 1, m000E
    m0018 = Space$(LOF(m0026) - m000E.M14B1 - 5)
    Get m0026, m000E.M14B1 + 6, m0018
    Close m0026
    If iVBVersion = 2 Then
      Mid$(m000E.M149E, 3, 1) = Chr$(&H2B)
    End If
    sub0220 l0062
  Else
    sub0201 l0060, gc03D6 & "*.FRM"
  End If
  For l005A = 1 To gv0F72
    If gv0F5A(l005A).M13DD = mc007A Then
      sub0201 l0062, gc03D2 & "RC" & Hex$(gv0F5A(l005A).M12BD)
      sub040D gv0F5A(l005A).M12BD
      If gv0DD4 = 0 Then
        gv7B4A = gv0F42(l005A)
        sub022F mc0086
      Else
        Get gv088E, gv0DDA, m000E
        l005E = Asc(m000E.M14A9)
        If p0056 Then
          gv0B72 = gv0FBC(gv0F5A(l005A).M12C5).M25D0
          l0058 = gv7B8C & gv0F42(l005A)
          sub047C l0062, gc03DA & gv0F42(l005A)
          Open l0058 For Binary As m0026
          Put m0026, , m000E
        Else
          sub047C l0062, gc03D6 & gv0F42(l005A)
          gv1028 = gv1028 + 1
          gv0B72 = gv1028
          ReDim Preserve gv0FF6(LBound(gv0FF6) To gv1028 + l005E)
          If gv0FBC(gv0F5A(l005A).M12C5).M25D0 <> gv0F5A(l005A).M12BD Then
            sub01F8
          End If
          gv0FBC(gv0F5A(l005A).M12C5).M25D0 = gv1028
          sub040D gv0F5A(l005A).M12BD + 1
          If gv0DD4 = 0 Then
            sub022F gc039A
            
          End If
          For l005C = 0 To l005E
            Get gv088E, , gv0A30
            If Asc(gv0A30) = 0 Then
              l0058 = "control" & Format$(l005C)
            Else
              l0058 = Space$(Asc(gv0A30))
              Get gv088E, , l0058
            End If
            gv0FF6(gv1028).M1D68 = l0058
            gv1028 = gv1028 + 1
          Next l005C
          gv0FBC(gv0F5A(l005A).M12C5).M1D68 = gv0FF6(gv0B72).M1D68
          sub040D gv0F5A(l005A).M12BD
        End If
        Get gv088E, gv0DDA, m000E
        sub0385 p0056
        Get gv088E, , gv0A30
        If p0056 Then Put m0026, , gv0A30
        If Asc(gv0A30) = 1 Then
          sub0385 p0056
          Do
            Get gv088E, , gv0A30
            If p0056 Then Put m0026, , gv0A30
            Select Case Asc(gv0A30)
            Case 1, 3
              sub0385 p0056
            Case 2
            Case 4, 5
              Exit Do
            Case Else
              gv7B4A = gv0F42(l005A)
              sub022F mc0086
              Exit Do
            End Select
          Loop
        End If
        If Asc(gv0A30) = 5 Then
          sub0385 p0056
          Do
            Get gv088E, , gv0A30
            If p0056 Then Put m0026, , gv0A30
            Select Case Asc(gv0A30)
            Case 2
              sub0385 p0056
            Case 3
            Case 4
              Exit Do
            Case Else
              gv7B4A = gv0F42(l005A)
              sub022F mc0086
              Exit Do
            End Select
          Loop
        End If
        If Asc(gv0A30) <> 4 Then
          sub022F mc0086
        End If
        If p0056 Then
          m000E.M14B1 = Seek(m0026) - 6
          Put m0026, , m0018
          Put m0026, 1, m000E
          Close m0026
        End If
      End If
      sub0220 l0062
    End If
  Next l005A
  sub0220 l0060
End Sub

Sub sub04F2()
Dim l0036 As Integer
Dim l0038 As String
Dim l003A As Integer
Dim l003C As Integer
Dim l003E As String
Dim l0040 As String
Dim l0044  As Integer
Dim l0046  As Integer
' Const mc004E = 64 ' &H40%
' Const mc0054 = 32 ' &H20%
  For l0036 = 1 To gv7AA2
    l0038 = fn0D20(l0036)
    If Right$(l0038, 4) = ".FRM" Then
      l0040 = fn0FBE(l0038, "FRM")
      l003E = fn0FBE(l0038, "BAS")
      If Len(Dir$(l003E)) Then
        l003A = FreeFile
        sub0201 l0044, l0040
        Open l0040 For Binary As l003A
        l0038 = Input$(7, l003A)
        If l0038 = "VERSION" Then
          Close l003A
          Open l0040 For Append As l003A
          l003C = FreeFile
          sub0201 l0046, " + " & l003E
          Open l003E For Input As l003C
          While Not EOF(l003C)
            Line Input #l003C, l0038
            Print #l003A, l0038
          Wend
          Close l003C
          Kill l003E
        Else
          MsgBox l0040, mc004E, gc043E
          gv7A1C = gv7A1C And Not mc0054
        End If
        Close l003A
        sub0220 l0046
        sub0220 l0044
      End If
    End If
  Next
End Sub

