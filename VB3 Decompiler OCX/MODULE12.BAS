Attribute VB_Name = "Module12"
' Module12
Option Explicit
Const mc000E = 3 ' &H3%
Const mc0010 = 4 ' &H4%
Const mc003C = 17 ' &H11%
Const mc003E = 18 ' &H12%
Const mc0040 = 19 ' &H13%
Const mc0042 = 20 ' &H14%
Const mc0044 = 21 ' &H15%
Const mc004A = 31 ' &H1F%
Dim m004C As Integer
Dim m008C As Long
Dim m0090 As Long
Const mc0096 = 1 ' &H1%
Const mc0098 = 2 ' &H2%
Const mc009A = 3 ' &H3%
Dim m009E As T2C2F
Dim m00A6(3) As String
Const mc00F2 = 1 ' &H1%
Const mc011A = 16384 ' &H4000%
Const mc015A = 4096 ' &H1000%
Const mc015E = 64 ' &H40%
Const mc0162 = 4 ' &H4%

Sub sub0427(p012E As Integer, ByVal p0130 As String)
Dim l0132 As Integer
Dim l0134 As Integer
Dim l0136 As Integer
Dim l0138 As Integer
Dim l013A As Integer
Dim l013C As Integer
Dim l013E As Long
Dim l0142 As T2B7A
Dim l0144  As Integer
Dim l0148  As Integer
Dim l014A  As Integer
Dim l014C  As Integer
Dim l014E As String
' Const mc015A = 4096 ' &H1000%
' Const mc015E = 64 ' &H40%
' Const mc0162 = 4 ' &H4%
  l013A = fn02FF(p0130, ".")
  If l013A = 0 Then
    p0130 = p0130 & ".300"
  Else
    p0130 = Left$(p0130, l013A) & "300"
  End If
  If p012E Then sub047C p012E, gc03D2 & p0130
On Error GoTo L15B1C

L154F8:
  l014E = gv0886
  gv7B4A = l014E & p0130
  l013A = Len(Dir$(gv7B4A))
  If l013A = 0 Then
    l014E = gv0BBA
    gv7B4A = l014E & p0130
    l013A = Len(Dir$(gv7B4A))
  End If
  If l013A = 0 Then Error 1000
  l013E = FileLen(gv7B4A)
  l0132 = FreeFile
  Open gv7B4A For Binary Access Read As l0132
  Get l0132, , l0144
  Get l0132, , l013C
  If l0144 = mc0010 Then
    l013E = 0
    If l013E = 0 Then
      sub022F mc011A
      Error 1002
    End If
  ElseIf l0144 <> mc000E Then
    Error 1001
  End If
  ReDim Preserve gv13EC(gv141E + l013C)
  For l0134 = 1 To l013C
    If l0144 = mc000E Then
      Get l0132, , l0142
    Else
      Get l0132, , l014A
      Get l0132, , l014C
      sub0B15 l0148, l014A, l014C, l0142
    End If
    If gv7A2A Then
      If p012E Then sub047C p012E, gc03D2 & l0142.M29AC
      frmMain.control8.AddItem l0142.M29AC
    End If
    If l0142.M2B9D >= 0 Then
      If l0144 = mc000E Then
        l0142.M2B94 = gv14A8
        ReDim Preserve gv1476(gv14A8 + l0142.M2B9D + 1)
        ReDim Preserve gv1442(LBound(gv1442) To gv1474 + l0142.M2B9D + 1)
      End If
      For l0138 = 0 To l0142.M2B9D
        If l0144 = mc000E Then
          Get l0132, , l013A
          If l013A > 0 Then
            gv1474 = gv1474 + 1
            Get l0132, , gv1442(gv1474)
            l013A = gv1474
          End If
          gv1476(gv14A8) = l013A
          gv14A8 = gv14A8 + 1
        Else
          l013A = gv1476(l0142.M2B94 + l0138)
        End If
        If l013A > 0 Then
          If (gv1442(l013A).M13DD And &HFF) = &H70 Then
            Select Case gv1442(l013A).M1D68
            Case "Fields": gv1442(l013A).M13DD = mc0040 Or gc129A
            Case "TableDefs": gv1442(l013A).M13DD = mc0042 Or gc129A
            Case "Indexes": gv1442(l013A).M13DD = mc0044 Or gc129A
            Case Else: sub022F mc015A
            End Select
          End If
        End If
      Next l0138
    Else
      l0142.M2B94 = 0
    End If
    If l0142.M2BBD >= 0 Then
      If l0144 = mc000E Then
        l0142.M2BB3 = gv1540
        ReDim Preserve gv150E(gv1540 + l0142.M2BBD + 1)
        ReDim Preserve gv14DA(LBound(gv14DA) To gv150C + l0142.M2BBD + 1)
        For l0136 = 0 To l0142.M2BBD
          Get l0132, , l013A
          If l013A > 0 Then
            gv150C = gv150C + 1
            l013A = gv150C
            Get l0132, , gv14DA(gv150C)
          End If
          gv150E(gv1540) = l013A
          gv1540 = gv1540 + 1
        Next l0136
      End If
    Else
      l0142.M2BB3 = 0
    End If
    gv141E = gv141E + 1
    gv13EC(gv141E) = l0142
  Next l0134

L15AEA:
  If l0132 Then Close l0132
  If l0148 Then Close l0148
Exit Sub

L15B1C:
  If l0132 Then Close l0132: l0132 = 0
  If l0148 Then Close l0148: l0148 = 0
  gv7B30 = gv7B30 And Not mc015E
  sub022F mc015E
  If gv7B48 = mc0162 Then Resume L154F8
  Resume L15AEA
End Sub

Sub sub064E(ByVal pv00DA As Integer)
  Select Case pv00DA And &HC000
  Case 0, &H8000
    pv00DA = pv00DA And &HFF
    m009E.M13DD = mc0098
    m009E.M1D70 = pv00DA
    m009E.M1D68 = gv0FBC(pv00DA).M1D68
  Case &H4000
    If pv00DA = &H4006 Then
      m009E.M13DD = mc0098
      m009E.M1D70 = gv7AA6.M12C5
      m009E.M1D68 = "me"
    Else
      m009E.M13DD = mc0098
      pv00DA = pv00DA And &HFF
      m009E.M1D70 = pv00DA
      m009E.M1D68 = gv0FBC(pv00DA).M1D68
    End If
  Case Else
    Beep
    m009E.M13DD = 0
  End Select
End Sub

Sub sub0671()
Dim l012C As Integer
  Select Case m009E.M13DD
  Case mc0096
    l012C = gv1442(m009E.M1D70).M13DD
    If l012C And gc129A Then
      l012C = l012C Xor gc129A
    Else
    End If
  Case mc0098
    l012C = m009E.M1D70
    If l012C < 0 Then
      l012C = -l012C
    ElseIf l012C = 20 Then
      l012C = 0
    Else
      If gv0FBC(l012C).M25C6 < 0 Then
        If -gv0FBC(l012C).M25C6 = gv0FBC(l012C - 1).M25C6 Then
          l012C = l012C - 1
        Else
          l012C = -gv0FBC(l012C).M25C6
        End If
      Else
        l012C = gv0FBC(l012C).M25DA
        If l012C < 0 Then
          l012C = -l012C
        Else
          l012C = Abs(gv0FBC(l012C).M25DA)
        End If
      End If
      If l012C Then
        m009E.M1D68 = m009E.M1D68 & "<." & gv0FBC(l012C).M1D68 & ">"
      Else
        m009E.M1D68 = m009E.M1D68 & "<.defcoll>"
      End If
    End If
  Case mc009A
    l012C = m009E.M1D70
  End Select
  m009E.M1D68 = m009E.M1D68 & "()"
  If l012C Then
    m009E.M1D70 = l012C
  Else
    m009E.M13DD = 0
  End If
End Sub

Function fn0B57(p00C6 As Integer, p00C8 As Long, p00CA As Integer) As String
Dim l00CE As String * 256
Dim l00D0 As Integer
  Get p00C6, p00C8 + fn0335(p00CA), l00CE
  l00D0 = InStr(l00CE, Chr$(0))
  If l00D0 Then
    fn0B57 = Left$(l00CE, l00D0 - 1)
  Else
    MsgBox gc038A
    fn0B57 = l00CE
  End If
End Function

Function fn0B63(p00D8 As Integer) As String
  fn0B63 = gv12BC(gv1442(p00D8).M13DD And gc129C)
End Function

Function fn0B6F(p00EA As Integer) As String
Dim l00EC As Integer
Dim l00EE As Integer
Dim l00F0 As String
' Const mc00F2 = 1 ' &H1%
  l00EC = m009E.M1D70
  l00EE = p00EA And &H3FFF
  Select Case p00EA
  Case -2
    If l00EC Then
      If m009E.M13DD = mc0098 Then
        l00EC = gv0FBC(l00EC).M25DA
      Else
        l00EC = 0
      End If
    End If
    If l00EC Then
      m009E.M1D68 = m009E.M1D68 & "<." & gv0FBC(l00EC).M1D68 & ">"
      m009E.M1D70 = l00EC
    Else
      m009E.M1D68 = m009E.M1D68 & "<.hide2>"
      m009E.M13DD = 0
      sub01F8
    End If
    Exit Function
  Case &HC0FE
    l00F0 = "Controls"
    m009E.M13DD = mc0098
    l00EC = 5
  Case &HC0FD
    l00F0 = "Count"
    m009E.M13DD = 0
    l00EC = 0
  Case Else
    If l00EC = 0 Then
      sub022F mc00F2
    End If
    Select Case m009E.M13DD
    Case mc0098
      Select Case p00EA And &HC000
      Case &HC000
        l00EC = gv0FBC(l00EC).M25C6
L148A8:
        If l00EC > 0 Then
          If l00EE > gv13EC(l00EC).M2B9D Then
            l00EC = 0
          Else
            l00EC = gv13EC(l00EC).M2B94
            l00EC = gv1476(l00EC + l00EE)
            l00F0 = gv1442(l00EC).M1D68
          End If
          m009E.M13DD = mc0096
        End If
      Case &H8000
L1494C:
        l00EC = gv0FBC(l00EC).M25D0
        If l00EC = 0 Then
          sub01F8
        Else
          l00EC = l00EC + l00EE
          m009E.M13DD = mc009A
          l00F0 = gv0FF6(l00EC).M1D68
        End If
      Case 0
        Select Case l00EC
        Case 0
          sub01F8
        Case 1, 2, 4
          l00EC = 0
          If l00EE <= UBound(gv1076) Then
            l00EC = gv1076(l00EE).M2635
            If l00EC Then
              l00EE = gv1076(l00EE).M263C
              l00EC = gv0FBC(l00EC).M25D0
              l00F0 = gv0FF6(l00EC + l00EE).M1D68
              l00EC = 4
              m009E.M13DD = mc0098
            End If
          End If
          If l00EC = 0 Then
            If l00EE <= UBound(gv1042) Then
              l00EC = gv1042(l00EE)
              If l00EC Then
                m009E.M13DD = mc0096
                l00F0 = gv1442(l00EC).M1D68
              End If
            End If
          End If
          If l00EC = 0 Then
            sub01F8
          End If
        Case Else
          l00F0 = fn03DF(l00EE)
        End Select
      Case Else
        sub01F8
      End Select
    Case mc009A
      Select Case p00EA And &HC000
      Case &HC000
      Case Else
      End Select
      l00EC = Abs(gv0FF6(l00EC).M25C6)
      GoTo L148A8
    Case Else
      If l00EE <= UBound(gv1042) Then
        l00EC = gv1042(l00EE)
        m009E.M13DD = mc0096
        l00F0 = gv1442(l00EC).M1D68
      Else
        sub01F8
        l00F0 = fn03DF(p00EA)
        m009E.M13DD = 0
      End If
    End Select
  End Select
  m009E.M1D70 = l00EC
  If m009E.M13DD = mc0096 Then
    l00EC = gv1442(l00EC).M13DD
    Select Case l00EC And gc129C
    Case mc003C: l00EC = &HB
    Case mc003E: l00EC = &HE
    Case mc0040: l00EC = &HF
    Case mc0042: l00EC = &HC
    Case mc004A: l00EC = 1
    Case m004C: l00EC = 4
    Case gc12BA: l00EC = &H14
    Case mc0044: l00EC = &H10
    Case Else: l00EC = 0
    End Select
    If l00EC Then
      m009E.M1D70 = l00EC
      m009E.M13DD = mc0098
    End If
  End If
  If Len(l00F0) = 0 Then
    l00F0 = "P?" & Hex$(p00EA)
    m009E.M13DD = 0
  End If
  m009E.M1D68 = m009E.M1D68 & "." & l00F0
  fn0B6F = l00F0
End Function

Sub sub0B15(p1 As Integer, p2 As Integer, p3 As Integer, p4 As T2B7A)
' called (p1 As type, p2 As type, p3 As type, p4 As type)
End Sub

Sub sub0B24()
End Sub

Sub sub0B31(p010C As Integer)
Dim l010E As Integer
Dim l0110 As Integer
Dim l0112  As Integer
' Const mc011A = 16384 ' &H4000%
  gv1474 = 0: gv14A8 = 1
  ReDim gv1442(0), gv1476(0)
  gv150C = 0: gv1540 = 1
  ReDim gv14DA(0), gv150E(0)
  gv141E = 0
  ReDim gv13EC(0)
  l010E = FreeFile
  gv7B4A = gv0BBA & gc1208
  If p010C Then sub047C p010C, gc03D2 & gv7B4A
  Open gv7B4A For Binary Access Read As l010E
  Get l010E, , l0110
  If l0110 <> (mc000E Or &H8000) Then
    Close l010E
    sub022F mc011A
    'End
  End If
  Get l010E, , l0112
  For l0110 = 0 To l0112
    Get l010E, , l0112
    gv12D2(l0110) = Space$(l0112)
    Get l010E, , gv12D2(l0110)
  Next l0110
  Get l010E, , l0112
  For l0110 = 0 To l0112
    Get l010E, , l0112
    gv12BC(l0110) = Space$(l0112)
    Get l010E, , gv12BC(l0110)
  Next l0110
  Get l010E, , l0112: ReDim gv1442(-l0112 To 0)
  m008C = fn0335(-l0112) - 1
  For l0110 = 1 To l0112 + 1
    Get l010E, , gv1442(1 - l0110)
  Next l0110
  Get l010E, , l0112: ReDim gv14DA(-l0112 To 0)
  m0090 = fn0335(-l0112) - 1
  For l0110 = 1 To l0112 + 1
    Get l010E, , gv14DA(1 - l0110)
  Next
  Close l010E
  If gv12BC(mc0040) = "" Then
    gv12BC(mc0040) = "Field"
    gv12BC(mc0042) = "TableDef"
    gv12BC(mc0044) = "Index"
  End If
End Sub

Sub sub0B3E(p00BC As Integer)
  m00A6(0) = "<no/std obj>"
  m00A6(mc009A) = "<ctrl>"
  m00A6(mc0098) = "<class>"
  m00A6(mc0096) = "<prop>"
  sub0B31 p00BC
End Sub

Sub sub0B4B()
End Sub

