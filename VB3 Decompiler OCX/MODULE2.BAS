Attribute VB_Name = "Module2"
' Module2
Option Explicit
Const mc0006 = 10 ' &HA%
Const mc0008 = 8 ' &H8%
Const mc00C6 = 229 ' &HE5%
Const mc00CE = 16384 ' &H4000%
Const mc00DE = 32 ' &H20%

Sub sub015B()
Dim l00D2 As Integer
Dim l00D4 As Integer
Dim l00D6 As Integer
' Const mc00DE = 32 ' &H20%
  If gv7AB6 <= 0 Then gv091C = False
  If Not gv091C Then Exit Sub
  l00D4 = FreeFile
  Open fn0FBE("", Hex$(gv7AB6)) For Binary As l00D4
  l00D6 = mc0008: Put l00D4, 1, l00D6
  l00D6 = gv09FA: Put l00D4, , l00D6
  Put l00D4, , gv7952
  For l00D2 = 1 To gv09FA
    Put l00D4, , gv7988(l00D2)
  Next
  l00D6 = UBound(gv7956)
  Put l00D4, , l00D6
  For l00D2 = 1 To l00D6
    l00D6 = Len(gv7956(l00D2))
    Put l00D4, , l00D6
    If l00D6 Then Put l00D4, , gv7956(l00D2)
  Next l00D2
  For l00D2 = 1 To gv09FA
    Put l00D4, , gv09FC(l00D2)
  Next
  l00D2 = 0: Put l00D4, , l00D2
  Close l00D4
  gv091C = False
End Sub

Sub sub016C()
  If gv091A Then sub0178
  If gv091C Then sub015B
End Sub

Sub sub0178()
Dim l009A As Integer
Dim l009C As Integer
Dim l009E As Integer
  Screen.MousePointer = 11
  l009C = FreeFile
  Open fn0FBE("", "0") For Binary Access Write As l009C
  l009A = mc0006: Put l009C, , l009A
  Put l009C, , gv7A1C
  l009E = UBound(gv7876): Put l009C, , l009E
  For l009A = 1 To l009E
    l009E = Len(gv7876(l009A)): Put l009C, , l009E
    If l009E Then
      Put l009C, , gv7876(l009A)
    End If
  Next l009A
  l009E = Len(gv78DA): Put l009C, , l009E
  Put l009C, , gv78DA
  For l009A = 0 To gv09C2
    Put l009C, , gv78A8(l009A)
  Next
  l009E = UBound(gv790A): Put l009C, , l009E
  For l009A = 0 To l009E
    Put l009C, , gv790A(l009A)
  Next l009A
  l009E = UBound(gv1042)
  Put l009C, , l009E
  For l009A = 1 To l009E
    Put l009C, , gv1042(l009A)
  Next
  l009E = UBound(gv1076)
  Put l009C, , l009E
  For l009A = 1 To l009E
    Put l009C, , gv1076(l009A)
  Next
  For l009A = 1 To gv7AEE
    Put l009C, , gv7ABC(l009A)
  Next
  l009A = 0: Put l009C, , l009A
  Close l009C
  Screen.MousePointer = 0
  gv091A = False
End Sub

Sub sub0189()
  gv1140 = 0
  Do Until fn0A66() And mc00DE
  Loop
End Sub

Sub sub019A(p00E0 As Long, p00E2 As Long)
  If gv06B4 Then
    p00E0 = gv06B4 + fn0335(gv7AF2.M1538)
    p00E2 = gv06B4 + fn0335(gv7AF2.M1561)
  Else
    p00E0 = 0: p00E2 = 0
  End If
End Sub

Sub sub01A8(p003C As Integer)
  If gv091C Then sub015B
  gv7AA6 = gv7A70(p003C)
  gv7AB6 = p003C
  gv7AB8 = gv7AA6.M24AC
  sub01B7
End Sub

Sub sub01B7()
Dim l00A0 As Integer
Dim l00A2 As String
Dim l00A4 As Integer
Dim l00A6 As Integer
Dim l00A8 As Integer
Dim l00AA As Integer
' Const mc00C6 = 229 ' &HE5%
' Const mc00CE = 16384 ' &H4000%
Dim l00D0 As Integer
  If gv7AB6 <= 0 Then Exit Sub
  l00A2 = fn0D20(gv7AB6)
  If gv7A2A Then frmMain.control4.Value = 1
  gv0BA2 = gv0B8E + gv7AA6.M24C8
  Get gv088E, gv0BA2, l00A6: gv09F6 = fn0335(l00A6) 'size of local to global var map
  gv09FA = (gv09F6 \ 2) - 1
  gv0BA2 = gv0BA2 + 2
  If gv09FA < 0 Then gv09FA = 0  'map cnt
  ReDim gv7956(0), gv7988(gv09FA)
  gv7952 = String$(gv09FA, 0) & Chr$(mc00C6)
  ReDim gv09FC(gv09FA)
  gv091C = True  'map table
  If gv09F6 < 6 Then Exit Sub
  sub0201 l00AA, gc03D2 & l00A2  'loading ...
  ReDim gv09C4(gv09FA)
  For l00A0 = 0 To gv09FA
    Get gv088E, , gv09C4(l00A0)
  Next l00A0
  l00A4 = FreeFile
  l00A2 = fn0FBE("", Hex$(gv7AB6))   'filename to var write
On Error Resume Next
  'decompiling inheritance when decompiling over exiting files
 ' If FileLen(l00A2) > 0 Then Kill (l00A2)
  ' end inheritance killing
  If FileLen(l00A2) = 0 Then Error 1000
  If Err Then
    On Error GoTo 0
    sub0210
    sub0220 l00AA
    Exit Sub
  End If
On Error GoTo 0
  Open l00A2 For Binary Access Read As l00A4
  Get l00A4, , l00A8
  If l00A8 < 6 Then
    Close l00A4
    gv7B4A = l00A2
    sub022F mc00CE
    sub0220 l00AA
    Exit Sub
  End If
  gv091C = False
  Get l00A4, , l00A6
  If l00A6 Then
    If l00A6 <> gv09FA Then sub01F8
    Get l00A4, , gv7952
  End If
  If l00A8 >= 8 Then
    For l00A0 = 1 To gv09FA
      Get l00A4, , gv7988(l00A0)
    Next
    Get l00A4, , l00A6
    ReDim gv7956(l00A6)
    For l00A0 = 1 To l00A6
      Get l00A4, , l00A6
      If l00A6 Then
        gv7956(l00A0) = Space$(l00A6)
        Get l00A4, , gv7956(l00A0)
      End If
    Next
  Else
    l00A6 = gv09FA
    For l00A0 = 1 To l00A6
      Get l00A4, , l00A6
      If l00A6 Then
        l00D0 = UBound(gv7956) + 1
        ReDim Preserve gv7956(l00D0)
        gv7956(l00D0) = Space$(l00A6)
        Get l00A4, , gv7956(l00D0)
        gv7988(l00A0) = l00D0
      End If
    Next l00A0
    gv091C = True
  End If
  If l00A8 = 6 Then
    l00A2 = Space$(gv09FA)
    Get l00A4, , l00A2
    For l00A0 = 1 To gv09FA
      gv09FC(l00A0) = Asc(Mid$(l00A2, l00A0, 1))
    Next
    gv091C = True
  Else
    For l00A0 = 1 To gv09FA
      Get l00A4, , gv09FC(l00A0)
    Next
  End If
  Close l00A4
  sub0220 l00AA
End Sub

Sub sub01C8(p004E As Integer)
  gv7B2A = p004E
  gv7B2C = gv7ABC(p004E).M24E0
  Get gv088E, fn0E71(gv7B2C), gv7AF2
  If gv7ABC(p004E).M1C2B Then
    sub0088 gv7ABC(p004E).M1C2B  'seek to segment
    sub01D7
  End If
End Sub

Sub sub01E7()
Dim l0064 As Integer
Dim l0066 As Integer
Dim l0068  As Integer
Dim l006A  As Integer
Dim l006C As String
  gv091A = True
  l0066 = FreeFile
  l006C = fn0FBE("", "0")
On Error Resume Next
  If FileLen(l006C) = 0 Then Exit Sub
  If Err Then Exit Sub
On Error GoTo 0
  Open l006C For Binary Access Read As l0066
  Get l0066, , l006A
  If l006A > 7 Then
    Get l0066, , gv7A1C
    Screen.MousePointer = 11
    Get l0066, , l0068
    ReDim gv7876(l0068)
    For l0064 = 1 To l0068
      Get l0066, , l0068
      If l0068 Then
        gv7876(l0064) = Space$(l0068)
        Get l0066, , gv7876(l0064)
      End If
    Next l0064
    Get l0066, , l0068
    If l0068 <> gv09C2 + 1 Then sub01F8
    gv78DA = Space$(l0068)
    Get l0066, , gv78DA
    For l0064 = 0 To gv09C2
      Get l0066, , gv78A8(l0064)
    Next
    Get l0066, , l0068
    ReDim gv790A(l0068)
    For l0064 = 0 To l0068
      Get l0066, , gv790A(l0064)
    Next
    Get l0066, , l0068
    ReDim gv1042(l0068)
    For l0064 = 1 To l0068
      Get l0066, , gv1042(l0064)
    Next l0064
    Get l0066, , l0068
    ReDim gv1076(l0068)
    For l0064 = 1 To l0068
      Get l0066, , gv1076(l0064)
    Next
    For l0064 = 1 To gv7AEE
      Get l0066, , gv7ABC(l0064)
    Next
    gv091A = False
    Screen.MousePointer = 0
  End If
  Close l0066
End Sub

Sub sub023C()
  gv7A70(gv7AB6) = gv7AA6
  gv091A = True
End Sub

