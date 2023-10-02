Attribute VB_Name = "Module7"
' Module7
Option Explicit

Sub sub05FF (p0070 As String)
  gv0C5A = gv0C5A & p0070
End Sub

Sub sub0632 (p0060 As Integer)
  gv0C5A = Space$(p0060)
  gv0C5E = p0060
End Sub

Sub sub063C (p0082 As Integer)
Dim l0084 As Integer
  l0084 = p0082 - Len(gv0C5A)
  If l0084 > 0 Then gv0C5A = gv0C5A & Space$(l0084)
 End Sub

Sub sub082B (p0086 As Integer, p0088 As Long, p008A As Integer, p008C As Integer)
  sub083A p0086, p0088, fn0335(p008A), fn0335(p008C)
  p008A = p008C
End Sub

Sub sub083A (p0010 As Integer, p0012 As Long, p0014 As Long, p0016 As Long)
Dim l001A As Integer
Dim l001C As Integer
Dim l001E As Integer
Dim l0020 As Long
Dim l0024 As String
Dim l0026 As String
  l0026 = Hex$(p0016): If Len(l0026) < 4 Then l0026 = Space$(4)
  Seek p0010, p0012 + p0014
  l0020 = p0016 - p0014
  If gv0C52 = 0 And l0020 > 512 Then l0020 = 512
  l001C = 16
  l0024 = Space$(l001C)
  While l0020 > 0
    If l0020 < l001C Then l001C = l0020: l0024 = Space$(l001C)
    Get p0010, , l0024
    For l001A = 1 To l001C
      l001E = Asc(Mid$(l0024, l001A, 1))
      sub05FF fn029A(l001E)
      If l001E < 32 Or l001E > &H7E Then Mid$(l0024, l001A, 1) = " "
    Next l001A
    If 49 > Len(gv0C5A) Then sub05FF Space$(49 - Len(gv0C5A))
    RSet l0026 = Hex$(p0014)
    gv0C5A = l0026 & gv0C5A & l0024: sub0868
    p0014 = p0014 + l001C
    l0020 = l0020 - l001C
  Wend
  If p0014 <> p0016 Then
    gv0C5A = "...": sub0868
    p0014 = p0016
  End If
  Seek p0010, p0012 + p0014
End Sub

Sub sub084A (p0090 As Integer, p0092 As Long, p0094 As Long, p0096 As Long)
  If (p0094 Xor p0096) And 1 Then
    sub083A p0090, p0092, p0094, p0096
  Else
    sub0858 p0090, p0092, p0094, p0096
  End If
End Sub

Sub sub0858 (p0098 As Integer, p009A As Long, p009C As Long, p009E As Long)
Dim l00A0 As Integer
Dim l00A2 As Integer
Dim l00A4 As Long
Dim l00A6 As Integer
Dim l00A8 As String
  l00A8 = Hex$(p009E): If Len(l00A8) < 4 Then l00A8 = Space$(4)
  l00A4 = p009E - p009C
  If gv0C52 = 0 And l00A4 > 512 Then l00A4 = 512
  Seek p0098, p009A + p009C
  l00A2 = 16
  While l00A4 > 0
    RSet l00A8 = Hex$(p009C): gv0C5A = l00A8
    If l00A4 < l00A2 Then l00A2 = l00A4
    For l00A0 = 1 To l00A2 Step 2
      Get p0098, , l00A6
      sub05FF fn02AC(l00A6)
    Next l00A0
    p009C = p009C + l00A2
    l00A4 = l00A4 - l00A2
    sub0868
  Wend
  If p009C < p009E Then
    gv0C5A = "...": sub0868
  End If
  p009C = p009E
  Seek p0098, p009A + p009C
End Sub

Sub sub0868 ()
On Error Resume Next
  If gv0C52 Then
    Print #gv0C52, gv0C5A
  Else
    If Len(gv0C5A) Then
      gv0C56.AddItem gv0C5A
    End If
  End If
  gv0C5A = ""
End Sub

Sub sub0872 (p002A As String, p002C As Integer)
Dim l002E As Integer
  l002E = InStr(p002C, p002A, Chr$(0)) - p002C
  sub05FF " '" & Mid$(p002A, p002C, l002E) & "'"
  sub0868
End Sub

Sub sub087E ()
  If gv0C52 Then
    Close gv0C52
    gv0C52 = 0
  End If
  Set gv0C56 = Nothing
  Screen.MousePointer = 0
End Sub

Sub sub0888 (p0034 As Integer, p0036 As Long, p0038 As Long)
Dim l003A As Integer
Dim l003C As Integer
Dim l003E As String
Dim l0040 As Integer
Dim l0042 As Long
Dim l0044 As Long
  If p0036 = 0 Then Exit Sub
  Seek p0034, p0036
  l0042 = p0038 - p0036
  If gv0C52 = 0 And l0042 > 512 Then l0042 = 512
  l003C = 16
  l003E = Space$(l003C)
  While l0042 > 0
    If l0044 < &H1000 Then
      gv0C5A = fn02B5(Int(l0044))
    Else
      gv0C5A = Hex$(l0044)
    End If
    If l0042 < 16 Then l003C = l0042: l003E = Space$(l0042)
    Get p0034, , l003E
    For l003A = 1 To l003C
      l0040 = Asc(Mid$(l003E, l003A, 1))
      sub05FF fn029A(l0040)
      If l0040 < 32 Then Mid$(l003E, l003A, 1) = " "
    Next l003A
    If 54 > Len(gv0C5A) Then sub05FF Space$(54 - Len(gv0C5A))
    sub05FF l003E
    sub0868
    l0044 = l0044 + l003C
    l0042 = l0042 - l003C
  Wend
End Sub

Sub sub0892 (p0048 As Integer, p004A As Long, p004C As Long)
Dim l0050  As Integer
Dim l0052 As Integer
Dim l0056 As Long
Dim l0058 As Long
Dim l005A  As Integer
  l0058 = 0
  Seek p0048, p004A
  l0056 = p004C - p004A
  If gv0C52 = 0 And l0056 > 512 Then l0056 = 512
  l0052 = 16
  While l0056 > 0
    If l0058 < &H1000 Then
      gv0C5A = fn02B5(Int(l0058))
    Else
      gv0C5A = Hex$(l0058)
    End If
    If l0056 < l0052 Then l0052 = l0056
    For l0050 = 1 To l0052 Step 2
      Get p0048, , l005A
      sub05FF fn02AC(l005A)
    Next l0050
    l0058 = l0058 + l0052
    l0056 = l0056 - l0052
    sub0868
  Wend
End Sub

Sub sub089C (p005E As String)
  If gv0C52 Then
    MsgBox "file already open"
    Close gv0C52
  End If
  gv0C52 = FreeFile
  Screen.MousePointer = 11
  Open p005E For Output As gv0C52
End Sub

Sub sub08A5 (p0064 As Control)
  Set gv0C56 = p0064
  p0064.Clear
End Sub

Sub sub08AF (p0066 As String, p0068 As Integer, p006A As String)
Dim l006C As Integer
Dim l006E As Integer
  For l006C = 1 To Len(p006A)
    Select Case Mid$(p006A, l006C, 1)
    Case "b"
      sub05FF Right$(Hex$(Asc(Mid$(p0066, p0068, 1)) + &H100), 2)
      p0068 = p0068 + 1
    Case "w"
      sub05FF Right$(Hex$(Asc(Mid$(p0066, p0068 + 1, 1)) + &H100), 2)
      sub05FF Right$(Hex$(Asc(Mid$(p0066, p0068, 1)) + &H100), 2)
      p0068 = p0068 + 2
    Case "l"
      sub05FF Right$(Hex$(Asc(Mid$(p0066, p0068 + 3, 1)) + &H100), 2)
      sub05FF Right$(Hex$(Asc(Mid$(p0066, p0068 + 2, 1)) + &H100), 2)
      sub05FF Right$(Hex$(Asc(Mid$(p0066, p0068 + 1, 1)) + &H100), 2)
      sub05FF Right$(Hex$(Asc(Mid$(p0066, p0068, 1)) + &H100), 2)
      p0068 = p0068 + 4
    Case "n"
      sub0868
    Case "s"
      l006E = Asc(Mid$(p0066, p0068, 1))
      sub05FF "'" & Mid$(p0066, p0068 + 1, l006E - 1) & "'"
      p0068 = p0068 + l006E + 1
    Case "t"
      l006E = Len(gv0C5A) Mod 14
      If l006E Then sub05FF Space$(14 - l006E)
    Case "z"
      l006E = InStr(p0068, p0066, Chr$(0)) - p0068
      sub05FF "'" & Mid$(p0066, p0068, l006E) & "'"
    Case "Z"
      l006E = InStr(p0068, p0066, Chr$(0)) - p0068
      sub05FF "'" & Mid$(p0066, p0068, l006E) & "'"
      p0068 = p0068 + l006E + 1
    Case "\"
      p0068 = p0068 + 1
    Case Else
      sub05FF Mid$(p006A, l006C, 1)
    End Select
  Next l006C
End Sub

Sub sub08B9 (p0072 As Long, p0074() As String)
Dim l007A As Long
Dim l007C As Integer
Dim l0080 As Integer
  l007A = 1
  For l007C = 0 To UBound(p0074)
    If p0072 And l007A Then
      GoSub LFFDE
      sub05FF p0074(l007C)
      p0072 = p0072 Xor l007A
    End If
    l007A = l007A * 2
  Next
  If p0072 Then
    GoSub LFFDE
    sub05FF "0x" & Hex$(p0072)
  End If
Exit Sub

LFFDE:
  If l0080 Then
    sub05FF " | "
  Else
    sub05FF " ": l0080 = True
  End If
Return
End Sub

