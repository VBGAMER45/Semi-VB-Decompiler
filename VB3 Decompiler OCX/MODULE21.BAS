Attribute VB_Name = "Module21"
' Module21
Option Explicit

Sub sub0FB1()
Dim l001C As Integer
Dim l001E As String
  l001E = gv7B84
  l001C = fn02FF(l001E, "\")
  If l001C Then
    gv7B88 = Left(l001E, l001C)
    l001E = Mid$(l001E, l001C + 1)
  End If
  l001C = InStr(l001E, ".")
  If l001C Then l001E = Left$(l001E, l001C - 1)
  gv7B90 = l001E
End Sub

Function fn0FBE(p0024 As String, ByVal p0026 As String) As String
Dim l0028 As Integer
  'If gv7B8C = "" Then form2.Show 1
  gv7B8C = m_OutputFolder
  If Len(p0024$) Then
    l0028 = InStr(p0024, ".")
    If l0028 Then
      fn0FBE = gv7B8C & Left$(p0024, l0028) & p0026
    Else
      fn0FBE = gv7B8C & p0024 & "." & p0026
    End If
  Else
    fn0FBE = gv7B8C & gv7B90 & "." & p0026
  End If
End Function

Sub sub0FCA()
Dim l002E As Integer
Dim l0030 As Integer
Dim l0032 As String
Dim l0034 As String
  gv7BAA = 0
  l0030 = FreeFile
  gv0886 = gv7B88
  Open gv7B84 For Input As #l0030
  While Not EOF(l0030)
    Line Input #l0030, l0032
    l002E = InStr(l0032, "=")
    If l002E Then
      If Mid$(l0032, l002E + 1, 1) = """" Then
        l0034 = Mid$(l0032, l002E + 2)
        l0034 = Left$(l0034, Len(l0034) - 1)
        Select Case Left$(l0032, l002E - 1)
        Case "ExeName"
          gv088A = l0034
        Case "Path"
          gv0886 = l0034
        End Select
      End If
    Else
      If Right$(l0032, 1) <> "X" And Right$(l0032, 3) <> "DLL" Then
        gv7BAA = gv7BAA + 1
        gv7B94(gv7BAA) = l0032
      End If
    End If
  Wend
  Close l0030
  l002E = InStr(gv088A, ".")
  If l002E Then gv088A = Left$(gv088A, l002E - 1)
  If Right$(gv0886, 1) <> "\" Then gv0886 = gv0886 & "\"
  gv0882 = gv0886 & gv088A & ".EXE"
End Sub
