Attribute VB_Name = "Module15"
' Module15
Option Explicit
Const mc0008 = -32768 ' &H8000%

Sub sub0D9D(p0006 As Integer)
' Const mc0008 =-32768 ' &H8000%
  If p0006 Then Exit Sub
  sub022F mc0008
  'End
End Sub

Sub sub0394(p000A As Integer, p000C As Integer, p000E As Integer)
  If p000E Then
    gv0FF6(p000A).M25C6 = -p000C
  Else
    gv0FF6(p000A).M25C6 = p000C
  End If
End Sub
