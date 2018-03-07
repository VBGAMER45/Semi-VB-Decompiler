Attribute VB_Name = "modTools"
Option Explicit


Public Function GetPart(DataStr As String, DataId As Long, Separator As String) As Variant
    Dim pointer As Long, i As Long
    On Error GoTo errHandler
    For i = 1 To DataId
        pointer = InStr(pointer + 1, DataStr, Separator)
    Next i
    GetPart = Mid$(DataStr, pointer + Len(Separator), IIf(InStr(pointer + 1, DataStr, Separator) = 0, Len(DataStr), InStr(pointer + 1, DataStr, Separator)) - pointer - Len(Separator))
    Exit Function
errHandler:
    GetPart = False
End Function




