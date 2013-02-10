Attribute VB_Name = "modTimer"
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'***********************************
Option Explicit

Public Function ProccessTimer(ByVal F As Integer, ByVal Opcode As Byte) As Long
    On Error GoTo errHandle
        Dim bData As Byte
        Dim iData As Integer
    Select Case Opcode
        Case 1
            Get #F, , iData
            Call frmMain.AddText("Index = " & iData)
        Case 3 'Interval
            Call frmMain.AddText("Interval = " & gVBFormFile.GetLong(Loc(F)))
        Case 5
            Get #F, , bData
            Call frmMain.AddText("Tag = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 7
            Call frmMain.AddText("Left = " & gVBFormFile.GetLong(Loc(F)))
        Case 8
            Call frmMain.AddText("Top = " & gVBFormFile.GetLong(Loc(F)))
            
        Case 255
          
            Do
                Get #F, , bData

                If bData = 1 Then
                   
                ElseIf bData = 4 Then
                    gIdentSpaces = gIdentSpaces - 1
                    Call frmMain.AddText("End")
                    gFormDone = True
                ElseIf bData = 3 Then
                    gIdentSpaces = gIdentSpaces - 1
                    Call frmMain.AddText("End")
                ElseIf bData = 2 Then
                    gIdentSpaces = gIdentSpaces - 1
                    Call frmMain.AddText("End")
                End If
           
            Loop While bData <> 0 And bData < 6
            Seek F, Loc(F)
        Case Else
            
            Call AddError("Error_Unknown Opcode_ProcessTimer: Opcode: " & Opcode & " Offset: " & Loc(F))
            bFirstFF = True
            Seek F, frmMain.ReturnControlEnd
    End Select
    
Exit Function
errHandle:
    Call AddError("Error_ProcessTimer: Opcode: " & Opcode & " Offset: " & Loc(F) & " Description= " & Err.Description)
            
End Function

