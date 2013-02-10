Attribute VB_Name = "modShape"
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'***********************************
Option Explicit

Public Function ProccessShape(ByVal F As Integer, ByVal Opcode As Byte) As Long
    On Error GoTo errHandle
        Dim bData As Byte
        Dim iData As Integer
    Select Case Opcode
        Case 1
            Get #F, , iData
            Call frmMain.AddText("Index = " & iData)
        Case 2 'BackColor
            Call frmMain.AddText("BackColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 3 'BorderColor
            Call frmMain.AddText("BorderColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 4
            Call modGlobals.GetControlSize(F)
        Case 8 'Visible
            Get #F, , bData
            Call frmMain.AddText("Visible = -1")
        Case 10 'Tag
            Get #F, , bData
            Call frmMain.AddText("Tag  = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))

        Case 11 'Shape
            Get #F, , bData
            Call frmMain.AddText("Shape = " & bData)
        Case 12 'DrawMode
            Get #F, , bData
            Call frmMain.AddText("DrawMode = " & bData)
        Case 13 'BorderStyle
            Get #F, , bData
            Call frmMain.AddText("BorderStyle = " & bData)
        Case 14 'BorderWidth
            Get #F, , iData
            Call frmMain.AddText("BorderWidth = " & iData)
        Case 15 'FillColor
            Call frmMain.AddText("FillColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 16 'BackStyle
            Get #F, , bData
            Call frmMain.AddText("BackStyle = " & bData)
        Case 17 'FillStyle
            Get #F, , bData
            Call frmMain.AddText("FillStyle = " & bData)
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
            
            Call AddError("Error_Unknown Opcode_ProcessShape: Opcode: " & Opcode & " Offset: " & Loc(F))
            bFirstFF = True
            Seek F, frmMain.ReturnControlEnd
    End Select
    
Exit Function
errHandle:
    Call AddError("Error_ProcessShape: Opcode: " & Opcode & " Offset: " & Loc(F) & " Description= " & Err.Description)
            
End Function

