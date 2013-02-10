Attribute VB_Name = "modImage"
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'***********************************
Option Explicit

Public Function ProccessImage(ByVal F As Integer, ByVal Opcode As Byte) As Long
    On Error GoTo errHandle
        Dim bData As Byte
        Dim iData As Integer
        Dim lSize As Long
        Dim bArray() As Byte
    Select Case Opcode
        Case 1
            Get #F, , iData
            Call frmMain.AddText("Index = " & iData)
        Case 2 'Picture
            lSize = gVBFormFile.GetLong(Loc(F))
            If lSize <> -1 Then
                
                ReDim bArray(lSize)
                Get #F, , bArray
                Seek #F, Loc(F)
                Call frmMain.AddText("Picture = " & frmMain.ReturnFormName() & ".frx:" & PadHex(Hex(gFrxAddress), 4))
                gFrxAddress = gFrxAddress + lSize
            End If
        Case 3
            Call modGlobals.GetControlSize(F)
        Case 7 'Enabled
            Get #F, , bData
            Call frmMain.AddText("Enabled = " & bData)
        Case 8 'Visible
            Get #F, , bData
            Call frmMain.AddText("Visible = " & bData)
        Case 9 'MousePointer
            Get #F, , bData
            Call frmMain.AddText("MousePointer = " & bData)
        Case 10 'Stretch
            Get #F, , bData
            Call frmMain.AddText("Stretch = -1")
        Case 12
            Get #F, , bData
            Call frmMain.AddText("DragMode = " & bData)
        Case 13 'DragIcon
            lSize = gVBFormFile.GetLong(Loc(F))
            If lSize <> -1 Then
                
                ReDim bArray(lSize)
                Get #F, , bArray
                Seek #F, Loc(F)
                Call frmMain.AddText("DragIcon = " & frmMain.ReturnFormName() & ".frx:" & PadHex(Hex(gFrxAddress), 4))
                gFrxAddress = gFrxAddress + lSize
            End If
        Case 14 'Tag
            Get #F, , bData
            Call frmMain.AddText("Tag  = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 15 'BorderStyle
            Get #F, , bData
            Call frmMain.AddText("BorderStyle = " & bData)
        Case 16 'DataSource
            Get #F, , bData
            Call frmMain.AddText("DataSource = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 17 'DataField
            Get #F, , bData
            Call frmMain.AddText("DataField = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
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
            
            Call AddError("Error_Unknown Opcode_ProcessImage: Opcode: " & Opcode & " Offset: " & Loc(F))
            bFirstFF = True
            Seek F, frmMain.ReturnControlEnd
    End Select
    
Exit Function
errHandle:
    Call AddError("Error_ProcessImage: Opcode: " & Opcode & " Offset: " & Loc(F) & " Description= " & Err.Description)
            
End Function

