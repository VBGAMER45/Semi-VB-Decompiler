Attribute VB_Name = "modHscroll"
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'***********************************
Option Explicit

Public Function ProccessHscroll(ByVal F As Integer, ByVal Opcode As Byte) As Long
    On Error GoTo errHandle
        Dim bData As Byte
        Dim iData As Integer
        Dim lSize As Long
        Dim bArray() As Byte
    Select Case Opcode
        Case 1
            Get #F, , iData
            Call frmMain.AddText("Index = " & iData)
        Case 2
            Call modGlobals.GetControlSize(F)
        Case 7 'Visible
            Get #F, , bData
            Call frmMain.AddText("Visible = " & bData)
        Case 9
            Get #F, , iData
            Call frmMain.AddText("TabIndex = " & iData)
        Case 10 'Min
            Get #F, , iData
            Call frmMain.AddText("Min = " & iData)
        Case 11 'Max
            Get #F, , iData
            Call frmMain.AddText("Max = " & iData)
        Case 12 'SmallChange
            Get #F, , iData
            Call frmMain.AddText("SmallChange = " & iData)
        Case 13 ' LargeChange
            Get #F, , iData
            Call frmMain.AddText("LargeChange = " & iData)
        Case 14 'Value
            Get #F, , iData
            Call frmMain.AddText("Value = " & iData)
        Case 16 'DragMode
            Get #F, , bData
            Call frmMain.AddText("DragMode = " & bData)
        Case 17 'DragIcon
            lSize = gVBFormFile.GetLong(Loc(F))
            If lSize <> -1 Then
                
                ReDim bArray(lSize)
                Get #F, , bArray
                Seek #F, Loc(F)
                Call frmMain.AddText("DragIcon = " & frmMain.ReturnFormName() & ".frx:" & PadHex(Hex(gFrxAddress), 4))
                gFrxAddress = gFrxAddress + lSize
            End If
        Case 19 'Tag
            Get #F, , bData
            Call frmMain.AddText("Tag  = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))

        Case 21 'HelpContextID
            Call frmMain.AddText("HelpContextID = " & gVBFormFile.GetLong(Loc(F)))
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
           
            Call AddError("Error_Unknown Opcode_ProcessHscroll: Opcode: " & Opcode & " Offset: " & Loc(F))
            bFirstFF = True
            Seek F, frmMain.ReturnControlEnd
    End Select
    
Exit Function
errHandle:
    Call AddError("Error_ProcessHscroll: Opcode: " & Opcode & " Offset: " & Loc(F) & " Description= " & Err.Description)
            
End Function

