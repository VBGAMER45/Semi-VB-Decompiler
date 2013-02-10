Attribute VB_Name = "modDriveListBox"
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'***********************************
Option Explicit

Public Function ProccessDriveListBox(ByVal F As Integer, ByVal Opcode As Byte) As Long
    On Error GoTo errHandle
        Dim bData As Byte
        Dim iData As Integer
        Dim lSize As Long
        Dim bArray() As Byte
    Select Case Opcode
        Case 1 'Index
            Get #F, , iData
            Call frmMain.AddText("Index = " & iData)
        Case 2 'BackColor
            Call frmMain.AddText("BackColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 3 'ForeColor
            Call frmMain.AddText("ForeColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 4
            Call modGlobals.GetControlSize(F)
        Case 8 'Enabled
            Get #F, , bData
            Call frmMain.AddText("Enabled = " & bData)
        Case 9 'Visible
            Get #F, , bData
            Call frmMain.AddText("Visible = " & bData)
        Case 10 'MousePointer
            Get #F, , bData
            Call frmMain.AddText("MousePointer = " & bData)
        Case 11 'TabIndex
            Get #F, , iData
            Call frmMain.AddText("TabIndex = " & iData)
        Case 16 'Font
            Call modGlobals.GetFontProperty(F)
        Case 23 'DragMode
            Get #F, , bData
            Call frmMain.AddText("DragMode = " & bData)
        Case 24 'DragIcon
            lSize = gVBFormFile.GetLong(Loc(F))
            If lSize <> -1 Then
                
                ReDim bArray(lSize)
                Get #F, , bArray
                Seek #F, Loc(F)
                Call frmMain.AddText("DragIcon = " & frmMain.ReturnFormName() & ".frx:" & PadHex(Hex(gFrxAddress), 4))
                gFrxAddress = gFrxAddress + lSize
            End If
        Case 25 'TabStop
            Get #F, , bData
            Call frmMain.AddText("TabStop = " & bData)
        Case 26 'Tag
            Get #F, , bData
            Call frmMain.AddText("Tag  = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))

        Case 28 'HelpContextID
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
            
            Call AddError("Error_Unknown Opcode_ProcessDriveListBox: Opcode: " & Opcode & " Offset: " & Loc(F))
            bFirstFF = True
            Seek F, frmMain.ReturnControlEnd
    End Select
    
Exit Function
errHandle:
    Call AddError("Error_ProcessDriveListBox: Opcode: " & Opcode & " Offset: " & Loc(F) & " Description= " & Err.Description)
            
End Function

