Attribute VB_Name = "modCommandButton"
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'***********************************
Option Explicit

Public Function ProccessCommandButton(ByVal F As Integer, ByVal Opcode As Byte) As Long
    On Error GoTo errHandle
        Dim bData As Byte
        Dim iData As Integer
        Dim lSize As Long
        Dim bArray() As Byte
    Select Case Opcode
        Case 0 'Caption
            Get #F, , bData
            Call frmMain.AddText("Caption  = " & Chr$(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr$(34))
        Case 2 'Index
            Get #F, , iData
            Call frmMain.AddText("Index = " & iData)
        Case 3 'BackColor
            Call frmMain.AddText("BackColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 4 'Control Size
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
        Case 11 'Font
            Call modGlobals.GetFontProperty(F)
        Case 17 'TabIndex
            Get #F, , iData
            Call frmMain.AddText("TabIndex = " & iData)
        Case 19 'Default
            Get #F, , bData
            Call frmMain.AddText("Default = -1")
        Case 20 'Cancel
            Get #F, , bData
            Call frmMain.AddText("Cancel = -1")
        Case 22 'DragMode
            Get #F, , bData
            Call frmMain.AddText("DragMode = " & bData)
        Case 23 'DragIcon
            lSize = gVBFormFile.GetLong(Loc(F))
            If lSize <> -1 Then
                
                ReDim bArray(lSize)
                Get #F, , bArray
                Seek #F, Loc(F)
                Call frmMain.AddText("DragIcon = " & frmMain.ReturnFormName() & ".frx:" & PadHex(Hex(gFrxAddress), 4))
                gFrxAddress = gFrxAddress + lSize
            End If
        Case 24 'TabStop
            Get #F, , bData
            Call frmMain.AddText("TabStop = " & bData)
        Case 25 'Tag
            Get #F, , bData
            Call frmMain.AddText("Tag = " & Chr$(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr$(34))
        Case 27 'HelpContextID
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
            
            Call AddError("Error_Unknown Opcode_ProcessCommandButton: Opcode: " & Opcode & " Offset: " & Loc(F))
            bFirstFF = True
            Seek F, frmMain.ReturnControlEnd
    End Select
    
Exit Function
errHandle:
    Call AddError("Error_ProcessCommandButton: Opcode: " & Opcode & " Offset: " & Loc(F) & " Description= " & Err.Description)
            
End Function

