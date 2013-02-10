Attribute VB_Name = "modFileListbox"
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'***********************************
Option Explicit

Public Function ProccessFileListBox(ByVal F As Integer, ByVal Opcode As Byte) As Long
    On Error GoTo errHandle
        Dim bData As Byte
        Dim iData As Integer
        Dim lSize As Long
        Dim bArray() As Byte
    Select Case Opcode
        Case 1
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
        Case 11
            Get #F, , iData
            Call frmMain.AddText("TabIndex = " & iData)
        Case 13 'Pattern
            Get #F, , bData
            Call frmMain.AddText("Pattern = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 15 'Normal
            Get #F, , bData
            Call frmMain.AddText("Normal = " & bData)
        Case 16 'ReadOnly
            Get #F, , bData
            Call frmMain.AddText("ReadOnly = " & bData)
        Case 17 'Archive
            Get #F, , bData
            Call frmMain.AddText("Archive = " & bData)
        Case 18 'Hidden
            Get #F, , bData
            Call frmMain.AddText("Hidden = -1")
        Case 19 'System
            Get #F, , bData
            Call frmMain.AddText("System = -1")
        Case 23 'Font
            Call modGlobals.GetFontProperty(F)
        Case 30 'DragMode
            Get #F, , bData
            Call frmMain.AddText("DragMode = " & bData)
        Case 31 'DragIcon
            lSize = gVBFormFile.GetLong(Loc(F))
            If lSize <> -1 Then
                
                ReDim bArray(lSize)
                Get #F, , bArray
                Seek #F, Loc(F)
                Call frmMain.AddText("DragIcon = " & frmMain.ReturnFormName() & ".frx:" & PadHex(Hex(gFrxAddress), 4))
                gFrxAddress = gFrxAddress + lSize
            End If
        Case 32 'TabStop
            Get #F, , bData
            Call frmMain.AddText("TabStop = " & bData)
        Case 33 'Tag
            Get #F, , bData
            Call frmMain.AddText("Tag = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 35 'HelpContextID
            Call frmMain.AddText("HelpContextID = " & gVBFormFile.GetLong(Loc(F)))
        Case 36 'MultiSelect
            Get #F, , bData
            Call frmMain.AddText("MultiSelect = " & bData)
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
            
            Call AddError("Error_Unknown Opcode_ProcessFileListBox: Opcode: " & Opcode & " Offset: " & Loc(F))
            bFirstFF = True
            Seek F, frmMain.ReturnControlEnd
    End Select
    
Exit Function
errHandle:
    Call AddError("Error_ProcessFileListBox: Opcode: " & Opcode & " Offset: " & Loc(F) & " Description= " & Err.Description)
            
End Function

