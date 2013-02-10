Attribute VB_Name = "modLabel"
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'***********************************
Option Explicit

Public Function ProccessLabel(ByVal F As Integer, ByVal Opcode As Byte) As Long
    On Error GoTo errHandle
        Dim bData As Byte
        Dim iData As Integer
        Dim lSize As Long
        Dim bArray() As Byte
    Select Case Opcode
        Case 0 'Caption
            Get #F, , bData
            Call frmMain.AddText("Caption  = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 2 'Index
            Get #F, , iData
            Call frmMain.AddText("Index = " & iData)
        Case 3
            Call frmMain.AddText("BackColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 4 ' ForeColor
            Call frmMain.AddText("ForeColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 5 'Control Size
            Call modGlobals.GetControlSize(F)
        Case 9
            Get #F, , bData
            Call frmMain.AddText("Enabled = " & bData)
        Case 10 'Visible
            Get #F, , bData
            Call frmMain.AddText("Visible = " & bData)
        Case 11
            Get #F, , bData
            Call frmMain.AddText("MousePointer = " & bData)
        Case 12 'FONT
            Call modGlobals.GetFontProperty(F)
        Case 18 'TabIndex
            Get #F, , iData
            Call frmMain.AddText("TabIndex = " & iData)
        Case 19
            Get #F, , bData
            Call frmMain.AddText("BorderStyle = " & bData)
        Case 20 ' Alignment
            Get #F, , bData
            Call frmMain.AddText("Alignment = " & bData)
        Case 22 'LinkItem
            Get #F, , bData
            Call frmMain.AddText("LinkItem = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 24
            Get #F, , bData
            Call frmMain.AddText("AutoSize = -1")
        Case 26
            Get #F, , bData
            Call frmMain.AddText("DragMode = " & bData)
        Case 27
            lSize = gVBFormFile.GetLong(Loc(F))
            If lSize <> -1 Then
                
                ReDim bArray(lSize)
                Get #F, , bArray
                Seek #F, Loc(F)
                Call frmMain.AddText("DragIcon = " & frmMain.ReturnFormName() & ".frx:" & PadHex(Hex$(gFrxAddress), 4))
                gFrxAddress = gFrxAddress + lSize
            End If
        Case 28
            Get #F, , iData
            Call frmMain.AddText("LinkTimeout = " & iData)
        Case 29
            Get #F, , bData
            Call frmMain.AddText("Tag = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 30
            Get #F, , bData
            Call frmMain.AddText("WordWrap = -1")
        Case 31 'BackStyle
            Get #F, , bData
            Call frmMain.AddText("BackStyle = " & bData)
        Case 32
            Get #F, , bData
            Call frmMain.AddText("DataSource = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 33
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
            
            Call AddError("Error_Unknown Opcode_ProcessLabel: Opcode: " & Opcode & " Offset: " & Loc(F))
            bFirstFF = True
            Seek F, frmMain.ReturnControlEnd
    End Select
    
Exit Function
errHandle:
    Call AddError("Error_ProcessLabel: Opcode: " & Opcode & " Offset: " & Loc(F) & " Description= " & Err.Description)
            
End Function

