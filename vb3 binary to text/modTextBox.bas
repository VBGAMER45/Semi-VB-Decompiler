Attribute VB_Name = "modTextBox"
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'***********************************
Option Explicit

Public Function ProccessTextBox(ByVal F As Integer, ByVal Opcode As Byte) As Long
    On Error GoTo errHandle
        Dim bData As Byte
        Dim iData As Integer
        Dim lSize As Long
        Dim bArray() As Byte
    Select Case Opcode
        Case 1 'Index
            Get #F, , iData
            Call frmMain.AddText("Index = " & iData)
        Case 2
            Call frmMain.AddText("BackColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 3 'ForeColor
            Call frmMain.AddText("ForeColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 4
            Call modGlobals.GetControlSize(F)
        Case 8
            Get #F, , bData
            Call frmMain.AddText("Enabled = " & bData)
        Case 9 'Visible
            Get #F, , bData
            Call frmMain.AddText("Visible = " & bData)
        Case 10 'MousePointer
            Get #F, , bData
            Call frmMain.AddText("MousePointer = " & bData)
        Case 11 'Text
            Get #F, , bData
            Call frmMain.AddText("Text = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 12
           Call modGlobals.GetFontProperty(F)
        Case 18
            Get #F, , iData
            Call frmMain.AddText("TabIndex = " & iData)
        Case 19
            Get #F, , bData
            Call frmMain.AddText("BorderStyle = " & bData)
        Case 21 'LinkItem
            Get #F, , bData
            Call frmMain.AddText("LinkItem = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 23 'MultiLine
            Get #F, , bData
            Call frmMain.AddText("MultiLine = -1")
        Case 24 'ScrollBars
            Get #F, , bData
            Call frmMain.AddText("ScrollBars = " & bData)
        Case 29
            Get #F, , bData
            Call frmMain.AddText("DragMode = " & bData)
        Case 30
            lSize = gVBFormFile.GetLong(Loc(F))
            If lSize <> -1 Then
                
                ReDim bArray(lSize)
                Get #F, , bArray
                Seek #F, Loc(F)
                Call frmMain.AddText("DragIcon = " & frmMain.ReturnFormName() & ".frx:" & PadHex(Hex$(gFrxAddress), 4))
                gFrxAddress = gFrxAddress + lSize
            End If
        Case 31
            Get #F, , iData
            Call frmMain.AddText("LinkTimeout = " & iData)
        Case 32 'TabStop
            Get #F, , bData
            Call frmMain.AddText("TabStop = " & bData)
        Case 33 'Tag
            Get #F, , bData
            Call frmMain.AddText("Tag = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 34 'PasswordChar
            Get #F, , bData
            Call frmMain.AddText("PasswordChar = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 35
            Get #F, , bData
            Call frmMain.AddText("HideSelection = " & bData)
        Case 36
            Get #F, , bData
            Call frmMain.AddText("Alignment = " & bData)
        Case 37 'MaxLength
            Call frmMain.AddText("MaxLength = " & gVBFormFile.GetLong(Loc(F)))
        Case 38 'HelpContextID
            Call frmMain.AddText("HelpContextID = " & gVBFormFile.GetLong(Loc(F)))
        Case 41 'DataSource
            Get #F, , bData
            Call frmMain.AddText("DataSource = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 42 'DataField
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
            
            Call AddError("Error_Unknown Opcode_ProcessTextBox: Opcode: " & Opcode & " Offset: " & Loc(F))
            bFirstFF = True
            Seek F, frmMain.ReturnControlEnd
    End Select
    
Exit Function
errHandle:
    Call AddError("Error_ProcessTextBox: Opcode: " & Opcode & " Offset: " & Loc(F) & " Description= " & Err.Description)
            
End Function

