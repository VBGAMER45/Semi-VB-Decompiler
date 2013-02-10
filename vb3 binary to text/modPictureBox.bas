Attribute VB_Name = "modPictureBox"
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'***********************************
Option Explicit

Public Function ProccessPictureBox(ByVal F As Integer, ByVal Opcode As Byte) As Long
    On Error GoTo errHandle
        Dim bData As Byte
        Dim iData As Integer
        Dim lSize As Long
        Dim bArray() As Byte
    Select Case Opcode
        Case 0 'Caption
            Get #F, , bData
            Call frmMain.AddText("Caption  = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 1
            Call frmMain.AddText("BackColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 2 'Index
            Get #F, , iData
            Call frmMain.AddText("Index = " & iData)
        Case 3 'Picture
            lSize = gVBFormFile.GetLong(Loc(F))
            If lSize <> -1 Then
                
                ReDim bArray(lSize)
                Get #F, , bArray
                Seek #F, Loc(F)
                Call frmMain.AddText("Picture = " & frmMain.ReturnFormName() & ".frx:" & PadHex(Hex(gFrxAddress), 4))
                gFrxAddress = gFrxAddress + lSize
            End If
        
        Case 4
            Call frmMain.AddText("ForeColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 5
            Call modGlobals.GetControlSize(F)
        Case 9
            Get #F, , bData
            Call frmMain.AddText("Enabled = 0")
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
        Case 26 'ScaleMode
            Get #F, , iData
            If iData <> 1 Then
                Call frmMain.AddText("ScaleMode = " & iData)
            End If
        Case 28
            Get #F, , bData
            Call frmMain.AddText("DrawStyle = " & bData)
        Case 29
            Get #F, , iData
            Call frmMain.AddText("DrawWidth = " & iData)
        Case 30
            Get #F, , bData
            Call frmMain.AddText("FillStyle = " & bData)
        Case 31
            Call frmMain.AddText("FillColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 32
            Get #F, , bData
            Call frmMain.AddText("DrawMode = " & bData)
        Case 35
            Get #F, , bData
            Call frmMain.AddText("AutoSize = -1")
        Case 36
            Get #F, , bData
            Call frmMain.AddText("BorderStyle = " & bData)
        Case 38
            Get #F, , bData
            Call frmMain.AddText("LinkItem = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 42 'DragMode
            Get #F, , bData
            Call frmMain.AddText("DragMode = " & bData)
        Case 43
            lSize = gVBFormFile.GetLong(Loc(F))
            If lSize <> -1 Then
                
                ReDim bArray(lSize)
                Get #F, , bArray
                Seek #F, Loc(F)
                Call frmMain.AddText("DragIcon = " & frmMain.ReturnFormName() & ".frx:" & PadHex(Hex(gFrxAddress), 4))
                gFrxAddress = gFrxAddress + lSize
            End If
        Case 44
            Get #F, , iData
            Call frmMain.AddText("LinkTimeout = " & iData)
        Case 45
            Get #F, , bData
            Call frmMain.AddText("TabStop = 0")
        Case 46
            Get #F, , bData
            Call frmMain.AddText("Tag  = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 48
            Get #F, , bData
            Call frmMain.AddText("ClipControls = " & bData)
        Case 49
            Call frmMain.AddText("HelpContextID = " & gVBFormFile.GetLong(Loc(F)))
        Case 50 'Align
            Get #F, , bData
            Call frmMain.AddText("Align = " & bData)
       Case 52 'DataSource
            Get #F, , bData
            Call frmMain.AddText("DataSource = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 53 'DataField
            Get #F, , bData
            Call frmMain.AddText("DataField = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 64
            Get #F, , bData
            Call frmMain.AddText("FontTransparent = -1")
        Case 66
            Get #F, , bData
        Case 98
            Get #F, , bData
            Call frmMain.AddText("AutoRedraw = -1")
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
            
            Call AddError("Error_Unknown Opcode_ProcessPictureBox: Opcode: " & Opcode & " Offset: " & Loc(F))
            bFirstFF = True
            Seek F, frmMain.ReturnControlEnd
    End Select
    
Exit Function
errHandle:
    Call AddError("Error_ProcessPictureBox: Opcode: " & Opcode & " Offset: " & Loc(F) & " Description= " & Err.Description)
            
End Function

