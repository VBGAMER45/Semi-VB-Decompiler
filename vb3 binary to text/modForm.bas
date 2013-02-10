Attribute VB_Name = "modForm"
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'***********************************
Option Explicit
Private Type FormSizeType
    ClientLeft As Long
    ClientTop As Long
    ClientWidth As Long
    ClientHeight As Long
End Type
Dim FormSize As FormSizeType
Global gFrxAddress As Long
Public Function ProccessForm(ByVal F As Integer, ByVal Opcode As Byte) As Long
    On Error GoTo errHandle
        Dim bData As Byte
        Dim iData As Integer
        Dim lSize As Long
        Dim bArray() As Byte
    Select Case Opcode
        Case 0 'Caption
            Get #F, , bData
            Call frmMain.AddText("Caption  = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 3 'BackColor
            Call frmMain.AddText("BackColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 4 'ForeColor
            Call frmMain.AddText("ForeColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 5
        'Size Opcode
            Get #F, , FormSize
            Call frmMain.AddText("ClientHeight = " & FormSize.ClientHeight)
            Call frmMain.AddText("ClientLeft = " & FormSize.ClientLeft)
            Call frmMain.AddText("ClientTop = " & FormSize.ClientTop)
            Call frmMain.AddText("ClientWidth = " & FormSize.ClientWidth)
        Case 9
            Get #F, , bData
            Call frmMain.AddText("Enabled = " & bData)
        Case 10 'WindowState
             Get #F, , bData
             Call frmMain.AddText("WindowState = " & bData)
        Case 11 'MousePointer
            Get #F, , bData
            Call frmMain.AddText("MousePointer = " & bData)
        Case 12 'FONT
            Call modGlobals.GetFontProperty(F)
        Case 25 'Scale Mode
            Get #F, , iData
            If iData <> 1 Then
                Call frmMain.AddText("ScaleMode  = " & iData)
            End If
        Case 27
            Get #F, , bData
            Call frmMain.AddText("DrawStyle = " & bData)
        Case 28
            Get #F, , iData
            Call frmMain.AddText("DrawWidth = " & iData)
        Case 29 'FillStyle
            Get #F, , bData
            Call frmMain.AddText("FillStyle = " & bData)
        Case 30
            Call frmMain.AddText("FillColor = " & gVBFormFile.GetLong(Loc(F)))
        Case 31
            Get #F, , bData
            Call frmMain.AddText("DrawMode = " & bData)
        Case 33 'Picture
            lSize = gVBFormFile.GetLong(Loc(F))
            If lSize <> -1 Then
                
                ReDim bArray(lSize)
                Get #F, , bArray
                Seek #F, Loc(F)
                Call frmMain.AddText("Picture = " & frmMain.ReturnFormName() & ".frx:" & PadHex(Hex(gFrxAddress), 4))
                gFrxAddress = gFrxAddress + lSize
            End If
        Case 34
            Get #F, , bData
            Call frmMain.AddText("BorderStyle = " & bData)
        Case 35 'Icon

            lSize = gVBFormFile.GetLong(Loc(F))
            If lSize <> -1 Then
                
                ReDim bArray(lSize)
                Get #F, , bArray
                Seek #F, Loc(F)
                Call frmMain.AddText("Icon = " & frmMain.ReturnFormName() & ".frx:" & PadHex(Hex(gFrxAddress), 4))
                gFrxAddress = gFrxAddress + lSize
                
            End If
            
        Case 36 'LinkTopic
            Get #F, , bData
            Call frmMain.AddText("LinkTopic  = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 37 'Link Mode
            Get #F, , bData
            Call frmMain.AddText("LinkMode = " & bData)
        Case 38 'MaxButton
            Get #F, , bData
            Call frmMain.AddText("MaxButton = 0")
        Case 39 'MinButton
            Get #F, , bData
            Call frmMain.AddText("MinButton = 0")
        Case 40 'ControlBox
            Get #F, , bData
            Call frmMain.AddText("ControlBox = -1")
        Case 46 'Visible
            Get #F, , bData
            Call frmMain.AddText("Visible = 0")
        Case 47 'Tag
            Get #F, , bData
            Call frmMain.AddText("Tag  = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 48 'MDIChild
            Get #F, , bData
            Call frmMain.AddText("MDIChild = -1")
        Case 49
            Get #F, , bData
            Call frmMain.AddText("KeyPreview = -1")
        Case 50
            Get #F, , bData
            Call frmMain.AddText("ClipControls = " & bData)
        Case 51 'HelpContext
            Call frmMain.AddText("HelpContextID = " & gVBFormFile.GetLong(Loc(F)))
        Case 53 'Left
            gVBFormFile.GetLong (Loc(F))
        Case 54 'Top
            gVBFormFile.GetLong (Loc(F))
        Case 55 'Width
            gVBFormFile.GetLong (Loc(F))
        Case 56 'Height
            gVBFormFile.GetLong (Loc(F))
        Case 64 'FontTransparent
            Get #F, , bData
            Call frmMain.AddText("FontTransparent = " & bData)
        Case 66
            Get #F, , bData
        Case 96
            Get #F, , bData
        Case 98
            Call frmMain.AddText("AutoRedraw = " & GetTrueFalse(F))
        Case 255
            Get #F, , bData
            If bData = 4 Then
                gIdentSpaces = gIdentSpaces - 1
                Call frmMain.AddText("End")
                gFormDone = True
            End If
        Case Else
            
            Call AddError("Error_Unknown Opcode_ProcessForm: Opcode: " & Opcode & " Offset: " & Loc(F))
            bFirstFF = True
            Seek F, frmMain.ReturnControlEnd
    End Select
    
Exit Function
errHandle:
    Call AddError("Error_ProcessForm: Opcode: " & Opcode & " Offset: " & Loc(F) & " Description= " & Err.Description)
            
End Function
