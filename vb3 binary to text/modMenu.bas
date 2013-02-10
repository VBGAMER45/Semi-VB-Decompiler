Attribute VB_Name = "modMenu"
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'***********************************
Option Explicit
Dim IdentMenu As Boolean

Public Function ProccessMenu(ByVal F As Integer, ByVal Opcode As Byte) As Long
    On Error GoTo errHandle
        Dim bData As Byte
        Dim iData As Integer
    Select Case Opcode
        Case 1 'Index
            Get #F, , iData
            Call frmMain.AddText("Index = " & iData)
        Case 2 'Visible
            Get #F, , bData
            Call frmMain.AddText("Visible = " & bData)
        Case 3 'Caption
            Get #F, , bData
            Call frmMain.AddText("Caption  = " & Chr(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr(34))
        Case 4 'Checked
            Get #F, , bData
            Call frmMain.AddText("Checked = -1")
        Case 5 'Enabled
            Get #F, , bData
            Call frmMain.AddText("Enabled = " & bData)
        Case 7
            Get #F, , bData
            IdentMenu = True
            'gIdentSpaces = gIdentSpaces + 1

        Case 8 'Shortcut
            Get #F, , iData
            Call frmMain.AddText("ShortCut = " & " ^" & Chr(64 + iData))
        Case 11 'WindowList
            Get #F, , bData
            Call frmMain.AddText("WindowList = -1")
        Case 12 'HelpContextID
            Call frmMain.AddText("HelpContextID = " & gVBFormFile.GetLong(Loc(F)))
        Case 255
           Dim iCounter As Integer
           Static iMenu As Integer
           iCounter = 0
            Do
                Get #F, , bData

                If bData = 1 Then
                   
                ElseIf bData = 4 Then
                    gIdentSpaces = gIdentSpaces - 1
                    Call frmMain.AddText("End")
                    gFormDone = True
                ElseIf bData = 3 Then
                    Dim l As Long
                    For l = 1 To iMenu
                        gIdentSpaces = gIdentSpaces - 1
                        Call frmMain.AddText("End")
                        
                    Next
                    iMenu = 0
                ElseIf bData = 2 Then
                
                'MsgBox iMenu
                    If IdentMenu = True Then
                        iMenu = iMenu + 1
                       gIdentSpaces = gIdentSpaces + 1
                        IdentMenu = False
                    Else
                        gIdentSpaces = gIdentSpaces - 1
                        Call frmMain.AddText("End")
                    End If
                    'If iCounter = 1 Then
                        'gIdentSpaces = gIdentSpaces - 1
                        'Call frmMain.AddText("End")
                    'Else
                    '     gIdentSpaces = gIdentSpaces + 1
                    'End If
                    
                 End If

              iCounter = iCounter + 1
            Loop While bData <> 0 And bData < 6
            Seek F, Loc(F)
        Case Else
            
            Call AddError("Error_Unknown Opcode_ProcessMenu: Opcode: " & Opcode & " Offset: " & Loc(F))
            bFirstFF = True
            Seek F, frmMain.ReturnControlEnd
    End Select
    
Exit Function
errHandle:
    Call AddError("Error_ProcessMenu: Opcode: " & Opcode & " Offset: " & Loc(F) & " Description= " & Err.Description)
            
End Function
