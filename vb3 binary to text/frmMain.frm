VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3D800911-77E3-43DE-82EA-7FC87C713180}#1.1#0"; "cPopMenu6.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB 1/2/3 Binary Form To Text Converter - VisualBasicZone.com"
   ClientHeight    =   6150
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8415
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtErrorLog 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   4680
      Width           =   8055
   End
   Begin RichTextLib.RichTextBox txtStorage 
      Height          =   1695
      Index           =   0
      Left            =   6000
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2990
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":014A
   End
   Begin VB.ListBox lstForms 
      Height          =   4155
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox txtOutput 
      Height          =   4155
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7329
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":01CC
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3720
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   4440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":024E
            Key             =   "NULL"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0568
            Key             =   "OPEN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0882
            Key             =   "SAVE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B9C
            Key             =   "PRINT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EB6
            Key             =   "CUT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11D0
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14EA
            Key             =   "PASTE"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1804
            Key             =   "FIND"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B1E
            Key             =   "HELP"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E38
            Key             =   "TICK"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2152
            Key             =   "FORM"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26EC
            Key             =   "PROJECT"
         EndProperty
      EndProperty
   End
   Begin cPopMenu6.PopMenu ctlPopMenu 
      Left            =   5160
      Top             =   0
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
   End
   Begin VB.Label lblErrorLog 
      Caption         =   "Error Log:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label lblFormName 
      Caption         =   "Form Name -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblForms 
      Caption         =   "Forms:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Form"
      End
      Begin VB.Menu mnuFileOpenProject 
         Caption         =   "Open &Project"
      End
      Begin VB.Menu mnuFileSaveAsText 
         Caption         =   "&Save Forms as Text"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsShowOffsets 
         Caption         =   "Show Gui Offsets"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOptionsExtractImages 
         Caption         =   "Extract Images"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOptionsProcessVBXControls 
         Caption         =   "Process VBX Controls"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuFileHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'***********************************
Option Explicit

'Needed for Browse For Folder
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
    End Type
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Dim strCurrentForm As String 'Hold The Current Form Name
Dim cEnd As Long 'Hold Control End Posistion
Dim m_lAboutId As Long 'Holds The About System Menu Id
Public Function ReturnControlEnd() As Long
'Return the file offset of the end of the current control
    ReturnControlEnd = cEnd
End Function

Private Sub ctlPopMenu_SystemMenuClick(ItemNumber As Long)
    Select Case ItemNumber
        Case m_lAboutId
            frmAbout.Show vbModal, Me
    End Select
End Sub

Private Sub Form_Load()
    Me.Tag = "vbgamer45"
    frmAbout.Tag = "http://www.visualbasiczone.com"
    'Ensure the ReadMe is always created
    Call modGlobals.PrintReadMe
    'Setup the custom menus
    With ctlPopMenu
        ' Add a new about item to the system menu:
        m_lAboutId = .SystemMenuAppendItem("&About...")
        .OfficeXpStyle = True
        
        ' Associate the image list:
        .ImageList = ilsIcons
        
        ' Parse through the VB designed menu and sub class the items:
        .SubClassMenu Me
        
        pSetIcon "FORM", "mnuFileOpen"
        pSetIcon "PROJECT", "mnuFileOpenProject"
        pSetIcon "SAVE", "mnuFileSaveAsText"

    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    End
End Sub

Private Sub lstForms_Click()
    If lstForms.ListIndex <> -1 Then
        Dim i As Integer
        For i = 0 To txtStorage.UBound
            If txtStorage(i).Tag = lstForms.Text Then
                lblFormName.Caption = "Form Name - " & lstForms.Text
                txtOutput.Text = "Version 2.00" & vbCrLf
                txtOutput.Text = txtOutput.Text & txtStorage(i).Text
            End If
        Next
    End If
End Sub

Private Sub mnuFileExit_Click()
        Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    On Error GoTo nofile:
        CD1.Filename = ""
        CD1.Filter = "VB Forms(*.frm)|*.frm|All Files(*.*)|*.*;"
        CD1.DialogTitle = "Select VB 1/2/3 Binary Form"
        CD1.ShowOpen
        If CD1.Filename = "" Then Exit Sub
        ReDim gVBXControlPath(0)
        Call ConvertVB3Form(CD1.Filename)
        Call modGlobals.DoFinalFormBuffer
        mnuFileSaveAsText.Enabled = True
    Exit Sub
nofile:
    MsgBox "Error_mnuFileOpen: " & Err.Description
End Sub

Private Sub mnuFileOpenProject_Click()
    On Error GoTo nofile:
        Dim i As Integer
        CD1.Filename = ""
        CD1.Filter = "Mak(*.mak)|*.mak|All Files(*.*)|*.*;"
        CD1.DialogTitle = "Select VB 2/3 Mak File"
        CD1.ShowOpen
        If CD1.Filename = "" Then Exit Sub
        'Clear the forms list
        lstForms.Clear
        'Clear the error log
        txtErrorLog.Text = ""
        txtOutput.Text = ""
        'Clear old storage buffers
        For i = 0 To txtStorage.UBound
            txtStorage(i).Text = vbNullString
            txtStorage(i).Tag = vbNullString
        Next
        For i = txtStorage.UBound To txtStorage.UBound + 1 Step -1
            Unload txtStorage(i)
        Next i
        Call OpenVBMakFile(CD1.Filename, CD1.FileTitle)
        Call modGlobals.DoFinalFormBuffer
        mnuFileSaveAsText.Enabled = True
        
    Exit Sub
nofile:
    MsgBox "Error_mnuFileOpenProject: " & Err.Description
End Sub


Private Sub mnuFileSaveAsText_Click()
On Error GoTo errHandle:
    Dim sPath As String
    Dim structFolder As BROWSEINFO
    Dim iNull As Integer
    Dim ret As Long
    Dim i As Long
    structFolder.hOwner = Me.hwnd
    structFolder.lpszTitle = "Browse for folder"
    structFolder.ulFlags = BIF_NEWDIALOGSTYLE  'To create make new folder option

    ret = SHBrowseForFolder(structFolder)
    If ret Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList ret, sPath
        'free the block of memory
        CoTaskMemFree ret
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    
    If sPath = vbNullString Then Exit Sub
    
    Dim F As Integer
    F = FreeFile
    If Right(sPath, 1) = "\" Then
        sPath = Left(sPath, Len(sPath) - 1)
    End If
    For i = 0 To txtStorage.UBound
        If txtStorage(i).Tag <> "" Then

            Open sPath & "\" & txtStorage(i).Tag & ".frm" For Output As #F
                Print #F, "Version 2.00"
                Print #F, txtStorage(i).Text
            Close #F
        End If
    Next
Exit Sub
errHandle:
    MsgBox "Error_mnuFileSaveAsText: " & Err.Description & " " & sPath
End Sub

Private Sub mnuHelpAbout_Click()
    'Show the about form
    frmAbout.Show vbModal, Me
End Sub
Public Sub ConvertVB3Form(ByVal strFilename As String, Optional MakFile As Boolean = False)
    On Error GoTo errHandle
    Dim F As Integer
    Dim frmHeader As FormHeader
    Dim strControlName As String
    Dim cHeader As ControlHeader
    Dim bType As Byte
    Dim i As Integer
    

    'F = FreeFile
    If MakFile = False Then
        'Clear the forms list
        lstForms.Clear
        'Clear the error log
        txtErrorLog.Text = ""
        txtOutput.Text = ""
        'Clear old storage buffers
        For i = 0 To txtStorage.UBound
            txtStorage(i).Text = vbNullString
            txtStorage(i).Tag = vbNullString
        Next
        For i = txtStorage.UBound To txtStorage.UBound + 1 Step -1
            Unload txtStorage(i)
        Next i
    End If
    
    'Setup the file class
    Set gVBFormFile = New clsFile
    gVBFormFile.Setup (strFilename)
    F = gVBFormFile.FileNumber
    'Reset Frx Address
    gFrxAddress = 0
    gFormDone = False
    'Get the form header
    Get #F, , frmHeader
    'Check if the file is a valid vb binary file
    If frmHeader.i1 <> -13057 And frmHeader.i1 <> -13277 Then
        MsgBox "This form is saved as Text! Does not need converting!", vbInformation
        Exit Sub
    End If
        
    'Main loop
    Do While gFormDone = False
        Dim cPos As Long 'Hold the control Position
        
        bFirstFF = False
        'Get control header for each control
        cPos = Loc(F)
        Get #F, , cHeader
        'Check if the control is a member of a control array
        If cHeader.IsArray = 128 Then
            Dim aHeader As ArrayControlHeader
            Get #F, cPos + 1, aHeader
            cHeader.ControlID = aHeader.ControlID
            cHeader.Length = aHeader.Length
            cHeader.NameLength = aHeader.NameLength
        End If
        'Get The control name
        strControlName = gVBFormFile.GetString(Loc(F), cHeader.NameLength, False)
        'Get the end of the current control offset
        cEnd = cPos + cHeader.Length
        Get #F, , bType
        
        'Check the control type
        Select Case bType
            
            Case vbForm
                'Hold the current Form name
                strCurrentForm = strControlName
                gIdentSpaces = 0
                'Load a new textbox to hold to the form's data
                Call modGlobals.LoadNewFormHolder(strCurrentForm)
                Call AddText("Begin Form " & strCurrentForm)
                gIdentSpaces = gIdentSpaces + 1
                lstForms.AddItem strCurrentForm
            Case vbLabel
                Call AddText("Begin Label " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
             Case vbCommandbutton
                Call AddText("Begin CommandButton " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
             Case vbMenu
                Call AddText("Begin Menu " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
             Case vbTextBox
                Call AddText("Begin TextBox " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
             Case vbFrame
                Call AddText("Begin Frame " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
             Case vbPictureBox
                Call AddText("Begin PictureBox " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
             Case vbCheckbox
                Call AddText("Begin CheckBox " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbOptionbutton
                Call AddText("Begin OptionButton " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbComboBox
                Call AddText("Begin ComboBox " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbListbox
                Call AddText("Begin ListBox " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbHscroll
                Call AddText("Begin Hscroll " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbVscroll
                Call AddText("Begin Vscroll " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbTimer
                Call AddText("Begin Timer " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbDriveListbox
                Call AddText("Begin DriveListBox  " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbDirectoryListbox
                Call AddText("Begin DirectoryListbox " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbFileListbox
                Call AddText("Begin FileListbox " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbShape
                Call AddText("Begin Shape " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbLine
                Call AddText("Begin Line " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbImage
                Call AddText("Begin Image " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
            Case vbMDIForm
                strCurrentForm = strControlName
                gIdentSpaces = 0
                Call modGlobals.LoadNewFormHolder(strCurrentForm)
                Call AddText("Begin MDIForm " & strControlName)
                lstForms.AddItem strCurrentForm
                gIdentSpaces = gIdentSpaces + 1
            Case Else
                'Add error to log if control type is not reconziged
                Call modGlobals.AddError("Error_Unknown Control Type! ControlType: " & bType)
                Call AddText("Begin " & strControlName)
                gIdentSpaces = gIdentSpaces + 1
                Seek #F, cEnd

        End Select

        Dim bOpcode As Byte
        Do Until (Loc(F) >= cEnd)
NextOpcode:
            Get #F, , bOpcode
            If bOpcode = 255 And bFirstFF = False And bType <> 255 Then
            
                bFirstFF = True
                GoTo NextOpcode
            End If
            Select Case bType
            
                Case vbForm
                    Call modForm.ProccessForm(F, bOpcode)
                Case vbCommandbutton
                    Call modCommandButton.ProccessCommandButton(F, bOpcode)
                Case vbLabel
                    Call modLabel.ProccessLabel(F, bOpcode)
                Case vbMenu
                    Call modMenu.ProccessMenu(F, bOpcode)
                Case vbTextBox
                    Call modTextBox.ProccessTextBox(F, bOpcode)
                Case vbFrame
                    Call modFrame.ProccessFrame(F, bOpcode)
                Case vbPictureBox
                    Call modPictureBox.ProccessPictureBox(F, bOpcode)
                Case vbCheckbox
                    Call modCheckBox.ProccessCheckBox(F, bOpcode)
                Case vbOptionbutton
                    Call modOptionbutton.ProccessOptionButton(F, bOpcode)
                Case vbComboBox
                    Call modComboBox.ProccessComboBox(F, bOpcode)
                Case vbListbox
                    Call modListBox.ProccessListBox(F, bOpcode)
                Case vbHscroll
                    Call modHscroll.ProccessHscroll(F, bOpcode)
                Case vbVscroll
                    Call modVscroll.ProccessVscroll(F, bOpcode)
                Case vbTimer
                    Call modTimer.ProccessTimer(F, bOpcode)
                Case vbDriveListbox
                    Call modDriveListBox.ProccessDriveListBox(F, bOpcode)
                Case vbDirectoryListbox
                    Call modDirectoryListbox.ProccessDirectoryListBox(F, bOpcode)
                Case vbFileListbox
                    Call modFileListbox.ProccessFileListBox(F, bOpcode)
                Case vbShape
                    Call modShape.ProccessShape(F, bOpcode)
                Case vbLine
                    Call modLine.ProccessLine(F, bOpcode)
                Case vbImage
                    Call modImage.ProccessImage(F, bOpcode)
                Case vbMDIForm
                    Call modMDIForm.ProccessMDIForm(F, bOpcode)
                Case Else
                    Call modCustom.ProccessCustom(F, bOpcode)
            End Select

        Loop
        
        
        
    Loop 'gFormDone
    If gIdentSpaces > 0 Then
        gIdentSpaces = 0
        Call AddText("End")
    End If
        
    Close #F
    
    
    Exit Sub
errHandle:
    MsgBox "Error_ConvertVB3Form: " & Err.Description

End Sub
Public Sub OpenVBMakFile(ByVal strFilename As String, ByVal strFileTitle As String)
    On Error GoTo errHandle
        'Clear the forms list
        lstForms.Clear
        'Clear the error log
        txtErrorLog.Text = ""
        'Clear old storage buffers
        Dim i As Integer
        For i = 0 To txtStorage.UBound
            txtStorage(i).Text = vbNullString
            txtStorage(i).Tag = vbNullString
        Next
        For i = txtStorage.UBound To txtStorage.UBound + 1 Step -1
            Unload txtStorage(i)
        Next i
        
        Dim F As Integer
        Dim strData As String
        Dim strFormList() As String
        ReDim strFormList(0)
        ReDim gVBXControlPath(0)
        F = FreeFile
        Open strFilename For Input As #F
            Do
                Line Input #F, strData
                'Get the Forms from the project
                If Right$(UCase$(strData), 4) = ".FRM" Then
                    strFormList(UBound(strFormList)) = strData
                    ReDim Preserve strFormList(UBound(strFormList) + 1)
                End If
                'Hold each VBX Control
                If Right$(UCase$(strData), 4) = ".VBX" Then
                    gVBXControlPath(UBound(gVBXControlPath)) = strData
                    ReDim Preserve gVBXControlPath(UBound(gVBXControlPath) + 1)
                End If
            Loop While Not EOF(F)
        Close #F
        
        Dim strFilePath As String
        For i = 0 To UBound(strFormList)
            If strFormList(i) <> "" Then
                strFilePath = Replace(strFilename, strFileTitle, "")
                Call ConvertVB3Form(strFilePath & strFormList(i), True)
            End If
        Next
    Exit Sub
errHandle:
    MsgBox "Error_OpenVBMakFile: " & Err.Description

End Sub
Public Sub AddText(ByVal strText As String)
'Add text to the current form's buffer
    If gIdentSpaces < 0 Then gIdentSpaces = 0
    strBuffer = strBuffer & Space(gIdentSpaces * 5) & strText & vbCrLf
End Sub

Private Sub mnuOptionsShowOffsets_Click()
    If mnuOptionsShowOffsets.Checked = True Then
        gShowOffsets = False
        mnuOptionsShowOffsets.Checked = False
    Else
        gShowOffsets = True
        mnuOptionsShowOffsets.Checked = True
    End If
End Sub
Public Function ReturnFormName() As String
    ReturnFormName = strCurrentForm
End Function
Private Sub pSetIcon( _
        ByVal sIconKey As String, _
        ByVal sMenuKey As String _
    )
Dim lIconIndex As Long
    lIconIndex = plGetIconIndex(sIconKey)
    ctlPopMenu.ItemIcon(sMenuKey) = lIconIndex
End Sub

Private Function plGetIconIndex( _
        ByVal sKey As String _
    ) As Long
    plGetIconIndex = ilsIcons.ListImages.Item(sKey).Index - 1
End Function


