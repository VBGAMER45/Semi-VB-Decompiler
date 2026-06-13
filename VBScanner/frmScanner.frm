VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmScanner
   Caption         =   "VB / .NET Software Scanner"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9135
   BeginProperty Font
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFolder
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton cmdBrowse
      Caption         =   "Browse..."
      Height          =   315
      Left            =   7020
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdScan
      Caption         =   "Scan"
      Height          =   315
      Left            =   7980
      TabIndex        =   2
      Top             =   120
      Width           =   1035
   End
   Begin VB.Frame fraTypes
      Caption         =   "File types"
      Height          =   675
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   2895
      Begin VB.CheckBox chkExe
         Caption         =   "EXE"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   280
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox chkDll
         Caption         =   "DLL"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   280
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox chkOcx
         Caption         =   "OCX"
         Height          =   255
         Left            =   1980
         TabIndex        =   6
         Top             =   280
         Value           =   1  'Checked
         Width           =   795
      End
   End
   Begin VB.Frame fraRuntime
      Caption         =   "Runtime to list"
      Height          =   675
      Left            =   3120
      TabIndex        =   7
      Top             =   540
      Width           =   5895
      Begin VB.CheckBox chkVB4
         Caption         =   "VB4"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   280
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox chkVB5
         Caption         =   "VB5"
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   280
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox chkVB6
         Caption         =   "VB6"
         Height          =   255
         Left            =   1980
         TabIndex        =   10
         Top             =   280
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox chkNet
         Caption         =   ".NET"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   280
         Value           =   1  'Checked
         Width           =   795
      End
   End
   Begin MSComctlLib.ListView lvResults
      Height          =   3375
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5953
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblFolder
      Caption         =   "Folder:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   160
      Width           =   615
   End
   Begin VB.Label lblDecLabel
      Caption         =   "Decompiler:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4840
      Width           =   915
   End
   Begin VB.TextBox txtDecompiler
      Height          =   315
      Left            =   1080
      TabIndex        =   13
      Top             =   4800
      Width           =   7335
   End
   Begin VB.CommandButton cmdBrowseExe
      Caption         =   "..."
      Height          =   315
      Left            =   8460
      TabIndex        =   14
      Top             =   4800
      Width           =   555
   End
   Begin VB.CommandButton cmdDecompile
      Caption         =   "Send to Semi VB Decompiler"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5220
      Width           =   2895
   End
   Begin VB.Label lblStatus
      Caption         =   "Ready."
      Height          =   315
      Left            =   3120
      TabIndex        =   16
      Top             =   5280
      Width           =   5895
   End
End
Attribute VB_Name = "frmScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************
' frmScanner - main window for VB Scanner
'
' Recursively scans a folder for EXE/DLL/OCX files, identifies the
' VB4/VB5/VB6/.NET ones (see modPE), lists them, and hands the
' selected file to Semi VB Decompiler on the command line.
'*********************************************
Option Explicit

Private mScanning As Boolean       ' a scan is running
Private mCancel As Boolean         ' user asked to stop
Private mScanned As Long           ' candidate files examined
Private mFound As Long             ' matches added to the list

Private Sub Form_Load()
    ' List view columns.
    With lvResults
        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        .HideSelection = False
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Name", 2600
        .ColumnHeaders.Add , , "Type", 800
        .ColumnHeaders.Add , , "Runtime", 1000
        .ColumnHeaders.Add , , "Folder", 4300
    End With

    ' Restore last folder.
    txtFolder.Text = IniGet("Paths", "Folder", "")

    ' Resolve the decompiler exe.
    Dim p As String
    p = IniGet("Paths", "Decompiler", "")
    If Len(p) = 0 Or Len(Dir$(p)) = 0 Then p = FindDecompiler()
    txtDecompiler.Text = p

    lblStatus.Caption = "Ready."
End Sub

' Try a couple of sensible default locations for SemiVBDecompiler.exe.
Private Function FindDecompiler() As String
    Dim cand As String
    cand = AddSlash(ParentFolder(App.Path)) & "SemiVBDecompiler.exe"
    If Len(Dir$(cand)) > 0 Then FindDecompiler = cand: Exit Function
    cand = AddSlash(App.Path) & "SemiVBDecompiler.exe"
    If Len(Dir$(cand)) > 0 Then FindDecompiler = cand: Exit Function
    FindDecompiler = ""
End Function

Private Sub cmdBrowse_Click()
    Dim s As String
    s = BrowseFolder(Me.hWnd, "Select the folder to scan")
    If Len(s) > 0 Then txtFolder.Text = s
End Sub

Private Sub cmdBrowseExe_Click()
    Dim s As String, startDir As String
    If Len(txtDecompiler.Text) > 0 Then startDir = ParentFolder(txtDecompiler.Text)
    s = BrowseForExe(Me.hWnd, startDir)
    If Len(s) > 0 Then
        txtDecompiler.Text = s
        IniPut "Paths", "Decompiler", s
    End If
End Sub

Private Sub txtDecompiler_Change()
    IniPut "Paths", "Decompiler", txtDecompiler.Text
End Sub

Private Sub cmdScan_Click()
    If mScanning Then
        mCancel = True
        Exit Sub
    End If

    Dim root As String
    root = Trim$(txtFolder.Text)
    If Len(root) = 0 Then
        MsgBox "Please choose a folder to scan.", vbInformation
        Exit Sub
    End If
    If Len(Dir$(root, vbDirectory)) = 0 Then
        MsgBox "Folder not found:" & vbCrLf & root, vbExclamation
        Exit Sub
    End If
    IniPut "Paths", "Folder", root

    ' Start.
    mScanning = True
    mCancel = False
    mScanned = 0
    mFound = 0
    lvResults.ListItems.Clear
    cmdScan.Caption = "Stop"
    SetControls False

    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then
        MsgBox "Scripting.FileSystemObject is unavailable on this system.", vbCritical
        EndScan
        Exit Sub
    End If

    ScanFolder root, fso

    EndScan
    If mCancel Then
        lblStatus.Caption = "Stopped. " & FoundSummary()
    Else
        lblStatus.Caption = "Done. " & FoundSummary()
    End If
End Sub

Private Function FoundSummary() As String
    FoundSummary = "Examined " & mScanned & " file(s), found " & mFound & " match(es)."
End Function

Private Sub EndScan()
    mScanning = False
    cmdScan.Caption = "Scan"
    SetControls True
End Sub

' Enable/disable inputs during a scan.
Private Sub SetControls(ByVal bEnabled As Boolean)
    txtFolder.Enabled = bEnabled
    cmdBrowse.Enabled = bEnabled
    chkExe.Enabled = bEnabled
    chkDll.Enabled = bEnabled
    chkOcx.Enabled = bEnabled
    chkVB4.Enabled = bEnabled
    chkVB5.Enabled = bEnabled
    chkVB6.Enabled = bEnabled
    chkNet.Enabled = bEnabled
    cmdDecompile.Enabled = bEnabled
End Sub

' Recursive directory walk using FSO (Dir$ is not re-entrant).
Private Sub ScanFolder(ByVal sFolder As String, ByVal fso As Object)
    Dim fld As Object, fil As Object, subf As Object
    On Error Resume Next
    Set fld = fso.GetFolder(sFolder)
    If fld Is Nothing Then Exit Sub

    For Each fil In fld.Files
        If mCancel Then Exit Sub
        ConsiderFile fil.Path
        mScanned = mScanned + 1
        If (mScanned Mod 20) = 0 Then
            lblStatus.Caption = "Scanning... " & FoundSummary()
            DoEvents
        End If
    Next

    For Each subf In fld.SubFolders
        If mCancel Then Exit Sub
        ScanFolder subf.Path, fso
    Next
End Sub

' Apply the type filter, identify the file, apply the runtime filter,
' and add a match to the list.
Private Sub ConsiderFile(ByVal sPath As String)
    Dim ext As String, dotPos As Long
    dotPos = InStrRev(sPath, ".")
    If dotPos = 0 Then Exit Sub
    ext = LCase$(Mid$(sPath, dotPos + 1))

    Select Case ext
        Case "exe": If chkExe.Value <> vbChecked Then Exit Sub
        Case "dll": If chkDll.Value <> vbChecked Then Exit Sub
        Case "ocx": If chkOcx.Value <> vbChecked Then Exit Sub
        Case Else: Exit Sub
    End Select

    Dim kind As Long
    kind = modPE.IdentifyFile(sPath)
    If kind = RT_NONE Then Exit Sub

    Select Case kind
        Case RT_VB4: If chkVB4.Value <> vbChecked Then Exit Sub
        Case RT_VB5: If chkVB5.Value <> vbChecked Then Exit Sub
        Case RT_VB6: If chkVB6.Value <> vbChecked Then Exit Sub
        Case RT_NET: If chkNet.Value <> vbChecked Then Exit Sub
    End Select

    AddResult sPath, UCase$(ext), RuntimeName(kind)
    mFound = mFound + 1
End Sub

Private Sub AddResult(ByVal sPath As String, ByVal sExt As String, ByVal sRuntime As String)
    Dim p As Long, nm As String, fld As String
    p = InStrRev(sPath, "\")
    If p > 0 Then
        nm = Mid$(sPath, p + 1)
        fld = Left$(sPath, p - 1)
    Else
        nm = sPath
        fld = ""
    End If

    Dim it As ListItem
    Set it = lvResults.ListItems.Add(, , nm)
    it.SubItems(1) = sExt
    it.SubItems(2) = sRuntime
    it.SubItems(3) = fld
    it.Tag = sPath
End Sub

' Double-click a row to send it to the decompiler.
Private Sub lvResults_DblClick()
    cmdDecompile_Click
End Sub

Private Sub cmdDecompile_Click()
    Dim it As ListItem
    Set it = lvResults.SelectedItem
    If it Is Nothing Then
        MsgBox "Select a file in the list first.", vbInformation
        Exit Sub
    End If

    Dim sFile As String, sExe As String
    sFile = it.Tag
    sExe = Trim$(txtDecompiler.Text)

    If Len(sExe) = 0 Or Len(Dir$(sExe)) = 0 Then
        MsgBox "Semi VB Decompiler was not found." & vbCrLf & _
               "Set its path with the '...' button.", vbExclamation
        Exit Sub
    End If
    If Len(Dir$(sFile)) = 0 Then
        MsgBox "The selected file no longer exists:" & vbCrLf & sFile, vbExclamation
        Exit Sub
    End If

    On Error GoTo shellErr
    Shell """" & sExe & """ """ & sFile & """", vbNormalFocus
    lblStatus.Caption = "Launched: " & it.Text
    Exit Sub
shellErr:
    MsgBox "Could not launch the decompiler:" & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mCancel = True
End Sub
