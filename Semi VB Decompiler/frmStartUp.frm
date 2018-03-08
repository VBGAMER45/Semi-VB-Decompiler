VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStartUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StartUp Form Patcher"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2040
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboFormSwap 
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ListBox lstForms 
      Height          =   2595
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmdSavePatch 
      Caption         =   "Save Patched File"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblWarning 
      Caption         =   "Keep backups just in case!"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   4200
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Current Form:"
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
      Left            =   2160
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblCurrentForm 
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblFormList 
      Caption         =   "Form List"
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
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblObjectID 
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Swap Form Info with:"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblInfo2 
      Caption         =   "ObjectID"
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
      Left            =   2160
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblFormPointer 
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Form Pointer"
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
      Left            =   2160
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblStartupForm 
      Caption         =   "StartUp Form ="
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   4455
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pGuiTable() As tGuiTable
Dim CurrentFormIndex As Long
Dim NewFormIndex As Long
Private Sub cboFormSwap_Change()

    If lblCurrentForm.Caption = "" Then
        MsgBox "You need to select a form from the list", vbInformation
        cboFormSwap.Text = ""
        Exit Sub
    End If

    If cboFormSwap.Text <> "" Then
        Dim i As Integer
        For i = 0 To lstForms.ListCount - 1
            If lstForms.List(i) = cboFormSwap.Text Then
                NewFormIndex = i
                Exit For
            End If
        Next
    
    End If
End Sub

Private Sub cboFormSwap_Click()

    If lblCurrentForm.Caption = "" Then
        MsgBox "You need to select a form from the list", vbInformation
        cboFormSwap.Text = ""
        Exit Sub
    End If

    If cboFormSwap.Text <> "" Then
        Dim i As Integer
        For i = 0 To lstForms.ListCount - 1
            If lstForms.List(i) = cboFormSwap.Text Then
                NewFormIndex = i
                Exit For
            End If
        Next
    
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSavePatch_Click()

    If CurrentFormIndex = -1 Or NewFormIndex = -1 Then
        MsgBox "No changes made...", vbInformation
    Else
        CD1.DialogTitle = "Save As"
        CD1.Filename = vbNullString
        CD1.Filter = "VB Files(*.exe,*.ocx,*.dll)|*.exe;*.ocx;*.dll"
        CD1.ShowSave
        
        If CD1.Filename = vbNullString Then Exit Sub
        On Error Resume Next
        'Copy the exe to the temp directory
        Close
        FileCopy SFilePath, App.Path & "\dump\" & SFile & "\" & SFile
        Dim F As Long
        F = FreeFile
        Dim oFormPointer As Long
        Dim oFormObjectID As Long
        
        'Make the changes
        oFormPointer = pGuiTable(CurrentFormIndex).aFormPointer
        oFormObjectID = pGuiTable(CurrentFormIndex).lObjectID
        pGuiTable(CurrentFormIndex).aFormPointer = pGuiTable(NewFormIndex).aFormPointer
        pGuiTable(CurrentFormIndex).lObjectID = pGuiTable(NewFormIndex).lObjectID
        'pGuiTable(NewFormIndex).aFormPointer = oFormPointer
        pGuiTable(NewFormIndex).lObjectID = oFormObjectID
        
        Open App.Path & "\dump\" & SFile & "\" & SFile For Binary Access Write As F
                Seek F, gVBHeader.aGuiTable - OptHeader.ImageBase + 1
                Put #F, , pGuiTable
        Close #F
        'Save the file
        FileCopy App.Path & "\dump\" & SFile & "\" & SFile, CD1.Filename
        'Kill the temp file
        Kill App.Path & "\dump\" & SFile & "\" & SFile
        MsgBox "File Patched", vbInformation
    End If
End Sub

Private Sub Form_Load()
On Error GoTo nofile:
'Load Form GUI Objects
    lblStartupForm.Caption = "StartUp Form = " & AppData.StartUpName
    Dim F As Long
    F = FreeFile
    CurrentFormIndex = -1
    NewFormIndex = -1
    Open SFilePath For Binary Access Read As #F
        Seek F, gVBHeader.aGuiTable - OptHeader.ImageBase + 1
        ReDim pGuiTable(gVBHeader.FormCount - 1)
        Get #F, , pGuiTable
    
    
    'Get the Form Names
        Dim i As Long
        Dim cControlHeader As ControlHeader
        For i = 0 To UBound(pGuiTable)
            Seek F, pGuiTable(i).aFormPointer + 94 - OptHeader.ImageBase
            Get #F, , cControlHeader
            lstForms.AddItem cControlHeader.cName
            cboFormSwap.AddItem cControlHeader.cName
            
        Next
    Close F
Exit Sub
nofile:
    MsgBox "Error_frmStartUp_Load: " & err.Description, vbExclamation
End Sub

Private Sub lstForms_Click()
    If lstForms.ListIndex <> -1 Then
        lblObjectID.Caption = pGuiTable(lstForms.ListIndex).lObjectID
        lblFormPointer.Caption = pGuiTable(lstForms.ListIndex).aFormPointer
        lblCurrentForm.Caption = lstForms.Text
        CurrentFormIndex = lstForms.ListIndex
    End If
End Sub


