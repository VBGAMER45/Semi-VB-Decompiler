VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmAdvDecompile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced Decompile"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAdvDecompile 
      Caption         =   "Decompile program by offset.  For VB5/6"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3240
         ScaleHeight     =   375
         ScaleWidth      =   1455
         TabIndex        =   9
         Top             =   360
         Width           =   1455
         Begin VB.CommandButton cmdSelectFile 
            Caption         =   "&Select File"
            Height          =   375
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3480
         ScaleHeight     =   375
         ScaleWidth      =   1335
         TabIndex        =   7
         Top             =   2040
         Width           =   1335
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            Default         =   -1  'True
            Height          =   375
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1335
         End
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   2880
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtFileOffset 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "0"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtImageBase 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Text            =   "4194304"
         ToolTipText     =   "The Image Base of the EXE file"
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "If you get not a VB 4/5/6 file error try to enter +-1 of the file offset."
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Notes: Works at the offset of where VB5! is located.  This can be used on upx packed exe's if you dump the memory image to a file."
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "File Offset:"
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
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblImageBase 
         Caption         =   "Image Base:"
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
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAdvDecompile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************
'frmAdvDecompile
'Copyright VisualBasicZone.com 2004 - 2006
'*********************************************
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSelectFile_Click()
    If txtFileOffset.Text = 0 Then
        MsgBox "You need to enter the file offset first before selecting the Visual Basic File.", vbInformation
        Exit Sub
    End If
    
    CD1.Filename = vbNullString
    CD1.DialogTitle = "Select VB 4/5/6 exe"
    CD1.Filter = "All Files(*.*)|*.*;*.ocx;*.dll|All Files(*.*);"
    CD1.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
    CD1.ShowOpen
    
    If CD1.Filename = vbNullString Then Exit Sub
    
    If FileExists(CD1.Filename) = True Then
        Me.Hide
        If VerifyVBSig(CD1.Filename, txtFileOffset.Text) = True Then
            Call frmMain.OpenVBExe(CD1.Filename, CD1.FileTitle, True, txtImageBase.Text, txtFileOffset.Text)
            Unload Me
        
        Else
            MsgBox "Not a Valid VB5! offset", vbExclamation
        End If
    Else
        MsgBox "File Does not exist.", vbExclamation
    End If
    
End Sub
Private Function VerifyVBSig(ByVal strFile As String, Start As Long) As Boolean
On Error GoTo errHandle:
    Dim F As Long
    Dim vbSig As String * 4
    F = FreeFile
    Open strFile For Binary As #F
        Get #F, Start + 1, vbSig
        
    Close #F
    If vbSig = "VB5!" Then
        VerifyVBSig = True
    Else
        VerifyVBSig = False
    End If
Exit Function
errHandle:
MsgBox "Error_frmAdvDecompile_VerifyVBSig: " & err.Number & " " & err.Description
End Function


Private Sub txtFileOffset_Change()
    If IsNumeric(txtFileOffset.Text) = False Then txtFileOffset.Text = 0
    If txtFileOffset.Text < 0 Then txtFileOffset.Text = 0
End Sub

Private Sub txtImageBase_Change()
    If IsNumeric(txtImageBase.Text) = False Then txtImageBase.Text = 4194304

End Sub


