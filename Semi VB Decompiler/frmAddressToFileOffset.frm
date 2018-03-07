VERSION 5.00
Begin VB.Form frmAddressToFileOffset 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Address To File Offset Convertor"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameRVA 
      Caption         =   "RVA To File Pointer"
      Height          =   975
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtRVA 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "0"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblFileRva 
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox txtFileOffset 
      Height          =   285
      Left            =   4320
      TabIndex        =   6
      Text            =   "0"
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtMemoryOffset 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "0"
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtImageBase 
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Text            =   "4194304"
      ToolTipText     =   "The Image Base of the EXE file"
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblNote 
      Caption         =   "Note: These values are decimal format."
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   1200
      Width           =   3495
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
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblFileOffset 
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
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblMemoryOffset 
      Caption         =   "Memory Offset:"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmAddressToFileOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************
'frmAddressToFileOffset
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub txtFileOffset_Change()
    Dim dImageNum As Double
    Dim dFileNum As Double
    If IsNumeric(txtFileOffset.Text) = False Then txtFileOffset.Text = 0

    dImageNum = txtImageBase.Text
    dFileNum = txtFileOffset.Text
    txtMemoryOffset.Text = Trim$(Str$(dImageNum + dFileNum)) 'txtFileOffset.Text + txtImageBase.Text
End Sub

Private Sub txtImageBase_Change()
    If IsNumeric(txtImageBase.Text) = False Then txtImageBase.Text = 4194304
End Sub

Private Sub txtMemoryOffset_Change()
    Dim dImageNum As Double
    Dim dMemNum As Double
    dImageNum = txtImageBase.Text
    
    If IsNumeric(txtMemoryOffset.Text) = False Then txtMemoryOffset.Text = 0
    dMemNum = txtMemoryOffset.Text
    txtFileOffset.Text = Trim$(Str$(dMemNum - dImageNum))
End Sub

Private Sub txtRVA_Change()
    If IsNumeric(txtRVA.Text) = False Then txtRVA.Text = 0
    lblFileRva.Caption = "File Offset: " & GetPtrFromRVA2(txtRVA.Text)
    
End Sub


