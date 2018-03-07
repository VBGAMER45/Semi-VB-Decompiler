VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCheckUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check for Update"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4695
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCheckUpdate 
      Caption         =   "Check for &Update"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3960
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblLatestVersion 
      Caption         =   "Latest Version:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblYourVersion 
      Caption         =   "Your Version:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Image imgUpdate 
      Height          =   480
      Left            =   360
      Picture         =   "frmCheckUpdate.frx":0000
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblNote 
      Caption         =   "Make sure you are connected to the internet then press the check for update button."
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmCheckUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************
'frmCheckUpdate
'Copyright VisualBasicZone.com 2004 - 2006
'*********************************************
Option Explicit

Private Sub cmdCheckUpdate_Click()
    If Inet1.StillExecuting = True Then Exit Sub
    Dim strData As String
    Dim strVersion As String
    strVersion = App.major & "." & App.minor & "." & App.revision
    strData = Inet1.OpenURL("http://www.visualbasiczone.com/products/semivbdecompiler/update.txt")

    If strData = "" Then
        MsgBox "You are not connected to the internet. Please connect first to check for an update.", vbInformation
        Exit Sub
    End If
    lblLatestVersion.Caption = "Latest Version: " & strData
    lblLatestVersion.Visible = True
    
    If strData = strVersion Then
        MsgBox "This product is up to date.", vbInformation
    Else
        MsgBox "There is an update for this product." & vbCrLf & "Please use the download link contained in your email to get the update.", vbExclamation
    End If
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.lblYourVersion.Caption = Me.lblYourVersion.Caption & " " & App.major & "." & App.minor & "." & App.revision
End Sub


