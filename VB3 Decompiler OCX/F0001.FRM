VERSION 5.00
Begin VB.Form form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Select Project Path"
   ClientHeight    =   3990
   ClientLeft      =   870
   ClientTop       =   1530
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3990
   ScaleWidth      =   4200
   Begin VB.CommandButton control1 
      Appearance      =   0  'Flat
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   372
      Left            =   3000
      TabIndex        =   4
      Top             =   1380
      Width           =   1092
   End
   Begin VB.CommandButton control2 
      Appearance      =   0  'Flat
      Caption         =   "New"
      Height          =   372
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   1092
   End
   Begin VB.TextBox control3 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   2880
      TabIndex        =   2
      Top             =   180
      Width           =   1212
   End
   Begin VB.DirListBox control4 
      Appearance      =   0  'Flat
      Height          =   3264
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2652
   End
   Begin VB.DriveListBox control5 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   2652
   End
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' F0001.FRM
Option Explicit

Private Sub control1_Click()
  gv7B8C = control4.Path
  If Right$(gv7B8C, 1) <> "\" Then gv7B8C = gv7B8C & "\"
  Hide
End Sub

Private Sub control2_Click()
Dim l0022 As String
Dim l0024 As String
  l0022 = control3
  If Len(l0022) = 0 Then Beep: Exit Sub
  l0024 = control4.List(control4.ListIndex)
  If Right$(l0024, 1) <> "\" Then l0024 = l0024 & "\"
  l0024 = l0024 & l0022
On Error Resume Next
  MkDir l0024
  control4.Path = l0024
  If Err Then Beep: Exit Sub
  gv7B8C = l0024 & "\"
End Sub

Private Sub control6_Change()
  control4.Path = control5.Drive
End Sub


Private Sub Form_Load()

End Sub

