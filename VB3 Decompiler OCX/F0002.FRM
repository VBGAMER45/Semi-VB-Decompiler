VERSION 5.00
Begin VB.Form form3 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Open File"
   ClientHeight    =   3195
   ClientLeft      =   570
   ClientTop       =   2415
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form4"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3195
   ScaleWidth      =   4425
   Begin VB.CommandButton control1 
      Appearance      =   0  'Flat
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   372
      Left            =   3480
      TabIndex        =   5
      Top             =   2700
      Width           =   852
   End
   Begin VB.CommandButton control2 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2520
      TabIndex        =   4
      Top             =   2700
      Width           =   852
   End
   Begin VB.ComboBox control3 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.FileListBox control4 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1812
   End
   Begin VB.DirListBox control5 
      Appearance      =   0  'Flat
      Height          =   2616
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2292
   End
   Begin VB.DriveListBox control6 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2292
   End
End
Attribute VB_Name = "form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' F0002.FRM
Option Explicit

Private Sub control2_Click()
  form3.Tag = ""
  Hide
End Sub

Private Sub control1_Click()
Dim l001E As String
Dim l0020 As String
  l001E = control4.FileName
  If l001E = "" Then Beep: control4.SetFocus: Exit Sub
  l0020 = control4.Path
  If Right$(l0020, 1) <> "\" Then l0020 = l0020 & "\"
  form3.Tag = l0020 & l001E
  Hide
End Sub

Private Sub control3_Click()
Dim l002A As Integer
  l002A = control3.ListIndex
  If l002A < 0 Then Exit Sub
  control4.Pattern = control3.List(l002A)
End Sub

Private Sub control5_Change()
  control4.Path = control5.List(control5.ListIndex)
End Sub

Private Sub control5_Click()
  control4.Path = control5.List(control5.ListIndex)
End Sub

Private Sub control6_Change()
  control5.Path = control6.Drive
End Sub

Private Sub control4_DblClick()
  control1.Value = 1
End Sub

Private Sub Form_Activate()
Dim l0044 As String
Dim l0046 As Integer
  l0044 = Tag
  If Len(l0044) = 0 Then
    If control3.ListCount Then
      l0046 = control3.ListIndex
      If l0046 < 0 Then l0046 = 0: control3.ListIndex = 0
      l0044 = control3.List(l0046)
    End If
    If Len(l0044) = 0 Then l0044 = "*.*"
  End If
  Caption = gc03CE & l0044
  control4.Path = control5.List(control5.ListIndex)
  control4.Pattern = l0044
End Sub

Private Sub Form_Resize()
  If control3.ListCount Then
    control4.Top = control5.Top
    control4.Height = control1.Top - control4.Top
    control3.Visible = True
  End If
End Sub
