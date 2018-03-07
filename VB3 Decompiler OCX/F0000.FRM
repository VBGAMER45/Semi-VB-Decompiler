VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DoDi's VB3 Discompiler converted to VB6"
   ClientHeight    =   3975
   ClientLeft      =   150
   ClientTop       =   1575
   ClientWidth     =   9285
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
   Icon            =   "F0000.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3975
   ScaleWidth      =   9285
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ComboBox control1 
      Appearance      =   0  'Flat
      Height          =   2910
      Left            =   3240
      Style           =   1  'Simple Combo
      TabIndex        =   3
      Top             =   540
      Width           =   1812
   End
   Begin VB.CommandButton control2 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   252
      Left            =   8880
      TabIndex        =   7
      Top             =   300
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton control3 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   252
      Left            =   6720
      TabIndex        =   11
      Top             =   300
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton control4 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   252
      Left            =   4800
      TabIndex        =   10
      Top             =   300
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.ListBox control5 
      Appearance      =   0  'Flat
      Height          =   3180
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   3012
   End
   Begin VB.CommandButton control6 
      Appearance      =   0  'Flat
      Caption         =   "done"
      Height          =   252
      Left            =   3840
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.CommandButton control7 
      Appearance      =   0  'Flat
      Caption         =   "scan"
      Height          =   252
      Left            =   3240
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.ComboBox control8 
      Appearance      =   0  'Flat
      Height          =   2910
      Left            =   7080
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Top             =   540
      Width           =   2052
   End
   Begin VB.ComboBox control9 
      Appearance      =   0  'Flat
      Height          =   2910
      Left            =   5160
      Style           =   1  'Simple Combo
      TabIndex        =   4
      Top             =   540
      Width           =   1812
   End
   Begin VB.Label lblCredits 
      Caption         =   "Based off DoDi's Discompiler"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblFolderPath 
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   3600
      Width           =   7335
   End
   Begin VB.Label control10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Controls"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   7080
      TabIndex        =   2
      Top             =   300
      Width           =   1812
   End
   Begin VB.Label control11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subroutines"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   5160
      TabIndex        =   1
      Top             =   300
      Width           =   1524
   End
   Begin VB.Label control12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Modules"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   3240
      TabIndex        =   0
      Top             =   300
      Width           =   1344
   End
   Begin VB.Menu control13 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu control14 
         Caption         =   "&Open..."
         Visible         =   0   'False
      End
      Begin VB.Menu control15 
         Caption         =   "Im&prove"
         Enabled         =   0   'False
      End
      Begin VB.Menu control16 
         Caption         =   "&Combine Forms"
         Enabled         =   0   'False
      End
      Begin VB.Menu control17 
         Caption         =   "Start &VB"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu control18 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu control19 
         Caption         =   "&Info..."
         Visible         =   0   'False
      End
      Begin VB.Menu control20 
         Caption         =   "E&xit"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' F0000.FRM
Option Explicit
Dim m001A As Integer
Const mc0056 = 32 ' &H20%
Const mc0060 = 16 ' &H10%
Const mc0068 = 8 ' &H8%

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub control1_Change()
  control9.Clear
  sub1070 control1
End Sub

Private Sub control14_Click()
' Const mc0068 = 8 ' &H8%
  control16.Enabled = False
  gv7A1C = 0
  sub0E86 mc0060, Me
  If gv7A1C And mc0060 Then
    control15.Enabled = True
    control16.Enabled = True
  End If
End Sub
Public Sub DoDecompile()
  control16.Enabled = False
  gv7A1C = 0
  sub0E86 mc0060, Me
  If gv7A1C And mc0060 Then
    control15.Enabled = True
    control16.Enabled = True
  End If
End Sub
Private Sub control15_Click()
' Const mc0068 = 8 ' &H8%
  gv7A1C = gv7A1C And Not (mc0060 Or mc0068)
  sub0E86 mc0060, Me
End Sub

Private Sub control16_Click()
' Const mc0068 = 8 ' &H8%
  sub0E86 mc0056, Me
End Sub

Private Sub control19_Click()
' Const mc0068 = 8 ' &H8%
    frmAbout.Show
End Sub

Private Sub control20_Click()
' Const mc0068 = 8 ' &H8%
  'End
End Sub

Private Sub control3_Click()
  control9.AddItem fn0D70(gv7B2A)
  DoEvents
End Sub

Private Sub control4_Click()
  control1.Text = fn0D20(gv7AB6)
  DoEvents
End Sub

Private Sub control6_Click()
  m001A = False
End Sub

Private Sub control7_Click()
  m001A = 1
  Caption = gv0882
  control1.Clear
  control9.Clear
  control8.Clear
  DoEvents
End Sub

Private Sub control8_Change()
  sub1070 control8
End Sub

Private Sub control9_Change()
  sub1070 control9
End Sub

Private Sub Form_Load()
Dim l0048  As Integer
' Const mc0056 = 32 ' &H20%
' Const mc0060 = 16 ' &H10%
' Const mc0068 = 8 ' &H8%
  lblFolderPath.Caption = "Check this path for files generated: " & m_OutputFolder

  Set gv0BC0 = Me
  'Show
  Hide
  sub0201 l0048, gc03A6
  m001A = 1
  sub0EAF l0048
  sub0220 l0048
  Hide
End Sub



Private Sub sub1070(p0044 As ComboBox)
Static m0046 As Integer
  If m0046 Then Exit Sub
  m0046 = True
  If m001A Then
    p0044.AddItem p0044.Text
    p0044.ListIndex = p0044.NewIndex
    DoEvents
  Else
    Beep
  End If
  m0046 = False
End Sub


