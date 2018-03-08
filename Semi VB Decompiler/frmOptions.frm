VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Tag             =   "                                  vbgamer45"
   Begin MSComDlg.CommonDialog cdColor 
      Left            =   240
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Decompile Options"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNotes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkPCODE"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkShowColors"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkShowOffsets"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkDumpControls"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkSkipCOM"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkDisableNative"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Syntax Coloring"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdChangeComment"
      Tab(1).Control(1)=   "cmdChangeStrings"
      Tab(1).Control(2)=   "cmdChangeKeyword"
      Tab(1).Control(3)=   "lblComments"
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(5)=   "lblStrings"
      Tab(1).Control(6)=   "Label1"
      Tab(1).Control(7)=   "lblKeyword"
      Tab(1).Control(8)=   "lbl1"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "P-Code Options"
      TabPicture(2)   =   "frmOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkShowPCodeStringAddress"
      Tab(2).Control(1)=   "chkDisplayHex"
      Tab(2).Control(2)=   "chkDisplayAddress"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Native Options"
      TabPicture(3)   =   "frmOptions.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chkNativeStrings"
      Tab(3).Control(1)=   "chkNativeHex"
      Tab(3).Control(2)=   "chkNativeOffsets"
      Tab(3).ControlCount=   3
      Begin VB.CheckBox chkNativeStrings 
         Caption         =   "Show Native Strings."
         Enabled         =   0   'False
         Height          =   255
         Left            =   -74520
         TabIndex        =   23
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CheckBox chkNativeHex 
         Caption         =   "Show Native Disassembly Hex information."
         Height          =   255
         Left            =   -74520
         TabIndex        =   22
         Top             =   1080
         Width           =   3735
      End
      Begin VB.CheckBox chkNativeOffsets 
         Caption         =   "Show Native Disassembly Offsets."
         Height          =   375
         Left            =   -74520
         TabIndex        =   21
         Top             =   600
         Width           =   2895
      End
      Begin VB.CheckBox chkDisableNative 
         Caption         =   "Disable Native Disassembly"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2760
         Width           =   2655
      End
      Begin VB.CheckBox chkShowPCodeStringAddress 
         Caption         =   "Show P-Code String Address in the P-Code String List."
         Height          =   375
         Left            =   -74520
         TabIndex        =   19
         Top             =   1560
         Width           =   4335
      End
      Begin VB.CheckBox chkDisplayHex 
         Caption         =   "Display Hex information in P-Code output."
         Height          =   375
         Left            =   -74520
         TabIndex        =   18
         Top             =   1080
         Width           =   3615
      End
      Begin VB.CheckBox chkDisplayAddress 
         Caption         =   "Display Memory Address on P-Code output."
         Height          =   375
         Left            =   -74520
         TabIndex        =   17
         Top             =   600
         Width           =   3735
      End
      Begin VB.CommandButton cmdChangeComment 
         Caption         =   "Change"
         Height          =   255
         Left            =   -72360
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdChangeStrings 
         Caption         =   "Change"
         Height          =   255
         Left            =   -72360
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdChangeKeyword 
         Caption         =   "Change"
         Height          =   255
         Left            =   -72360
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox chkSkipCOM 
         Caption         =   "Skip COM and Control/Form Property Processing"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Disables processing of form and control properties."
         Top             =   480
         Width           =   3855
      End
      Begin VB.CheckBox chkDumpControls 
         Caption         =   "Dump Control/Form raw binary data"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Dumps the raw binary data of each control and form."
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox chkShowOffsets 
         Caption         =   "Show Offests and Gui Opcodes"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Shows file offsets of where certain information is contained in the exe."
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CheckBox chkShowColors 
         Caption         =   "Show Colors"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Shows syntax coloring for forms, modules, and classes"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox chkPCODE 
         Caption         =   "Disable P-Code Decompile"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Disables P-Code Analysis the setting is saved."
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label lblNotes 
         Caption         =   $"frmOptions.frx":0070
         Height          =   855
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   4455
      End
      Begin VB.Label lblComments 
         Caption         =   "'vbgamer45"
         Height          =   255
         Left            =   -73680
         TabIndex        =   14
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Comments:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   13
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblStrings 
         Caption         =   """Hello World"""
         Height          =   255
         Left            =   -73680
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Strings:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   10
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblKeyword 
         Caption         =   "Public Sub"
         Height          =   255
         Left            =   -73680
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl1 
         Caption         =   "Keywords:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   2535
      TabIndex        =   0
      Top             =   4560
      Width           =   1815
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************
'frmOptions
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************
Option Explicit

Private Sub chkDisplayAddress_Click()
    If chkDisplayAddress.value = vbChecked Then
        PCODEDisplayAddress = True
        Call SaveSetting("VB Decompiler", "Options", "ShowPCODEAddress", "TRUE")
    Else
        PCODEDisplayAddress = False
        Call SaveSetting("VB Decompiler", "Options", "ShowPCODEAddress", "FALSE")
    End If
End Sub

Private Sub chkDisplayHex_Click()
    If chkDisplayHex.value = vbChecked Then
        PCodeDisplayHexInfo = True
        Call SaveSetting("VB Decompiler", "Options", "ShowPCODEHex", "TRUE")
    
    Else
        PCodeDisplayHexInfo = False
        Call SaveSetting("VB Decompiler", "Options", "ShowPCODEHex", "FALSE")
    
    End If
End Sub

Private Sub chkDumpControls_Click()
    If chkDumpControls.value = vbChecked Then
        gDumpData = True
    Else
        gDumpData = False
    End If
End Sub

Private Sub chkNativeHex_Click()
    If chkNativeHex.value = vbChecked Then
        NativeShowHexInformation = True
    Else
        NativeShowHexInformation = False
    End If
End Sub

Private Sub chkNativeOffsets_Click()
    If chkNativeOffsets.value = vbChecked Then
        NativeShowOffsets = True
    Else
        NativeShowOffsets = False
    End If
End Sub

Private Sub chkPCODE_Click()
    If chkPCODE.value = vbChecked Then
        gPcodeDecompile = False
        Call SaveSetting("VB Decompiler", "Options", "DisablePCode", "FALSE")
    Else
        gPcodeDecompile = True
        Call SaveSetting("VB Decompiler", "Options", "DisablePCode", "TRUE")
    End If
End Sub

Private Sub chkShowColors_Click()
    If chkShowColors.value = vbChecked Then
        gShowColors = True
        Call SaveSetting("VB Decompiler", "Options", "ShowColors", "TRUE")
    Else
        gShowColors = False
        Call SaveSetting("VB Decompiler", "Options", "ShowColors", "FALSE")
    End If
End Sub

Private Sub chkShowOffsets_Click()
    If chkShowOffsets.value = vbChecked Then
        gShowOffsets = True
        Call SaveSetting("VB Decompiler", "Options", "ShowOffsets", "TRUE")
    Else
        gShowOffsets = False
        Call SaveSetting("VB Decompiler", "Options", "ShowOffsets", "FALSE")
    End If
End Sub

Private Sub chkShowPCodeStringAddress_Click()
    If chkShowPCodeStringAddress.value = vbChecked Then
        ShowPCodeStringAddress = True
    Else
        ShowPCodeStringAddress = False
    
    End If
End Sub

Private Sub chkSkipCOM_Click()
    If chkSkipCOM.value = vbChecked Then
        gSkipCom = True
    Else
        gSkipCom = False
    End If
End Sub

Private Sub cmdChangeComment_Click()
    cdColor.ShowColor
    If cdColor.Color = 0 Then Exit Sub
    lblComments.ForeColor = cdColor.Color
    modGlobals.CommentColor = cdColor.Color
End Sub

Private Sub cmdChangeKeyword_Click()
    cdColor.ShowColor
    If cdColor.Color = 0 Then Exit Sub
    lblKeyword.ForeColor = cdColor.Color
    modGlobals.InstrColor = cdColor.Color

End Sub

Private Sub cmdChangeStrings_Click()
    cdColor.ShowColor
    If cdColor.Color = 0 Then Exit Sub
    lblStrings.ForeColor = cdColor.Color
    modGlobals.StringColor = cdColor.Color
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If gSkipCom = True Then chkSkipCOM.value = vbChecked
    If gDumpData = True Then Me.chkDumpControls.value = vbChecked
    If gShowOffsets = True Then Me.chkShowOffsets.value = vbChecked
    If gShowColors = True Then Me.chkShowColors.value = vbChecked
    If gPcodeDecompile = False Then Me.chkPCODE.value = vbChecked
    
    If PCODEDisplayAddress = True Then Me.chkDisplayAddress.value = vbChecked
    If PCodeDisplayHexInfo = True Then Me.chkDisplayHex.value = vbChecked
    If ShowPCodeStringAddress = True Then Me.chkShowPCodeStringAddress.value = vbChecked
    
    If NativeShowOffsets = True Then Me.chkNativeOffsets.value = vbChecked
    If NativeShowHexInformation = True Then Me.chkNativeHex.value = vbChecked
    
    lblStrings.ForeColor = modGlobals.StringColor
    lblKeyword.ForeColor = modGlobals.InstrColor
    lblComments.ForeColor = modGlobals.CommentColor
End Sub


