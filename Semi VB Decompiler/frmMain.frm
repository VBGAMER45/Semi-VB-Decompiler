VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{FCEA04FA-85AF-4857-AF33-3842A581C8BC}#1.0#0"; "pePropertySheet.ocx"
Begin VB.Form frmMain 
   Caption         =   "Semi VB Decompiler - VisualBasicZone.com"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7890
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1                   vbgamer45"
   ScaleHeight     =   6375
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Tag             =   "                                   v b g a m e r 4 5"
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   2280
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27A2
            Key             =   "NULL"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2ABC
            Key             =   "OPEN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DD6
            Key             =   "SAVE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30F0
            Key             =   "PRINT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":340A
            Key             =   "CUT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3724
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A3E
            Key             =   "PASTE"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D58
            Key             =   "FIND"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4072
            Key             =   "HELP"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":438C
            Key             =   "LOCK"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46A6
            Key             =   "GENERATE"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C40
            Key             =   "NET"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F5A
            Key             =   "DOCUMENT"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5274
            Key             =   "CALC"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5386
            Key             =   "PCODE"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5498
            Key             =   "TICK"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":57B2
            Key             =   "NCODE"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58C4
            Key             =   "REPORT"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtFinal 
      Height          =   1695
      Index           =   0
      Left            =   10400
      TabIndex        =   22
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2990
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":5E5E
   End
   Begin VB.Frame FrameStatus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Form Generating Status"
      Height          =   3135
      Left            =   1680
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdSkipProcedure 
         Caption         =   "&Skip Procedure"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         ToolTipText     =   "Use this option if a P-Code procedure hangs or takes too long to process. Sometimes you need two clicks."
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         ToolTipText     =   "To cancel VB decompiling in case it hangs."
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtStatus 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   240
         Width           =   4335
      End
   End
   Begin RichTextLib.RichTextBox txtBuffer 
      Height          =   855
      Left            =   8640
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":5EE0
   End
   Begin VB.ListBox lstMembers 
      Height          =   2400
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin RichTextLib.RichTextBox buffCodeAp 
      Height          =   1935
      Left            =   8760
      TabIndex        =   12
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   3413
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":5F62
   End
   Begin RichTextLib.RichTextBox buffCodeAv 
      Height          =   1575
      Left            =   8040
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":5FED
   End
   Begin RichTextLib.RichTextBox txtFunctions 
      Height          =   615
      Left            =   8400
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":6078
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   6105
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   476
      Style           =   1
      SimpleText      =   "Credits: VB Decompiling community... Sarge, Napalm, Mr. Unleaded, Moogman, _aLfa_, Alex Ionescu, Warning, and others..."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   5400
      Top             =   -360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglistControl 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   51
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":60FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6694
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C80
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":801A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":906C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":93BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9710
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A62
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A106
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A45A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A7AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AB00
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AE52
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B1A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B4F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B848
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BB9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BEEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C23E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C592
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C8E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CC36
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CF88
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D2DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D62C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D97E
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E1F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EA62
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F2D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FB46
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":103B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10F7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":117EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12060
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":128D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12F78
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13512
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14046
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":145E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1473A
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14A8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15026
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":155C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15B5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":160EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":165B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvProject 
      Height          =   6075
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   10716
      _Version        =   393217
      Indentation     =   617
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "imglistControl"
      BorderStyle     =   1
      Appearance      =   0
      OLEDropMode     =   1
   End
   Begin TabDlg.SSTab sstViewFile 
      Height          =   6075
      Left            =   3480
      TabIndex        =   1
      Tag             =   "T{20/21/}"
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   10716
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Code"
      TabPicture(0)   =   "frmMain.frx":16AB4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtCode"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Properties"
      TabPicture(1)   =   "frmMain.frx":16AD0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fxgEXEInfo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Preview"
      TabPicture(2)   =   "frmMain.frx":16AEC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picPreview"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Edit Object"
      TabPicture(3)   =   "frmMain.frx":16B08
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pePropTree"
      Tab(3).Control(1)=   "cmdColor(0)"
      Tab(3).Control(2)=   "txtEditArray(0)"
      Tab(3).Control(3)=   "lblHelpText"
      Tab(3).Control(4)=   "lblObjectName"
      Tab(3).Control(5)=   "lblArrayEdit(0)"
      Tab(3).ControlCount=   6
      Begin pePropertyEditor.pePropertyTree pePropTree 
         Height          =   4215
         Left            =   -74760
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   7435
         BackColor       =   -2147483644
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LockWindowUpdate=   0   'False
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   -71040
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtEditArray 
         Height          =   285
         Index           =   0
         Left            =   -73200
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5415
         Left            =   -74880
         ScaleHeight     =   5385
         ScaleWidth      =   3945
         TabIndex        =   13
         Top             =   480
         Width           =   3975
      End
      Begin RichTextLib.RichTextBox txtCode 
         Height          =   5535
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   9763
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMain.frx":16B24
      End
      Begin MSFlexGridLib.MSFlexGrid fxgEXEInfo 
         Height          =   5535
         Left            =   -74940
         TabIndex        =   2
         Top             =   480
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   9763
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   -2147483627
         ForeColorFixed  =   12829635
         GridColorFixed  =   8421504
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   1
         AllowUserResizing=   3
      End
      Begin MSComDlg.CommonDialog cdlShow 
         Left            =   -74940
         Top             =   7590
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin RichTextLib.RichTextBox txtResult 
         Height          =   675
         Left            =   -74940
         TabIndex        =   3
         Top             =   8190
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   1191
         _Version        =   393217
         BackColor       =   12632256
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":16BA6
      End
      Begin VB.Label lblHelpText 
         Height          =   615
         Left            =   -74640
         TabIndex        =   25
         Top             =   4815
         Width           =   3375
      End
      Begin VB.Label lblObjectName 
         Caption         =   "ObjectName:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   16
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label lblArrayEdit 
         Caption         =   "Property Name"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.ListBox lstTypeInfos 
      Height          =   2400
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Members:"
      Height          =   195
      Left            =   7920
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TypeInfos:"
      Height          =   195
      Left            =   7920
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileDebugProcess 
         Caption         =   "&Debug VB Process"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileGenerate 
         Caption         =   "&Generate vbp"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFileSaveExe 
         Caption         =   "&Save Exe Changes"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileExportMemoryMap 
         Caption         =   "&Export Memory Map"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileLanguage 
         Caption         =   "&Language"
         Begin VB.Menu mnuLanguageArray 
            Caption         =   "English"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFileAntiDecompiler 
         Caption         =   "&Anti VB Decompiler Protect"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent1 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent2 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent3 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent4 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsPCodeProcedure 
         Caption         =   "&P-Code Procedure Decompile"
      End
      Begin VB.Menu mnuToolsNativeProcdureDecompile 
         Caption         =   "&Native Procedure Decompile"
      End
      Begin VB.Menu mnuToolsDecompileFromOffset 
         Caption         =   "&Decompile from Offset"
      End
      Begin VB.Menu mnuToolsAddressToFileOffset 
         Caption         =   "Address To File Offset"
      End
      Begin VB.Menu mnuToolsStartupPatcher 
         Caption         =   "&Startup Form Patcher"
      End
      Begin VB.Menu mnuToolsViewReport 
         Caption         =   "&View File Report"
      End
      Begin VB.Menu mnuToolsNetConsole 
         Caption         =   "View .Net &Console"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOtherTools 
      Caption         =   "Othe&r Tools"
      Begin VB.Menu mnuOtherToolsObfuscator 
         Caption         =   "VB Obfuscator"
      End
      Begin VB.Menu mnuOtherToolsTypeLib 
         Caption         =   "Type Library Explorer"
      End
      Begin VB.Menu mnuOtherToolsVB123 
         Caption         =   "VB 1/2/3 Binary Form To Text"
      End
      Begin VB.Menu mnuOtherToolsApiAdd 
         Caption         =   "Api Add"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpReportBug 
         Caption         =   "&Report Bug"
      End
      Begin VB.Menu mnuHelpCheckForUpdates 
         Caption         =   "&Check for Update"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpSaveImage 
         Caption         =   "Save Image"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TODO
'REGISTRY ENTERY DLL OCX EXE REMBEMBER IDA


'*********************************************
'*Semi VB Decompiler
'*Copyright VisualBasicZone.com 2004-2007
'*By vbgamer45
'*Credits:
'*Some structures from decompiler.theautomaters.com  The VB Decompiling Community
'*Sarge for the PE Skeleton
'*Mr. Unleaded for MemoryMap
'*Napalm api files.
'*Moogman for TypeViewer
'*Brad Martinez for parts of modFrx
'*Alex Ionescu for his help for COM and structures
'*And from Warning for treeview
'*For Contact or request documentation email vbgamer45@gmail.com
'*********************************************

'The following is used for the browse for folder dialog
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

'Api List
Private Type APIRecord
    Function As String
    Definition As String
End Type
Const APIRecLen As Integer = 536
Dim APIRecTotal As Integer '= 1548

'Used for syntax highlighting
Dim LinesCheck() As String
Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Private m_lAboutId As Long

Private Sub cmdCancel_Click()
    CancelDecompile = True
End Sub

Private Sub cmdColor_Click(index As Integer)
    If cmdColor(index).tag = "c" Then
     CD1.Color = txtEditArray(index).Text
     CD1.ShowColor
     txtEditArray(index).Text = CD1.Color
    End If
    If cmdColor(index).tag = "f" Then
    'FontName
    On Error Resume Next
     CD1.FontName = txtEditArray(index).Text
     CD1.ShowFont
     txtEditArray(index).Text = CD1.FontName
    End If
    If cmdColor(index).tag = "p" Then
    'Picture property
        Dim strResponse As String
        Dim bUsed As Boolean
        strResponse = MsgBox("Are you sure you want to delete this image?", vbYesNo + vbInformation, "Delete Image?")
        bUsed = False
        If strResponse = vbYes Then
            For i = 0 To UBound(PictureChange)
                If lblArrayEdit(index).tag = PictureChange(i).offset Then
                    bUsed = True
                End If
            Next
            If bUsed = False Then
                ReDim Preserve PictureChange(UBound(PictureChange) + 1)
                PictureChange(UBound(PictureChange)).offset = lblArrayEdit(index).tag
            End If
            MsgBox "Image Deleted!", vbInformation
        End If
    
    End If
    
End Sub

Private Sub cmdSkipProcedure_Click()
    bSkipProcedure = True
End Sub


'Private Sub ctlPopMenu_SystemMenuClick(ItemNumber As Long)
'    Select Case ItemNumber
'        Case m_lAboutId
'            frmAbout.Show vbModal, Me
'    End Select
'End Sub

Private Sub Form_Load()
'*****************************
'Purpose: To set all our decompiler and load any functions that need to be loaded.
'*****************************
On Error GoTo errHandle
    Me.Caption = "Semi VB Decompiler - VisualBasicZone.com  Version: " & Version
    Call PrintReadMe
    'Setup Variables
    gSkipCom = False
    gDumpData = False
    gShowOffsets = False
    gShowColors = True
    gPcodeDecompile = True
    CancelDecompile = False
    ShowPCodeStringAddress = True
    NativeShowOffsets = True
    NativeShowHexInformation = True
    gVB4App = False
    gVB5App = False
    gVB6App = False
    'Set default syntax colors
    InstrColor = vbBlue '&H400000
    FuncColor = vbBlue Xor vbRed
    CommentColor = &H8000&
    StringColor = &H80&
    'Load Languages
    Call frmMain.LoadLanguageList
    'Setup language array
    Dim i As Integer
    mnuLanguageArray(0).Caption = gLanguageList(0)
    For i = 1 To UBound(gLanguageList)
        Load mnuLanguageArray(i)
        With mnuLanguageArray(i)
            .Caption = gLanguageList(i)
        End With
    Next
    'Load Default of English
    Call LoadLanguage("english")
    
    'Load the object type's
    Call modGlobals.LoadObjectType
    
    'Get the recent file list
    Dim Recent1Title As String
    Dim Recent2Title As String
    Dim Recent3Title As String
    Dim Recent4Title As String
    
    Recent1Title = GetSetting("VB Decompiler", "Options", "Recent1FileTitle", vbNullString)
    Recent2Title = GetSetting("VB Decompiler", "Options", "Recent2FileTitle", vbNullString)
    Recent3Title = GetSetting("VB Decompiler", "Options", "Recent3FileTitle", vbNullString)
    Recent4Title = GetSetting("VB Decompiler", "Options", "Recent4FileTitle", vbNullString)
    
    If Recent1Title <> vbNullString Then
        mnuFileRecent1.Visible = True
        mnuFileSep1.Visible = True
        mnuFileRecent1.Caption = Recent1Title
    End If
    If Recent2Title <> vbNullString Then
        mnuFileRecent2.Visible = True
        mnuFileRecent2.Caption = Recent2Title
    End If
    If Recent3Title <> vbNullString Then
        mnuFileRecent3.Visible = True
        mnuFileRecent3.Caption = Recent3Title
    End If
    If Recent4Title <> vbNullString Then
        mnuFileRecent4.Visible = True
        mnuFileRecent4.Caption = Recent4Title
    End If

   ' With ctlPopMenu
   '     ' Add a new about item to the system menu:
   '     m_lAboutId = .SystemMenuAppendItem("&About...")
   '     .OfficeXpStyle = True
        
        ' Associate the image list:
   '     .ImageList = ilsIcons
        
        ' Parse through the VB designed menu and sub class the items:
   '     .SubClassMenu Me
   '     pSetIcon "OPEN", "mnuFileOpen"
   '     pSetIcon "SAVE", "mnuFileSaveExe"
   '     pSetIcon "NET", 32
   '     pSetIcon "REPORT", 29
   '     pSetIcon "PCODE", 24
   '     pSetIcon "NCODE", 25
   '     pSetIcon "CALC", 27
   '     pSetIcon "GENERATE", "mnuFileGenerate"
   '     pSetIcon "LOCK", 14
   ' End With
    
    'Get Options from Registry
    Dim strPCodeDecompile As String
    strPCodeDecompile = GetSetting("VB Decompiler", "Options", "DisablePCode", "FALSE")
    If strPCodeDecompile = "FALSE" Then
        gPcodeDecompile = False
    Else
        gPcodeDecompile = True
    End If
    Dim strShowOffsets As String
    strShowOffsets = GetSetting("VB Decompiler", "Options", "ShowOffsets", "FALSE")
    If strShowOffsets = "TRUE" Then
        modGlobals.gShowOffsets = True
    Else
        modGlobals.gShowOffsets = False
    End If
    Dim strShowColors As String
    strShowColors = GetSetting("VB Decompiler", "Options", "ShowColors", "TRUE")
    If strShowColors = "TRUE" Then
        modGlobals.gShowColors = True
    Else
        modGlobals.gShowColors = False
    End If
    Dim strGetPcode As String
    strGetPcode = GetSetting("VB Decompiler", "Options", "ShowPCODEAddress", "TRUE")
    If strGetPcode = "TRUE" Then
        PCODEDisplayAddress = True
    Else
        PCODEDisplayAddress = False
    End If
    strGetPcode = GetSetting("VB Decompiler", "Options", "ShowPCODEHex", "TRUE")
    If strGetPcode = "TRUE" Then
        PCodeDisplayHexInfo = True
    Else
        PCodeDisplayHexInfo = False
    End If
    
    'Register the typelib
    Call modTypeLB.RegisterOLB(App.Path & "\data\VB6.OLB")
    'Regsiter VB4 Typelib
    Call modTypeLB.RegisterOLB(App.Path & "\data\VB32.OLB")
    Set cTypeInfo = New clsTypeLibInfo
    
    If Not cTypeInfo.OpenTypeLib(App.Path & "\data\VB6.OLB") Then
        Debug.Print "Couldn't open typelib."
        Exit Sub
    End If

    'Setup the COM Functions
    Set tliTypeLibInfo = New TypeLibInfo
    'GUID for vb6.olb used to find the gui opcodes of the standard controls
    tliTypeLibInfo.LoadRegTypeLib "{FCFB3D2E-A0FA-1068-A738-08002B3371B5}", 6, 0, 9
    Call ProcessTypeLibrary
    tliTypeLibInfo.AppObjString = "<Global>"

    'Load Events Opcodes for standard controls
    Call modGlobals.SetupEvents
    'Load the vb Function list
    Call modNative.VBFunction_Description_Init(App.Path & "\data\vb6api.txt")

    ReDim LinesCheck(0)
    LinesCheck(0) = txtCode
    gUpdateText = False


    
Exit Sub
errHandle:
    MsgBox "Error_frmMain_Load: " & err.Description
End Sub

Private Sub Form_Resize()
'*****************************
'Purpose: When the form is resized adjust all our controls.
'*****************************
    On Error Resume Next
    tvProject.Height = Me.Height - StatusBar1.Height - 700
    sstViewFile.Height = Me.Height - StatusBar1.Height - 700
    txtCode.Height = sstViewFile.Height - 420
    Me.fxgEXEInfo.Height = sstViewFile.Height - 600
    sstViewFile.Width = Me.Width - tvProject.Width - 200
    txtCode.Width = sstViewFile.Width - 200
    fxgEXEInfo.Width = sstViewFile.Width - 200
    picPreview.Width = sstViewFile.Width - 200
    picPreview.Height = sstViewFile.Height - 600
End Sub

Private Sub fxgEXEInfo_Click()
   Call Clipboard.SetText(fxgEXEInfo.Text)
   
End Sub



Private Sub lstTypeInfos_Click()
    MsgBox lstTypeInfos.ListIndex
End Sub

Private Sub mnuAddIns_Click()

End Sub

Private Sub mnuFileAntiDecompiler_Click()
'*****************************
'Purpose: Show save dialog and encypt the current exe
'*****************************
On Error GoTo errHandle
    CD1.Filename = vbNullString
    CD1.DialogTitle = "Save File As"
    CD1.Filter = "Exe Files(*.exe)|*.exe"
    
    CD1.ShowSave
    
    If CD1.Filename = vbNullString Then Exit Sub
    
    Call modAntiDecompiler.LoadCrypter
    Call modAntiDecompiler.EncryptExe(SFilePath, CD1.Filename)
Exit Sub
errHandle:
    MsgBox "Error_frmMain_mnuFileAntiDecompiler: " & err.Number & " " & err.Description
End Sub

Private Sub mnuFileExit_Click()
'*****************************
'Purpose: To exit the decompiler and  clear any used memory
'*****************************
    Unload frmAbout
    Unload frmAddressToFileOffset
    Unload frmAdvDecompile
    cTypeInfo.CloseTypeLib
    
    Unload Me
    End
End Sub

Private Sub mnuFileExportMemoryMap_Click()
'*****************************
'Purpose: To generate a Memory Map of the current exe file.
'*****************************
On Error GoTo errHandle:
    Set gVBFile = Nothing
    Set gVBFile = New clsFile
    Call gVBFile.Setup(SFilePath)
    Dim strTitle As String
    strTitle = Me.Caption
    Me.Caption = "Generating Memory Map...Please Wait..."
    
    Set gMemoryMap = New clsMemoryMap
 
    'hascollision = gMemoryMap.AddSector(AppData.PeHeaderOffset, Len(PEHeader), "pe")
    
    
    hascollision = gMemoryMap.AddSector(VBStartHeader.PushStartAddress - OptHeader.ImageBase, 102, "vb header")
    hascollision = gMemoryMap.AddSector(gVBHeader.aProjectInfo - OptHeader.ImageBase, 572, "project info")
    hascollision = gMemoryMap.AddSector(gProjectInfo.aObjectTable - OptHeader.ImageBase, 84, "objecttable")
    hascollision = gMemoryMap.AddSector(gVBHeader.aComRegisterData - OptHeader.ImageBase, Len(modGlobals.gCOMRegData), "ComRegisterData")


    
    gMemoryMap.ExportToHTML 'exports to File.Name & ".html"
    Me.Caption = strTitle

    Dim Response As String
    Response = MsgBox("Memory Map Created! Would you like to open it now?", vbYesNo + vbInformation)
    If Response = vbYes Then
        ShellExecute Me.hwnd, vbNullString, App.Path & "\" & gVBFile.ShortFileName & ".html", vbNullString, "C:\", SW_SHOWNORMAL
    End If
    
Exit Sub
errHandle:
    MsgBox "Error_frmMain_mnuFileExportMemoryMap: " & err.Number & " " & err.Description

End Sub

Private Sub mnuFileGenerate_Click()
'*****************************
'Purpose: To generate all the vb files from the decompiled exe.
'*****************************
'On Error GoTo errHandle:
    Dim sPath As String
    Dim structFolder As BROWSEINFO
    Dim iNull As Integer
    Dim ret As Long
    Dim g As Integer
    Dim i As Integer
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
    
    'Write The Project File
    If gVB4App = True Then
      Call WriteVBPVB4(sPath & "\" & ProjectName & ".vbp")
    Else
      Call WriteVBP(sPath & "\" & ProjectName & ".vbp")
    End If


    'Write the Forms
    If VBVersion = 4 Then
        For i = 0 To UBound(strVB4Forms)
            If strVB4Forms(i) <> "" Then
                Call modOutput.WriteForms(sPath & "\" & strVB4Forms(i) & ".frm", strVB4Forms(i), i)
            End If
        Next
        
    Else
        'Write VB5/6 Forms
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 1 Then
           
                    Call modOutput.WriteForms(sPath & "\" & gObjectNameArray(i) & ".frm", gObjectNameArray(i), i)
                    Exit For
                End If
            Next g
        Next
   
        
        'Write Forms frx files
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 1 Then
                    Call modOutput.WriteFormFrx(sPath, gObjectNameArray(i), i)
                    Exit For
                End If
            Next g
        Next
        'Write the modules
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 2 Then
                    Call modOutput.WriteModules(sPath & "\" & gObjectNameArray(i) & ".bas", gObjectNameArray(i), i)
                    Exit For
                End If
            Next g
        Next
        'Write the classes
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 3 Then
                    Call modOutput.WriteClasses(sPath & "\" & gObjectNameArray(i) & ".cls", gObjectNameArray(i), i)
                    Exit For
                End If
            Next g
        Next
        'Write the user controls
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 4 Then
                    Call modOutput.WriteUserControls(sPath & "\" & gObjectNameArray(i) & ".ctl", gObjectNameArray(i), i)
                    Exit For
                End If
            Next g
        Next
        'Write property pages
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 5 Then
                    Call modOutput.WritePropertyPage(sPath & "\" & gObjectNameArray(i) & ".pag", gObjectNameArray(i), i)
                    Exit For
                End If
            Next g
        Next
        'Write User Documents
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 6 Then
                    Call modOutput.WriteUserDocument(sPath & "\" & gObjectNameArray(i) & ".dob", gObjectNameArray(i), i)
                    Exit For
                End If
            Next g
        Next
        
        'Write Designers
        For i = 0 To UBound(gObject)
            If gObject(i).ObjectType = 17926147 Then
                Call modOutput.WriteDesigner(sPath & "\" & gObjectNameArray(i) & ".dsr", gObjectNameArray(i), i)
            End If
        Next i
    End If
    
    Dim strResponse As String
    strResponse = MsgBox("Project Generated. Do you want to open the project file now?", vbYesNo + vbInformation)
    If strResponse = vbYes Then
        ShellExecute Me.hwnd, vbNullString, sPath & "\" & ProjectName & ".vbp", vbNullString, "C:\", SW_SHOWNORMAL
    End If

Exit Sub
errHandle:
    MsgBox "Error_frmMain_mnuFileGenerate: " & err.Number & " " & err.Description
End Sub

Private Sub mnuFileOpen_Click()
'*****************************
'Purpose: Show Open Dialog and then call OpenVBExe
'*****************************
' On Error GoTo errHandle:
    CD1.Filename = vbNullString
    CD1.DialogTitle = "Select VB 4/5/6 file"
    CD1.Filter = "VB Files(*.exe,*.ocx,*.dll)|*.exe;*.ocx;*.dll|All Files(*.*)|*.*;"
    CD1.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
    CD1.ShowOpen
    
    If CD1.Filename = vbNullString Then Exit Sub
    
    If FileExists(CD1.Filename) = True Then
        Call OpenVBExe(CD1.Filename, CD1.FileTitle)
    Else
        MsgBox "File Does not exist", vbExclamation
    End If
Exit Sub
errHandle:
    MsgBox "Error_frmMain_mnuFileOpen: " & err.Number & " " & err.Description
    'Show the information that you got
    Call SetupTreeView
End Sub

Sub OpenVBExe(ByVal FilePath As String, ByVal FileTitle As String, Optional bAdvDecompile As Boolean = False, Optional lImageBase As Long, Optional lFileOffset As Long)
'################################################
'Purpose: Main function that gets all VB Sturtures
'#################################################
 Dim bFormEndUsed As Boolean
 Dim i As Long 'Loop Var
 Dim k As Long 'Loop Var
 Dim addr As Long 'Loop Var
 Dim StartOffset As Long 'Holds Address of first VB Struture
 Dim F As Integer 'FileNumber holder
 
    'Erase existing data
    bFormEndUsed = False
    
    
    For i = 0 To txtFinal.UBound
        txtFinal(i).Text = vbNullString
        txtFinal(i).tag = vbNullString
    Next
    For i = txtFinal.UBound To txtFinal.UBound + 1 Step -1
        Unload txtFinal(i)
    Next i
    
    mnuFileGenerate.Enabled = False
    mnuFileExportMemoryMap.Enabled = False
    mnuFileAntiDecompiler.Enabled = False
    mnuToolsNetConsole.Visible = False
    cmdSkipProcedure.Visible = False
    SFilePath = vbNullString
    SFile = vbNullString
    ReDim gControlNameArray(0) 'Treeveiw control list
    ReDim gControlOffset(0)
    ReDim gProcedureList(0)
    ReDim gOcxList(0)
    ReDim gObjectNameArray(0)
    ReDim gExternalObjectHolder(0)
    ReDim FrxPreview(0)
    'Reset Change Types
    ReDim ByteChange(0)
    ReDim BooleanChange(0)
    ReDim IntegerChange(0)
    ReDim LongChange(0)
    ReDim SingleChange(0)
    ReDim StringChange(0)
    ReDim PictureChange(0)
    'Pcode
    ReDim EventProcList(0)
    ReDim SubNamelist(0)
    'Native
    ReDim modNative.gNativeProcArray(0)
    Close
    'clear the nodes
    tvProject.Nodes.Clear
    'Save name and path
    SFilePath = FilePath
    SFile = FileTitle
    
    'Reset the error flag
    ErrorFlag = False
    'Clear the error log
    Call modGlobals.ClearErrorLog

    gVB4App = False
    gVB5App = False
    gVB6App = False
    bISVBNET = False
    
    CancelDecompile = False
    'Get a file handle
    InFileNumber = FreeFile
    
    'Check for error
 ' On Error GoTo AnalyzeError
    
    'Access the file
    Open SFilePath For Binary As #InFileNumber


    'Is it a VB6 file?
    If CheckHeader() = True Then
        'Good file
        
        Close #InFileNumber
        If gVB4App = True Then
            OptHeader.ImageBase = mImageBaseAlign
            
            StartOffset = VBStartHeader.PushStartAddress '- mImageBaseAlign

            Call modVB4.OpenVB4EXE(SFilePath, StartOffset)
            MakeDir (App.Path & "\dump")
            MakeDir (App.Path & "\dump\" & FileTitle)
            mnuFileGenerate.Enabled = True
            mnuFileExportMemoryMap.Enabled = True
            'mnuFileAntiDecompiler.Enabled = True
            'Get FileVersion Info
            gFileInfo = modGlobals.FileInfo(SFilePath)
                
            'Hide Form Generation Status
            FrameStatus.Visible = False
            Call modOutput.DumpVBExeInfo(App.Path & "\dump\" & FileTitle & "\FileReport.txt", FileTitle)
            Call SetupTreeView
            'Add to recent files
            Call AddToRecentList(SFilePath, SFile)
        
            'Clear current data
            txtStatus.Text = vbNullString
            Exit Sub
        End If
    Else
        If VBVersion = 1 Then
            MsgBox "This Program is VB Version 1.0 and there are no decompilers for it.", vbCritical
        End If
        
       ' If VBVersion = 4 Then
       '     ' MsgBox "This Program is VB Version 4.0 and I have not made a decompiler for it yet.", vbCritical
       ' End If
        
        If VBVersion = 2 Or VBVersion = 3 Then
            MsgBox "This program is VB Version: " & VBVersion & " no decompiling option yet."
            'Decompile it
            'if neheader.
                ''Call VB3Decompile(SFilePath)
                ''Call AddToRecentList(SFilePath, SFile)
        
            'else
                'msgbox "This program is protected by DoDi"
            'End If
        End If
       'Bad file
        If bAdvDecompile = False Then
            'Get FileVersion Info
            gFileInfo = modGlobals.FileInfo(SFilePath)
            MsgBox "Not a VB 4/5/6 file.", vbOKOnly Or vbCritical Or vbApplicationModal, "Bad file!"
            gVB5App = False
            gVB4App = False
            gVB6App = False
            
        End If
        Close #InFileNumber
        If bAdvDecompile = False Then
            MakeDir (App.Path & "\dump")
            MakeDir (App.Path & "\dump\" & FileTitle)
            
            Call modOutput.DumpVBExeInfo(App.Path & "\dump\" & FileTitle & "\FileReport.txt", FileTitle)
            Call SetupTreeView
            Exit Sub
        End If
        
    End If
    

    If bAdvDecompile = True Then
        OptHeader.ImageBase = lImageBase
        StartOffset = lFileOffset
        gVB6App = True
    Else
        OptHeader.ImageBase = mImageBaseAlign
        StartOffset = VBStartHeader.PushStartAddress - OptHeader.ImageBase
       
    
    End If
    'VB Reformers....
    'OptHeader.ImageBase = frmAddressToFileOffset.txtImageBase.Text
    'StartOffset = 38388 '8508
    
    MakeDir (App.Path & "\dump")
    MakeDir (App.Path & "\dump\" & FileTitle)

   'Setup the VB File class
    Set gVBFile = New clsFile
    Call gVBFile.Setup(SFilePath)
    
    F = gVBFile.FileNumber
    Call modGlobals.GetStartUpName(F)
        'Goto begining of vb header
        Seek F, StartOffset + 1
        'Get the vb header
        Get #F, , gVBHeader

        'GetHelpFile
        
        Seek #F, StartOffset + 1 + gVBHeader.oHelpFile 'Loc(f) + gVBHeader.oHelpFile + 1
        
      
        HelpFile = GetUntilNull(F)
       
        'Get Project Name
        Seek #F, StartOffset + 1 + gVBHeader.oProjectName
        ProjectName = GetUntilNull(F)

        'Project Title
        Seek #F, StartOffset + 1 + gVBHeader.oProjectTitle
        ProjectTitle = GetUntilNull(F)

        'ExeName
        Seek #F, StartOffset + 1 + gVBHeader.oProjectExename
        ProjectExename = GetUntilNull(F)
        'Get ComRegisterData
        Seek #F, gVBHeader.aComRegisterData + 1 - OptHeader.ImageBase
        Get #F, , gCOMRegData
        Get #F, , gCOMRegInfo
        
        'Get ProjectDescription
        Seek #F, gVBHeader.aComRegisterData + 1 + gCOMRegData.oNTSProjectDescription - OptHeader.ImageBase
        ProjectDescription = GetUntilNull(F)

        
        'Get External Componetns
        '##########
        If gVBHeader.ExternalComponentCount > 0 Then
        Seek F, gVBHeader.aExternalComponentTable + 1 - OptHeader.ImageBase
            'MsgBox gVBHeader.aExternalComponentTable + 1 - OptHeader.ImageBase
            ReDim gOcxList(0)
            Dim AexternEnd As Long
            Dim bExternEnd As Long
            For i = 1 To gVBHeader.ExternalComponentCount
               bExternEnd = Loc(F)
               Dim cOcx As tComponent
               Get F, , cOcx
               'MsgBox bExternEnd + cOcx.oUuid
               AexternEnd = bExternEnd + 1 + cOcx.StructLength
               If cOcx.GUIDlength = 72 Then
                    Seek F, bExternEnd + 1 + cOcx.GUIDoffset
                    gOcxList(UBound(gOcxList)).strGuid = UCase$(GetGuidString(F, 36))
               End If
               
               Seek F, bExternEnd + 1 + cOcx.oUuid
               Dim sGuid1 As String
               'Dim sGuid2 As String
               sGuid1 = gVBFile.GetGUID(Loc(F))
               'sGuid2 = gVBFile.GetGUID(Loc(f))
               'Debug.Print sGuid1
               gOcxList(UBound(gOcxList)).strGuid = sGuid1

               Seek F, bExternEnd + 1 + cOcx.FileNameOffset
               gOcxList(UBound(gOcxList)).strocxName = GetUntilNull(F)
               'MsgBox gOcxList(UBound(gOcxList)).strGuid
               Seek F, bExternEnd + 1 + cOcx.SourceOffset
               gOcxList(UBound(gOcxList)).strLibname = GetUntilNull(F)
               Seek F, bExternEnd + 1 + cOcx.NameOffset
               gOcxList(UBound(gOcxList)).strName = GetUntilNull(F)
               ReDim Preserve gOcxList(UBound(gOcxList) + 1)
               Seek F, AexternEnd
            Next
        End If
        
        'Get Project Info Table
        Seek F, gVBHeader.aProjectInfo + 1 - OptHeader.ImageBase
        Get #F, , gProjectInfo
        'Begin Main Loop to get api list
        Dim nApi As Integer
        Dim APIFile As Integer
        Dim iAPILoop As Integer
        Dim tAPIRec As APIRecord
        Dim sTempFunctionName As String
        ReDim gApiList(0)
        APIFile = FreeFile

        Open App.Path + "\data\winapi.dat" For Random As #APIFile Len = APIRecLen
            APIRecTotal = LOF(APIFile) \ APIRecLen

            APIRecTotal = APIRecTotal + 1
        For nApi = 0 To gProjectInfo.ExternalCount - 1
  
            'Get External Table 'Number of Api Calls
            Seek F, gProjectInfo.aExternalTable + 1 + (nApi * 8) - OptHeader.ImageBase
            Get #F, , gExternalTable
            'Get External Library
            If gProjectInfo.ExternalCount > 0 And gExternalTable.Flag <> 6 Then
                Seek F, gExternalTable.aExternalLibrary + 1 - OptHeader.ImageBase
                
                Get #F, , gExternalLibrary
                If gExternalLibrary.aLibraryFunction <> 0 Then
                    Seek F, gExternalLibrary.aLibraryFunction + 1 - OptHeader.ImageBase
                    
                    sTempFunctionName = GetUntilNull(F)
                    For iAPILoop = 1 To APIRecTotal
                        Get #APIFile, iAPILoop, tAPIRec
                        If UCase$(tAPIRec.Function) = UCase$(sTempFunctionName) Then
                            gApiList(UBound(gApiList)).strFunctionName = tAPIRec.Definition
                            Exit For
                        End If
                    Next iAPILoop
                    If gApiList(UBound(gApiList)).strFunctionName = vbNullString Then
                        gApiList(UBound(gApiList)).strFunctionName = sTempFunctionName
                        Seek F, gExternalLibrary.aLibraryName + 1 - OptHeader.ImageBase
                        gApiList(UBound(gApiList)).strLibraryName = GetUntilNull(F)
                    End If
                    ReDim Preserve gApiList(UBound(gApiList) + 1)
                End If
            End If
        Next nApi 'End Api List Loop
        Close #APIFile
    
        'Get Object Table
        Seek F, gProjectInfo.aObjectTable + 1 - OptHeader.ImageBase
        Get #F, , gObjectTable
        
     
        'Resize for the number of objects...(forms,modules,classes)
        If gObjectTable.ObjectCount1 > 0 Then
        ReDim gObject(gObjectTable.ObjectCount1 - 1)
        ReDim gObjectNameArray(gObjectTable.ObjectCount1 - 1)
        ReDim gObjectProcCountArray(gObjectTable.ObjectCount1 - 1)
        ReDim gObjectInfoHolder(gObjectTable.ObjectCount1 - 1)
        End If
        'Get Object
        Seek F, gObjectTable.aObject + 1 - OptHeader.ImageBase
        Get #F, , gObject
       
        
        Dim loopC As Integer
        For loopC = 0 To UBound(gObject)
        'Get ObjectName
        Seek F, gObject(loopC).aObjectName + 1 - OptHeader.ImageBase
        gObjectNameArray(loopC) = GetUntilNull(F)
        gObjectProcCountArray(loopC) = gObject(loopC).ProcCount

        'Get Object Info
        Seek F, gObject(loopC).aObjectInfo + 1 - OptHeader.ImageBase
        Get #F, , gObjectInfo
        
        'Save the information for later on
        gObjectInfoHolder(loopC).aConstantPool = gObjectInfo.aConstantPool
        gObjectInfoHolder(loopC).aObject = gObjectInfo.aObject
        gObjectInfoHolder(loopC).aObjectTable = gObjectInfo.aObjectTable
        gObjectInfoHolder(loopC).aProcTable = gObjectInfo.aProcTable
        gObjectInfoHolder(loopC).aSmallRecord = gObjectInfo.aSmallRecord
        gObjectInfoHolder(loopC).Const1 = gObjectInfo.Const1
        gObjectInfoHolder(loopC).Flag1 = gObjectInfo.Flag1
        gObjectInfoHolder(loopC).iConstantsCount = gObjectInfo.iConstantsCount
        gObjectInfoHolder(loopC).iMaxConstants = gObjectInfo.iMaxConstants
        gObjectInfoHolder(loopC).Flag5 = gObjectInfo.Flag5
        gObjectInfoHolder(loopC).Flag6 = gObjectInfo.Flag6
        gObjectInfoHolder(loopC).Flag7 = gObjectInfo.Flag7
        gObjectInfoHolder(loopC).Null1 = gObjectInfo.Null1
        gObjectInfoHolder(loopC).Null2 = gObjectInfo.Null2
        gObjectInfoHolder(loopC).NumberOfProcs = gObjectInfo.NumberOfProcs
        gObjectInfoHolder(loopC).ObjectIndex = gObjectInfo.ObjectIndex
        gObjectInfoHolder(loopC).RunTimeLoaded = gObjectInfo.RunTimeLoaded
        
        'If gObjectInfo.aProcTable - OptHeader.ImageBase > 0 Then
            'Dim ProcCodeInfo As tCodeInfo
            'Seek f, gObjectInfo.aProcTable + 1 - OptHeader.ImageBase
            'Get f, , ProcCodeInfo
        'End If
        'If gObjectInfo.aConstantPool <> 0 Then
            'Seek f, gObjectInfo.aConstantPool + 1 - OptHeader.ImageBase
       ' End If
        
         'Get Optional Object Info
        Seek F, gObject(loopC).aObjectInfo + 57 - OptHeader.ImageBase
        
        'Decide if to get Optional Info or not
        If ((gObject(loopC).ObjectType And &H80) = &H80) Then
            
            Get #F, , gOptionalObjectInfo
            'Dim testLink() As tEventLink
            Dim LinkPCode() As MethodLinkPCode
            Dim LinkNative As MethodLinkNative
            
           ' MsgBox gOptionalObjectInfo.aEventLinkArray + 1 - OptHeader.ImageBase
            'MsgBox gOptionalObjectInfo.iEventCount
            'Resize the Arrays
            If gOptionalObjectInfo.iEventCount > 0 Then
                ReDim LinkPCode(gOptionalObjectInfo.iEventCount - 1)
                
            
                'MsgBox gOptionalObjectInfo.iEventCount
                If gOptionalObjectInfo.aEventLinkArray <> 0 And gOptionalObjectInfo.aEventLinkArray <> -1 Then
                    If gOptionalObjectInfo.aEventLinkArray + 1 - OptHeader.ImageBase > 0 Then
                        Seek F, gOptionalObjectInfo.aEventLinkArray + 1 - OptHeader.ImageBase
                        If gProjectInfo.aNativeCode = 0 Then
                        'P-Code
                            Get F, , LinkPCode
                        Else
                        'Native
                            'MsgBox Loc(F)
                            Dim lNative() As Long
                            ReDim lNative(gOptionalObjectInfo.iEventCount - 1)
                            Seek F, Loc(F) + 1
                            'MsgBox Loc(f)
                            'Get F, Loc(F) + 1, lNative

                            Dim currPos As Long
                            Dim lastOffset As Long
                            
                            For i = 0 To UBound(lNative)
                                
                                Get F, Loc(F) + 1, lNative(i)
                               ' frmNativeDecompile.lstProcedures.AddItem lNative(i) - OptHeader.ImageBase
                               ' frmNativeDecompile.lstProcedures.AddItem lNative(i) + lastOffset + currPos + 6 + 1
                                'lastOffset = lNative(i) - OptHeader.ImageBase
                                'frmNativeDecompile.lstProcedures.AddItem lNative(i) + currPos - OptHeader.ImageBase
                                'MsgBox lNative(i) + 1 - OptHeader.ImageBase
                            'MsgBox LinkNative(i).jmpOffset + Loc(F) + 5 - OptHeader.ImageBase
                            Next i
                            For i = 0 To UBound(lNative)
                            On Error Resume Next
                                'MsgBox lNative(i) - OptHeader.ImageBase
                                Get F, lNative(i) + 1 - OptHeader.ImageBase, LinkNative
                                'MsgBox LinkNative.jmpOpCode
                                'MsgBox LinkNative.jmpOffset + Loc(f) + 5
                                 currPos = Loc(F) + 1
     
                                 gNativeProcArray(UBound(gNativeProcArray)).sName = gObjectNameArray(loopC) & ".proc:" & LinkNative.jmpoffset + 5 + currPos + 5 + OptHeader.ImageBase
                                 'Debug.Print gNativeProcArray(UBound(gNativeProcArray)).sName
                                 gNativeProcArray(UBound(gNativeProcArray)).offset = LinkNative.jmpoffset + 5 + currPos + 5 + OptHeader.ImageBase
                                 ReDim Preserve gNativeProcArray(UBound(gNativeProcArray) + 1)
                                 'frmNativeDecompile.lstProcedures.AddItem LinkNative.jmpOffset + 5 + currPos + 5 + OptHeader.ImageBase
                            Next i
                        End If
                    
                    
                    'For i = 0 To UBound(LinkPCode)
                       ' MsgBox LinkPCode(i).movAddress '+ 1 - OptHeader.ImageBase
                    'Next
                    End If
                End If
            End If
        End If

        'Address PublicBytes
        'Notes aPublicBytes points to a structure of 2 integers (iStringBytes and iVarBytes) and this structure tells how many pointers will be in memory at aModulePublic.
        If gObject(loopC).aPublicBytes <> 0 Then
            Seek #F, gObject(loopC).aPublicBytes + 1 - OptHeader.ImageBase
            Dim iStringBytes As Integer, iVarBytes As Integer
            Get F, , iStringBytes
            Get F, , iVarBytes
       
            If gObject(loopC).aModulePublic <> 0 Then
                Seek #F, gObject(loopC).aModulePublic + 1 - OptHeader.ImageBase
               
            End If
        End If
        
        'Resize the control array
        'Check if its a form
        If gObject(loopC).ObjectType = 98435 Or gObject(loopC).ObjectType = 17926147 Or gObject(loopC).ObjectType = 98467 Or gObject(loopC).ObjectType = 98499 Then
   
          If gOptionalObjectInfo.ControlCount < 5000 And gOptionalObjectInfo.ControlCount <> 0 Then
            ReDim gControl(gOptionalObjectInfo.ControlCount - 1)
            'Get Control Array
            Seek F, gOptionalObjectInfo.aControlArray + 1 - OptHeader.ImageBase
            Get #F, , gControl
   
            'Resize Event Table array
            ReDim gEventTable(UBound(gControl))

            Dim ControlName As String
            
            For i = 0 To UBound(gControl)
                'Get Event Table
               Seek F, gControl(i).aEventTable + 1 - OptHeader.ImageBase
               ' ReDim gEventTable(i).aEventPointer(gControl(i).EventCount - 1)
                ReDim taEventPointer(gControl(i).EventCount - 1)
                'MsgBox gOptionalObjectInfo.iEventCount & " " & gControl(i).EventCount
                Get #F, , gEventTable(i)
                Get #F, , taEventPointer
      
                If gControl(i).aName + 1 - OptHeader.ImageBase > 0 Then
                 Seek F, gControl(i).aName + 1 - OptHeader.ImageBase
                 ControlName = GetUntilNull(F)
                 Dim strGuid As String
                 Seek F, gControl(i).aGUID + 1 - OptHeader.ImageBase
                 strGuid = modGlobals.ReturnGuid(F)

                For k = 0 To UBound(taEventPointer)
                    If taEventPointer(k) <> 0 Then
                    '  MsgBox "Good:" & ControlName & " " & taEventPointer(k) + 1 - OptHeader.ImageBase & " #" & k
                       'MsgBox "Offset: " & taEventPointer(k) + 1 - OptHeader.ImageBase
                        Dim pointerAevent As tEventPointer
                        Seek F, taEventPointer(k) + 1 - OptHeader.ImageBase
                        Get F, , pointerAevent

                        If pointerAevent.aEvent <> 0 Then
               
                          ' Debug.Print strGuid
                           'Debug.Print "Event #: " & k
                           'Debug.Print "Max Events: " & UBound(taEventPointer)
       
                            If GetEventNumber(strGuid, CInt(k)) = -1 Then
                             SubNamelist(UBound(SubNamelist)).strName = gObjectNameArray(loopC) & "." & ControlName & "_Event" & CInt(k)
                             SubNamelist(UBound(SubNamelist)).offset = pointerAevent.aEvent
                                                      
                                gProcedureList(UBound(gProcedureList)).strProcedureName = ControlName & "_Event" & CInt(k)
                                gProcedureList(UBound(gProcedureList)).strParent = gObjectNameArray(loopC)
                                ReDim Preserve gProcedureList(UBound(gProcedureList) + 1)
                            Else
                                SubNamelist(UBound(SubNamelist)).strName = gObjectNameArray(loopC) & "." & ControlName & "_" & getEventComplete(App.Path & "\data\VB6.OLB", strGuid, GetEventNumber(strGuid, CInt(k)))
                                SubNamelist(UBound(SubNamelist)).offset = pointerAevent.aEvent
                                gProcedureList(UBound(gProcedureList)).strProcedureName = ControlName & "_" & getEventComplete(App.Path & "\data\VB6.OLB", strGuid, GetEventNumber(strGuid, CInt(k)))
                                gProcedureList(UBound(gProcedureList)).strParent = gObjectNameArray(loopC)
                                ReDim Preserve gProcedureList(UBound(gProcedureList) + 1)
                            
                            End If
                          ' MsgBox gObjectNameArray(loopC) & "." & ControlName & "_" & getEventComplete(App.Path & "\data\VB6.OLB", strGuid, GetEventNumber(strGuid, CInt(k)))
                           ' Dim k234 As Byte
                           ' For k234 = 0 To 40
                           '  Debug.Print "#" & k234 & " " & getEventComplete(App.Path & "\data\VB6.OLB", strGuid, CInt(k234))
                           ' Next
                          
                            ReDim Preserve SubNamelist(UBound(SubNamelist) + 1)
                            EventProcList(UBound(EventProcList)) = pointerAevent.aEvent 'taEventPointer(k)
                            ReDim Preserve EventProcList(UBound(EventProcList) + 1)
                        End If
                    End If

                Next
                 

                 'Save the control information for the treeview
                 ReDim Preserve gControlNameArray(UBound(gControlNameArray) + 1)
                 gControlNameArray(UBound(gControlNameArray)).strControlName = ControlName
                 gControlNameArray(UBound(gControlNameArray)).strParentForm = gObjectNameArray(loopC)
                 gControlNameArray(UBound(gControlNameArray)).strGuid = strGuid
                End If
            Next
            End If
        End If
        'Get Proc Names
        
        If gObject(loopC).ProcCount <> 0 Then

            If gObject(loopC).aProcNamesArray <> 0 Then
            Dim AddressProcNamesArray() As Long
            ReDim AddressProcNamesArray(gObject(loopC).ProcCount - 1)
            
            
            Seek F, gObject(loopC).aProcNamesArray + 1 - OptHeader.ImageBase
            Get F, , AddressProcNamesArray
   
                For addr = 0 To UBound(AddressProcNamesArray)
                   
                    If AddressProcNamesArray(addr) = 0 Then
                    
                    Else
                        If (AddressProcNamesArray(addr) - OptHeader.ImageBase) < 0 Then
                           ' MsgBox AddressProcNamesArray(addr)
                        Else
                            Seek F, AddressProcNamesArray(addr) + 1 - OptHeader.ImageBase
                            
                        
                            gProcedureList(UBound(gProcedureList)).strProcedureName = GetUntilNull(F)
                            
                            gProcedureList(UBound(gProcedureList)).strParent = gObjectNameArray(loopC)
                            SubNamelist(UBound(SubNamelist)).strName = gProcedureList(UBound(gProcedureList)).strProcedureName
                            SubNamelist(UBound(SubNamelist)).offset = AddressProcNamesArray(addr)
                            ReDim Preserve SubNamelist(UBound(SubNamelist) + 1)
                            
                            ReDim Preserve gProcedureList(UBound(gProcedureList) + 1)
                        End If
                    End If
                Next
         
            End If

        
        End If
        Next loopC

        'Main Loop to Get all Form's Properties
        FrameStatus.Visible = True
        txtStatus.Text = vbNullString
        Call ProccessControls(F)
        Call modGlobals.DoFinalFormBuffer
    Close F

    
    'Set the compile type either pcode or ncode
    If gProjectInfo.aNativeCode <> 0 Then
        AppData.CompileType = "Native"
        'Begin Native Decompile
        Call modNative.Decode(SFilePath)
    Else
        AppData.CompileType = "PCode"
        'Begin Pcode Decompile
        txtStatus.Text = txtStatus.Text & "Begin PCode Decompile" & vbCrLf
        Call modPCode.init
        'Decompile the file
        If gPcodeDecompile = True Then
            Call modPCode.Decode(SFilePath)
        End If
        txtStatus.Text = txtStatus.Text & "End PCode Decompile" & vbCrLf
        
    End If
    
    

    mnuFileGenerate.Enabled = True
    mnuFileExportMemoryMap.Enabled = True
    'mnuFileAntiDecompiler.Enabled = True
    'Get FileVersion Info
    gFileInfo = modGlobals.FileInfo(SFilePath)
    
    'Hide Form Generation Status
    FrameStatus.Visible = False
    
    Call SetupTreeView
    Call modOutput.DumpVBExeInfo(App.Path & "\dump\" & FileTitle & "\FileReport.txt", FileTitle)
    
    'Write Error log
    Call modGlobals.WriteErrorLog(App.Path & "\dump\" & FileTitle & "\ErrorLog.txt")
    
    'Add to recent files
    Call AddToRecentList(SFilePath, SFile)

    'Clear current data
    txtStatus.Text = vbNullString
    
  Exit Sub
    
AnalyzeError:
    FrameStatus.Visible = False
    MsgBox "Analyze error " & err.Description, vbCritical Or vbOKOnly, "Source file error"
    
    Call SetupTreeView
    MakeDir (App.Path & "\dump")
    MakeDir (App.Path & "\dump\" & FileTitle)
    Close
    Call modOutput.DumpVBExeInfo(App.Path & "\dump\" & FileTitle & "\FileReport.txt", FileTitle)
    Close
End Sub
Sub AddToRecentList(Filename As String, FileTitle As String)
'*****************************
'Purpose: Add a Filename to the recently access list via the registry
'*****************************
On Error GoTo errHandle
    Dim Recent1File As String
    Dim Recent1Title As String
    Dim Recent2File As String
    Dim Recent2Title As String
    Dim Recent3File As String
    Dim Recent3Title As String
    
    mnuFileSep1.Visible = True
    mnuFileRecent1.Visible = True
    
    Recent1Title = GetSetting("VB Decompiler", "Options", "Recent1FileTitle", vbNullString)
    Recent2Title = GetSetting("VB Decompiler", "Options", "Recent2FileTitle", vbNullString)
    Recent3Title = GetSetting("VB Decompiler", "Options", "Recent3FileTitle", vbNullString)
    Recent1File = GetSetting("VB Decompiler", "Options", "Recent1File", vbNullString)
    Recent2File = GetSetting("VB Decompiler", "Options", "Recent2File", vbNullString)
    Recent3File = GetSetting("VB Decompiler", "Options", "Recent3File", vbNullString)
    
   
    
    If Recent1Title <> vbNullString Then
        mnuFileRecent2.Visible = True
    End If
    If Recent2Title <> vbNullString Then
        mnuFileRecent3.Visible = True
    End If
    If Recent3Title <> vbNullString Then
        mnuFileRecent4.Visible = True
    End If


    Call SaveSetting("VB Decompiler", "Options", "Recent4File", Recent3File)
    Call SaveSetting("VB Decompiler", "Options", "Recent4FileTitle", Recent3Title)
    Call SaveSetting("VB Decompiler", "Options", "Recent3File", Recent2File)
    Call SaveSetting("VB Decompiler", "Options", "Recent3FileTitle", Recent2Title)
    Call SaveSetting("VB Decompiler", "Options", "Recent2File", Recent1File)
    Call SaveSetting("VB Decompiler", "Options", "Recent2FileTitle", Recent1Title)


    Call SaveSetting("VB Decompiler", "Options", "Recent1File", Filename)
    Call SaveSetting("VB Decompiler", "Options", "Recent1FileTitle", FileTitle)
    
    
    
    mnuFileRecent4.Caption = mnuFileRecent3.Caption
    mnuFileRecent3.Caption = mnuFileRecent2.Caption
    mnuFileRecent2.Caption = mnuFileRecent1.Caption
    mnuFileRecent1.Caption = FileTitle

    
    'ctlPopMenu.Caption(19) = ctlPopMenu.Caption(18)
    'ctlPopMenu.Caption(18) = ctlPopMenu.Caption(17)
    'ctlPopMenu.Caption(17) = ctlPopMenu.Caption(16)
    'ctlPopMenu.Caption(16) = FileTitle

    
Exit Sub
errHandle:
    MsgBox "Error_frmMain_AddToRecentList: " & err.Description

End Sub
Sub MakeDir(ByVal Path As String)
'*****************************
'Purpose: To make a dir without erroring
'*****************************
On Error Resume Next
    MkDir (Path)

End Sub

Private Sub mnuFileRecent1_Click()
'*****************************
'Purpose: To load a recent file if it exists
'*****************************
    Dim RecentTitle As String
    Dim RecentFile As String
    RecentTitle = GetSetting("VB Decompiler", "Options", "Recent1FileTitle", vbNullString)
    RecentFile = GetSetting("VB Decompiler", "Options", "Recent1File", vbNullString)
    If FileExists(RecentFile) = True Then
        Call OpenVBExe(RecentFile, RecentTitle)
    Else
        MsgBox "File no longer exists!", vbExclamation
    End If
End Sub

Private Sub mnuFileRecent2_Click()
'*****************************
'Purpose: To load a recent file if it exists
'*****************************
Dim RecentTitle As String
Dim RecentFile As String
    RecentTitle = GetSetting("VB Decompiler", "Options", "Recent2FileTitle", vbNullString)
    RecentFile = GetSetting("VB Decompiler", "Options", "Recent2File", vbNullString)
    If FileExists(RecentFile) = True Then
        Call OpenVBExe(RecentFile, RecentTitle)
    Else
        MsgBox "File no longer exists!", vbExclamation
    End If
End Sub

Private Sub mnuFileRecent3_Click()
'*****************************
'Purpose: To load a recent file if it exists
'*****************************
Dim RecentTitle As String
Dim RecentFile As String
    RecentTitle = GetSetting("VB Decompiler", "Options", "Recent3FileTitle", vbNullString)
    RecentFile = GetSetting("VB Decompiler", "Options", "Recent3File", vbNullString)
    If FileExists(RecentFile) = True Then
        Call OpenVBExe(RecentFile, RecentTitle)
    Else
        MsgBox "File no longer exists!", vbExclamation
    End If
End Sub

Private Sub mnuFileRecent4_Click()
'*****************************
'Purpose: To load a recent file if it exists
'*****************************
    Dim RecentTitle As String
    Dim RecentFile As String
    RecentTitle = GetSetting("VB Decompiler", "Options", "Recent4FileTitle", vbNullString)
    RecentFile = GetSetting("VB Decompiler", "Options", "Recent4File", vbNullString)
    If FileExists(RecentFile) = True Then
        Call OpenVBExe(RecentFile, RecentTitle)
    Else
        MsgBox "File no longer exists!", vbExclamation
    End If
End Sub

Private Sub mnuFileSaveExe_Click()
'#####################################
'Purpose: Save Changes to the Form's Gui
'And generates a Patch Report
'#####################################
    CD1.DialogTitle = "Save As"
    CD1.Filename = vbNullString
    CD1.Filter = "VB Files(*.exe,*.ocx,*.dll)|*.exe;*.ocx;*.dll"
    CD1.ShowSave
    
    If CD1.Filename = vbNullString Then Exit Sub
    On Error Resume Next
    'Copy the exe to the temp directory
    Close
    FileCopy SFilePath, App.Path & "\dump\" & SFile & "\" & SFile
    
    'Make the changes
    fFile = FreeFile
    Dim i As Integer
    Dim NewByte As Byte
    Dim bArray() As Byte
    Open App.Path & "\dump\" & SFile & "\" & SFile For Binary Access Write As fFile
        If UBound(StringChange) > 0 Then
            For i = 1 To UBound(StringChange)
                Seek fFile, StringChange(i).offset '+ 1
                
                ReDim bArray(Len(StringChange(i).sString) - 1)
                For g = 0 To Len(StringChange(i).sString) - 1
                    bArray(g) = Asc(Mid$(StringChange(i).sString, 1 + g, 1))
                Next g
                Put fFile, , bArray
                'Put fFile, , StringChange(I).sString
            Next
        End If
        If UBound(ByteChange) > 0 Then
            For i = 1 To UBound(ByteChange)
                Seek fFile, ByteChange(i).offset + 1
                Put fFile, , ByteChange(i).bByte
            Next
        End If
        
        If UBound(BooleanChange) > 0 Then
           
            For i = 1 To UBound(BooleanChange)
                Seek fFile, BooleanChange(i).offset + 1
                 'MsgBox "yo" & BooleanChange(i).Offset
                If BooleanChange(i).bBool = True Then
                    NewByte = 255
                    Put fFile, , NewByte
                Else
                    NewByte = 0
                    Put fFile, , NewByte
                End If
                'Put fFile, , ByteChange(i).bByte
            Next i
        End If
        If UBound(IntegerChange) > 0 Then
            For i = 1 To UBound(IntegerChange)
                Seek fFile, IntegerChange(i).offset + 1
                Put fFile, , IntegerChange(i).iInt
            Next
        End If
        If UBound(LongChange) > 0 Then
            For i = 1 To UBound(LongChange)
                Seek fFile, LongChange(i).offset + 1
                Put fFile, , LongChange(i).lLong
            Next
        End If
        If UBound(SingleChange) > 0 Then
            For i = 1 To UBound(SingleChange)
                Seek fFile, SingleChange(i).offset + 1
                Put fFile, , SingleChange(i).sSingle
            Next
        End If
        If UBound(PictureChange) > 0 Then
            Dim FileNum As Integer
            Dim lPictureSize As Long
            'Dim iByte As Byte
            'iByte = 255
            FileNum = FreeFile
            
            Dim picHeader As typePictureHeader
            Open App.Path & "\dump\" & SFile & "\" & SFile For Binary Access Read As #FileNum
            For i = 1 To UBound(PictureChange)
                Get #FileNum, PictureChange(i).offset - 3, lPictureSize
                Get #FileNum, PictureChange(i).offset, picHeader
                'MsgBox lPictureSize - Len(picHeader)
                'MsgBox PictureChange(i).offset
                'ReDim bArray(lPictureSize - Len(picHeader) - 1)
                ReDim bArray(lPictureSize - 1)
              
               '' 'Put fFile, PictureChange(i).offset - 5, 0
                'iByte = 255
               '''Put fFile, PictureChange(i).offset - 4, iByte
                'iByte = 2
               '''Put fFile, PictureChange(i).offset - 3, iByte

                Put fFile, PictureChange(i).offset, bArray
                Put fFile, PictureChange(i).offset - 5, 0
                Put fFile, PictureChange(i).offset - 4, 255
                Put fFile, PictureChange(i).offset - 3, 2

            Next
            
            Close #FileNum
        
        End If
        
        
    Close fFile
    
    'Save the file
    FileCopy App.Path & "\dump\" & SFile & "\" & SFile, CD1.Filename
    'Kill the temp file
    Kill App.Path & "\dump\" & SFile & "\" & SFile
    
    'Write Patch Report
    
    fFile = FreeFile
   
    Open App.Path & "\dump\" & SFile & "\PatchReport.txt" For Output As fFile
        Print #fFile, "File Patch Report from Semi VB Decompiler by VisualBasicZone.com"
        Print #fFile, "------------------------------------------------------"
        Print #fFile, "Filename=" & SFile
        Print #fFile, ""
        Print #fFile, "Byte Changes"
        For i = 0 To UBound(ByteChange)
                If i <> 0 Then
                    Print #fFile, "Offset:" & ByteChange(i).offset & " Changed to: " & ByteChange(i).bByte
                End If
        Next i
        Print #fFile, ""
        Print #fFile, "Boolean Changes"
        For i = 0 To UBound(BooleanChange)
                If i <> 0 Then
                    Print #fFile, "Offset:" & BooleanChange(i).offset & " Changed to: " & BooleanChange(i).bBool
                End If
        Next i
        Print #fFile, ""
        Print #fFile, "Integer Changes"
        For i = 0 To UBound(IntegerChange)
                If i <> 0 Then
                    Print #fFile, "Offset:" & IntegerChange(i).offset & " Changed to: " & IntegerChange(i).iInt
                End If
        Next i
        Print #fFile, ""
        Print #fFile, "Long Changes"
        For i = 0 To UBound(LongChange)
                If i <> 0 Then
                    Print #fFile, "Offset:" & LongChange(i).offset & " Changed to: " & LongChange(i).lLong
                End If
        Next i
        Print #fFile, ""
        Print #fFile, "Single Changes"
        For i = 0 To UBound(SingleChange)
                If i <> 0 Then
                    Print #fFile, "Offset:" & SingleChange(i).offset & " Changed to: " & SingleChange(i).sSingle
                End If
        Next i
        Print #fFile, ""
        Print #fFile, "String Changes"
        For i = 0 To UBound(StringChange)
                If i <> 0 Then
                    Print #fFile, "Offset:" & StringChange(i).offset & " Changed to: " & StringChange(i).sString
                End If
        Next i
        Print #fFile, ""
        Print #fFile, "Picture Changes"
        For i = 0 To UBound(PictureChange)
                If i <> 0 Then
                    Print #fFile, "Offset:" & PictureChange(i).offset & " Changed to: nothing"
                End If
        Next i
        
    Close fFile
    
    MsgBox "Done changes saved check the patch report for details.", vbInformation
End Sub

Private Sub mnuHelpAbout_Click()
'*****************************
'Purpose: Show my Cool about screen.
'*****************************
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpCheckForUpdates_Click()
    frmCheckUpdate.Show vbModal, Me
    
End Sub

Private Sub mnuHelpReportBug_Click()
    ShellExecute Me.hwnd, vbNullString, "mailto:support@visualbasiczone.com", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub mnuLanguageArray_Click(index As Integer)
    Call LoadLanguage(mnuLanguageArray(index).Caption)
End Sub

Private Sub mnuOptions_Click()
'*****************************
'Purpose: Show the options form
'*****************************
    frmOptions.Show vbModal, Me
    
End Sub

Private Sub mnuOtherToolsApiAdd_Click()
On Error GoTo errHandle
    Call Shell(App.Path & "\ApiLoader.exe", vbNormalFocus)

Exit Sub
errHandle:
    MsgBox "Error_ApiAdd" & err.Number & " " & err.Description
End Sub

Private Sub mnuOtherToolsObfuscator_Click()
On Error GoTo errHandle
    Call Shell(App.Path & "\VBObfuscator.exe", vbNormalFocus)
Exit Sub
errHandle:
    MsgBox "Error_" & err.Number & " " & err.Description
End Sub

Private Sub mnuOtherToolsTypeLib_Click()
On Error GoTo errHandle
    Call Shell(App.Path & "\TypeLibraryExplorer.exe", vbNormalFocus)
Exit Sub
errHandle:
    MsgBox "Error_" & err.Number & " " & err.Description
End Sub

Private Sub mnuOtherToolsVB123_Click()
On Error GoTo errHandle
    Call Shell(App.Path & "\prjVBBinaryToText.exe", vbNormalFocus)

Exit Sub
errHandle:
    MsgBox "Error_" & err.Number & " " & err.Description
End Sub

Private Sub mnuPopUpSaveImage_Click()
'*****************************
'Purpose: Save an image from the preview picture box
'*****************************
On Error GoTo errHandle
    CD1.Filename = vbNullString
    CD1.DialogTitle = "Save Image"
    CD1.Filter = "Image Files(*.*)|*.*;"
    CD1.ShowSave
    If CD1.Filename = vbNullString Then Exit Sub
    
    Call SavePicture(picPreview.Image, CD1.Filename)
Exit Sub
errHandle:
    MsgBox "Error_frmMain_mnuPopupSaveImage: " & err.Number & " " & err.Description
End Sub

Private Sub mnuToolsAddressToFileOffset_Click()
    frmAddressToFileOffset.Show vbModal, Me
End Sub

Private Sub mnuToolsDecompileFromOffset_Click()
    frmAdvDecompile.Show vbModal, Me
End Sub

Private Sub mnuToolsNativeProcdureDecompile_Click()

    If SFilePath = vbNullString Then
        MsgBox "No File Loaded", vbInformation
        Exit Sub
    End If
    If gVB4App = True Or gVB5App = True Or gVB6App = True Then
        If modGlobals.gProjectInfo.aNativeCode <> 0 Then
            frmNativeDecompile.Show vbModal, Me
        Else
            MsgBox "This is a P-Code compiled exe! Not a Native one!", vbExclamation
        End If
    Else
        MsgBox "Not a VB Native Program", vbInformation
    End If
    
    
End Sub

Private Sub mnuToolsNetConsole_Click()
On Error GoTo errHandle
    If SFilePath = vbNullString Then
        MsgBox "No File Loaded", vbInformation
        Exit Sub
    End If
    ShellExecute Me.hwnd, vbNullString, App.Path & "\dump\" & SFile & "\netconsole.txt", vbNullString, "C:\", SW_SHOWNORMAL
    Exit Sub
errHandle:
Exit Sub
End Sub

Private Sub mnuToolsPCodeProcedure_Click()
'*****************************
'Purpose: Show the Procedure Decompile View
'*****************************
    If gPcodeDecompile = False Then
        MsgBox "P-Code Decompiling is disabled. Please goto Options and enable it then reopen the file.", vbInformation
        Exit Sub
    End If

    If SFilePath = vbNullString Then
        MsgBox "No File Loaded", vbInformation
        Exit Sub
    End If
    If gVB4App = True Or gVB5App = True Or gVB6App = True Then
        If modGlobals.gProjectInfo.aNativeCode = 0 Then
            frmPcode.Show vbModal, Me
        Else
            MsgBox "This is a Native compiled exe! Not a P-Code one!", vbExclamation
        End If
    Else
        MsgBox "Not a VB P-Code Program", vbInformation
    End If
    
End Sub

Private Sub mnuToolsStartupPatcher_Click()
    If SFilePath = vbNullString Then
        MsgBox "No File Loaded", vbInformation
        Exit Sub
    End If
    If gVB5App = True Or gVB6App = True Then
        frmStartUp.Show vbModal, Me
    Else
        MsgBox "Startup Patching is only for VB 5/6 Files"
    
    End If
End Sub

Private Sub mnuToolsViewReport_Click()
On Error GoTo errHandle
    If SFilePath = vbNullString Then
        MsgBox "No File Loaded", vbInformation
        Exit Sub
    End If
    ShellExecute Me.hwnd, vbNullString, App.Path & "\dump\" & SFile & "\FileReport.txt", vbNullString, "C:\", SW_SHOWNORMAL
    Exit Sub
errHandle:
Exit Sub
End Sub

Private Sub pePropTree_PropertyChanged(oPropItem As pePropertyEditor.CPropertyItem)
On Error Resume Next
    Dim i As Long
    For i = 0 To lblArrayEdit.UBound
        If lblArrayEdit(i).Caption = oPropItem.Caption Then
            txtEditArray(i).Text = oPropItem.value
            oPropItem.value = txtEditArray(i).Text
            Exit For
        End If
    Next
End Sub

Private Sub pePropTree_PropertySelected(oPropItem As pePropertyEditor.CPropertyItem)
On Error GoTo errHandle
    lblHelpText.Caption = oPropItem.HelpString
Exit Sub
errHandle:
Exit Sub
End Sub

Private Sub picPreview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
End Sub
Private Sub tvProject_NodeClick(ByVal Node As MSComctlLib.Node)
'*****************************
'Purpose: To show the contents of each struture and textbox data
'*****************************
On Error Resume Next

    Dim i As Long
    Dim tblPath() As String
    Dim strBuffer As String
    txtCode.SelStart = 0
    txtCode.SelColor = vbBlack
    
    If CurrentItem <> tvProject.SelectedItem.Key Then
        tblPath = Split(tvProject.SelectedItem.Key, "/")
        CurrentItem = tvProject.SelectedItem.Key

        Select Case tblPath(1)
            Case "VERSIONINFO"
                sstViewFile.TabVisible(1) = True
                sstViewFile.TabVisible(0) = False
                sstViewFile.TabVisible(2) = False
                sstViewFile.TabVisible(3) = False
                fxgEXEInfo.Visible = True
                fxgEXEInfo.ColAlignment(1) = 0
                fxgEXEInfo.Clear
                fxgEXEInfo.Rows = 2
                fxgEXEInfo.ColWidth(0) = 2500
                fxgEXEInfo.TextArray(0) = "Name"
                fxgEXEInfo.TextArray(1) = "Value"
                        fxgEXEInfo.ColWidth(0) = 2000
                        fxgEXEInfo.ColWidth(1) = 2500
                        fxgEXEInfo.TextArray(2) = "CompanyName"
                        fxgEXEInfo.TextArray(3) = gFileInfo.CompanyName
                        fxgEXEInfo.AddItem "FileDescription"
                        fxgEXEInfo.TextArray(5) = gFileInfo.FileDescription
                        fxgEXEInfo.AddItem "FileVersion"
                        fxgEXEInfo.TextArray(7) = gFileInfo.FileVersion
                        fxgEXEInfo.AddItem "InternalName"
                        fxgEXEInfo.TextArray(9) = gFileInfo.InternalName
                        fxgEXEInfo.AddItem "LanguageID"
                        fxgEXEInfo.TextArray(11) = gFileInfo.LanguageID
                        fxgEXEInfo.AddItem "LegalCopyright"
                        fxgEXEInfo.TextArray(13) = gFileInfo.LegalCopyright
                        fxgEXEInfo.AddItem "OrigionalFileName"
                        fxgEXEInfo.TextArray(15) = gFileInfo.OrigionalFileName
                        fxgEXEInfo.AddItem "ProductName"
                        fxgEXEInfo.TextArray(17) = gFileInfo.ProductName
                        fxgEXEInfo.AddItem "ProductVersion"
                        fxgEXEInfo.TextArray(19) = gFileInfo.ProductVersion
                        fxgEXEInfo.AddItem "Comments"
                        fxgEXEInfo.TextArray(21) = gFileInfo.Comments
                        fxgEXEInfo.AddItem "LegalTrademark"
                        fxgEXEInfo.TextArray(23) = gFileInfo.LegalTradeMark
            Case "STRUCT"
                sstViewFile.TabVisible(1) = True
                sstViewFile.TabVisible(0) = False
                sstViewFile.TabVisible(2) = False
                sstViewFile.TabVisible(3) = False
                fxgEXEInfo.Visible = True
                fxgEXEInfo.ColAlignment(1) = 0
                fxgEXEInfo.Clear
                fxgEXEInfo.Rows = 2
                fxgEXEInfo.ColWidth(0) = 2500
                fxgEXEInfo.TextArray(0) = "Name"
                fxgEXEInfo.TextArray(1) = "Value"
                Select Case tblPath(2)
                    Case "", "VBHEADER"
                        If gVB4App = False Then
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "Signature"
                        fxgEXEInfo.TextArray(3) = gVBHeader.signature
                        fxgEXEInfo.AddItem "Address SubMain"
                        fxgEXEInfo.TextArray(5) = gVBHeader.aSubMain
                        fxgEXEInfo.AddItem "Address ExternalComponentTable"
                        fxgEXEInfo.TextArray(7) = gVBHeader.aExternalComponentTable
                        fxgEXEInfo.AddItem "Address GUITable"
                        fxgEXEInfo.TextArray(9) = gVBHeader.aGuiTable
                        fxgEXEInfo.AddItem "Address ComRegisterData"
                        fxgEXEInfo.TextArray(11) = gVBHeader.aComRegisterData
                        fxgEXEInfo.AddItem "Address ProjectInfo"
                        fxgEXEInfo.TextArray(13) = gVBHeader.aProjectInfo
                        fxgEXEInfo.AddItem "BackupLanguageDLL"
                        fxgEXEInfo.TextArray(15) = gVBHeader.BackupLanguageDLL
                        fxgEXEInfo.AddItem "BackupLanguageID"
                        fxgEXEInfo.TextArray(17) = gVBHeader.BackupLanguageID
                        fxgEXEInfo.AddItem "ExternalComponentCount"
                        fxgEXEInfo.TextArray(19) = gVBHeader.ExternalComponentCount
                        fxgEXEInfo.AddItem "Flag MDLIntObjs"
                        fxgEXEInfo.TextArray(21) = gVBHeader.fMDLIntObjs
                        fxgEXEInfo.AddItem "Flag MDLIntObjs2"
                        fxgEXEInfo.TextArray(23) = gVBHeader.fMDLIntObjs2
                        fxgEXEInfo.AddItem "FormCount"
                        fxgEXEInfo.TextArray(25) = gVBHeader.FormCount
                        fxgEXEInfo.AddItem "LanguageDLL"
                        fxgEXEInfo.TextArray(27) = gVBHeader.LanguageDLL
                        fxgEXEInfo.AddItem "LanguageID"
                        fxgEXEInfo.TextArray(29) = gVBHeader.LanguageID
                        fxgEXEInfo.AddItem "Offset HelpFile"
                        fxgEXEInfo.TextArray(31) = gVBHeader.oHelpFile
                        fxgEXEInfo.AddItem "Offset ProjectExename"
                        fxgEXEInfo.TextArray(33) = gVBHeader.oProjectExename
                        fxgEXEInfo.AddItem "Offset ProjectName"
                        fxgEXEInfo.TextArray(35) = gVBHeader.oProjectName
                        fxgEXEInfo.AddItem "Offset ProjectTitle"
                        fxgEXEInfo.TextArray(37) = gVBHeader.oProjectTitle
                        fxgEXEInfo.AddItem "RuntimeDLLVersion"
                        fxgEXEInfo.TextArray(39) = gVBHeader.RuntimeDLLVersion
                        fxgEXEInfo.AddItem "RuntimeBuild"
                        fxgEXEInfo.TextArray(41) = gVBHeader.RuntimeBuild
                        fxgEXEInfo.AddItem "ThreadCount"
                        fxgEXEInfo.TextArray(43) = gVBHeader.ThreadCount
                        fxgEXEInfo.AddItem "ThreadFlags"
                        fxgEXEInfo.TextArray(45) = gVBHeader.ThreadFlags
                        fxgEXEInfo.AddItem "ThunkCount"
                        fxgEXEInfo.TextArray(47) = gVBHeader.ThunkCount
                        Else
                            Call ShowVB4Header
                        End If
                    Case "VB4HEADER"
                        If gVB4App = True Then
                            Call ShowVB4Header
                        End If
                    Case "VBPROJECTINFO"
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "Address EndOfCode"
                        fxgEXEInfo.TextArray(3) = gProjectInfo.aEndOfCode
                        fxgEXEInfo.AddItem "Address ExternalTable"
                        fxgEXEInfo.TextArray(5) = gProjectInfo.aExternalTable
                        fxgEXEInfo.AddItem "Address NativeCode"
                        fxgEXEInfo.TextArray(7) = gProjectInfo.aNativeCode
                        fxgEXEInfo.AddItem "Address ObjectTable"
                        fxgEXEInfo.TextArray(9) = gProjectInfo.aObjectTable
                        fxgEXEInfo.AddItem "Address StartOfCode"
                        fxgEXEInfo.TextArray(11) = gProjectInfo.aStartOfCode
                        fxgEXEInfo.AddItem "Address VBAExceptionhandler"
                        fxgEXEInfo.TextArray(13) = gProjectInfo.aVBAExceptionhandler
                        fxgEXEInfo.AddItem "ExternalCount"
                        fxgEXEInfo.TextArray(15) = gProjectInfo.ExternalCount
                        fxgEXEInfo.AddItem "Flag1"
                        fxgEXEInfo.TextArray(17) = gProjectInfo.Flag1
                        fxgEXEInfo.AddItem "Flag2"
                        fxgEXEInfo.TextArray(19) = gProjectInfo.Flag2
                        fxgEXEInfo.AddItem "Flag3"
                        fxgEXEInfo.TextArray(21) = gProjectInfo.Flag3
                        fxgEXEInfo.AddItem "Null1"
                        fxgEXEInfo.TextArray(23) = gProjectInfo.Null1
                        fxgEXEInfo.AddItem "NullSpacer"
                        fxgEXEInfo.TextArray(25) = gProjectInfo.NullSpacer
                        fxgEXEInfo.AddItem "oProjectLocation"
                        fxgEXEInfo.TextArray(27) = gProjectInfo.oProjectLocation
                        fxgEXEInfo.AddItem "OriginalPathName"
                        fxgEXEInfo.TextArray(29) = gProjectInfo.OriginalPathName
                        fxgEXEInfo.AddItem "Signature"
                        fxgEXEInfo.TextArray(31) = gProjectInfo.signature
                        fxgEXEInfo.AddItem "ThreadSpace"
                        fxgEXEInfo.TextArray(33) = gProjectInfo.ThreadSpace
                    Case "VBCOMREGDATA"
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "TlbVerMajor"
                        fxgEXEInfo.TextArray(3) = gCOMRegData.iTlbVerMajor
                        fxgEXEInfo.AddItem "iTlbVerMinor"
                        fxgEXEInfo.TextArray(5) = gCOMRegData.iTlbVerMinor
                        fxgEXEInfo.AddItem "Padding1"
                        fxgEXEInfo.TextArray(7) = gCOMRegData.iPadding1
                        fxgEXEInfo.AddItem "Padding2"
                        fxgEXEInfo.TextArray(9) = gCOMRegData.iPadding2
                        fxgEXEInfo.AddItem "Padding3"
                        fxgEXEInfo.TextArray(11) = gCOMRegData.lPadding3
                        fxgEXEInfo.AddItem "lTlbLcid"
                        fxgEXEInfo.TextArray(13) = gCOMRegData.lTlbLcid
                        fxgEXEInfo.AddItem "Offset to NTSHelpDirectory"
                        fxgEXEInfo.TextArray(15) = gCOMRegData.oNTSHelpDirectory
                        fxgEXEInfo.AddItem "Offset to NTSProjectDescription"
                        fxgEXEInfo.TextArray(17) = gCOMRegData.oNTSProjectDescription
                        fxgEXEInfo.AddItem "Offset to NTSProjectName"
                        fxgEXEInfo.TextArray(19) = gCOMRegData.oNTSProjectName
                        fxgEXEInfo.AddItem "Offset to RegInfo"
                        fxgEXEInfo.TextArray(21) = gCOMRegData.oRegInfo
                        fxgEXEInfo.AddItem "uuidProjectClsId"
                        For i = 0 To UBound(gCOMRegData.uuidProjectClsId)
                            fxgEXEInfo.TextArray(23) = fxgEXEInfo.TextArray(23) & gCOMRegData.uuidProjectClsId(i)
                        Next
                    Case "VBCOMREGINFO"
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "fClassType"
                        fxgEXEInfo.TextArray(3) = gCOMRegInfo.fClassType
                        fxgEXEInfo.AddItem "fIsControl"
                        fxgEXEInfo.TextArray(5) = gCOMRegInfo.fIsControl
                        fxgEXEInfo.AddItem "fIsDesigner"
                        fxgEXEInfo.TextArray(7) = gCOMRegInfo.fIsDesigner
                        fxgEXEInfo.AddItem "fIsInterface"
                        fxgEXEInfo.TextArray(9) = gCOMRegInfo.fIsInterface
                        fxgEXEInfo.AddItem "fObjectType"
                        fxgEXEInfo.TextArray(11) = gCOMRegInfo.fObjectType
                        fxgEXEInfo.AddItem "iDefaultIcon"
                        fxgEXEInfo.TextArray(13) = gCOMRegInfo.iDefaultIcon
                        fxgEXEInfo.AddItem "iToolboxBitmap32"
                        fxgEXEInfo.TextArray(15) = gCOMRegInfo.iToolboxBitmap32
                        fxgEXEInfo.AddItem "lInstancing"
                        fxgEXEInfo.TextArray(17) = gCOMRegInfo.lInstancing
                        fxgEXEInfo.AddItem "lMiscStatus"
                        fxgEXEInfo.TextArray(19) = gCOMRegInfo.lMiscStatus
                        fxgEXEInfo.AddItem "lObjectID"
                        fxgEXEInfo.TextArray(21) = gCOMRegInfo.lObjectID
                        fxgEXEInfo.AddItem "Offset to ControlClsID"
                        fxgEXEInfo.TextArray(23) = gCOMRegInfo.oControlClsID
                        fxgEXEInfo.AddItem "Offset to DesignerData"
                        fxgEXEInfo.TextArray(25) = gCOMRegInfo.oDesignerData
                        fxgEXEInfo.AddItem "Offset to NextObject"
                        fxgEXEInfo.TextArray(27) = gCOMRegInfo.oNextObject
                        fxgEXEInfo.AddItem "Offset to ObjectClsID"
                        fxgEXEInfo.TextArray(29) = gCOMRegInfo.oObjectClsID
                        fxgEXEInfo.AddItem "Offset to ObjectDescription"
                        fxgEXEInfo.TextArray(31) = gCOMRegInfo.oObjectDescription
                        fxgEXEInfo.AddItem "Offset to ObjectName"
                        fxgEXEInfo.TextArray(33) = gCOMRegInfo.oObjectName
                        fxgEXEInfo.AddItem "uuidObjectClsID"
                        For i = 0 To UBound(gCOMRegInfo.uuidObjectClsID)
                            fxgEXEInfo.TextArray(35) = fxgEXEInfo.TextArray(35) & gCOMRegInfo.uuidObjectClsID(i)
                        Next
                    Case "VBOBJECTABLE"
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "Address of ExecProj"
                        fxgEXEInfo.TextArray(3) = gObjectTable.aExecProj
                        fxgEXEInfo.AddItem "Address of ProjectInfo2"
                        fxgEXEInfo.TextArray(5) = gObjectTable.aProjectInfo2
                        fxgEXEInfo.AddItem "Address of ProjectObject Size"
                        fxgEXEInfo.TextArray(7) = gObjectTable.lpProjectObject
                        fxgEXEInfo.AddItem "Address of First Object"
                        fxgEXEInfo.TextArray(9) = gObjectTable.aObject
                        fxgEXEInfo.AddItem "Address of ProjectName"
                        fxgEXEInfo.TextArray(11) = gObjectTable.aProjectName
                        fxgEXEInfo.AddItem "Const1"
                        fxgEXEInfo.TextArray(13) = gObjectTable.Const1
                        fxgEXEInfo.AddItem "Flag CompileType"
                        fxgEXEInfo.TextArray(15) = gObjectTable.fCompileType
                        fxgEXEInfo.AddItem "Const3"
                        fxgEXEInfo.TextArray(17) = gObjectTable.Const3
                        fxgEXEInfo.AddItem "Flag1"
                        fxgEXEInfo.TextArray(19) = gObjectTable.Flag1
                        fxgEXEInfo.AddItem "Flag2"
                        fxgEXEInfo.TextArray(21) = gObjectTable.Flag2
                        fxgEXEInfo.AddItem "Flag3"
                        fxgEXEInfo.TextArray(23) = gObjectTable.Flag3
                        fxgEXEInfo.AddItem "Flag4"
                        fxgEXEInfo.TextArray(25) = gObjectTable.Flag4
                        fxgEXEInfo.AddItem "LangID1"
                        fxgEXEInfo.TextArray(27) = gObjectTable.LangID1
                        fxgEXEInfo.AddItem "LangID2"
                        fxgEXEInfo.TextArray(29) = gObjectTable.LangID2
                        fxgEXEInfo.AddItem "Null1"
                        fxgEXEInfo.TextArray(31) = gObjectTable.lNull1
                        fxgEXEInfo.AddItem "Null2"
                        fxgEXEInfo.TextArray(33) = gObjectTable.Null2
                        fxgEXEInfo.AddItem "Null3"
                        fxgEXEInfo.TextArray(35) = gObjectTable.Null3
                        fxgEXEInfo.AddItem "Null4"
                        fxgEXEInfo.TextArray(37) = gObjectTable.Null4
                        fxgEXEInfo.AddItem "Null5"
                        fxgEXEInfo.TextArray(39) = gObjectTable.Null5
                        fxgEXEInfo.AddItem "Null6"
                        fxgEXEInfo.TextArray(41) = gObjectTable.Null6
                        fxgEXEInfo.AddItem "ObjectCount1"
                        fxgEXEInfo.TextArray(43) = gObjectTable.ObjectCount1
                        fxgEXEInfo.AddItem "CompiledObjects"
                        fxgEXEInfo.TextArray(45) = gObjectTable.iCompiledObjects
                        fxgEXEInfo.AddItem "ObjectsInUse"
                        fxgEXEInfo.TextArray(47) = gObjectTable.iObjectsInUse
                Case "VBOBJECTS"
                        If tblPath(3) <> vbNullString And UBound(tblPath) = 4 Then

                            Dim objSel As Long
                            objSel = Val(tblPath(3))
                        
                            fxgEXEInfo.ColWidth(0) = 2500
                            fxgEXEInfo.TextArray(2) = "Address of ModulePublic"
                            fxgEXEInfo.TextArray(3) = gObject(objSel).aModulePublic
                            fxgEXEInfo.AddItem "Address of ModuleStatic"
                            fxgEXEInfo.TextArray(5) = gObject(objSel).aModuleStatic
                            fxgEXEInfo.AddItem "Address of ObjectInfo"
                            fxgEXEInfo.TextArray(7) = gObject(objSel).aObjectInfo
                            fxgEXEInfo.AddItem "Address of ObjectName"
                            fxgEXEInfo.TextArray(9) = gObject(objSel).aObjectName
                            fxgEXEInfo.AddItem "Address Proc Name Array"
                            fxgEXEInfo.TextArray(11) = gObject(objSel).aProcNamesArray
                            fxgEXEInfo.AddItem "Const1"
                            fxgEXEInfo.TextArray(13) = gObject(objSel).Const1
                            fxgEXEInfo.AddItem "Address of PublicBytes"
                            fxgEXEInfo.TextArray(15) = gObject(objSel).aPublicBytes
                            fxgEXEInfo.AddItem "Address of StaticBytes"
                            fxgEXEInfo.TextArray(17) = gObject(objSel).aStaticBytes
                            fxgEXEInfo.AddItem "Offset of StaticVars"
                            fxgEXEInfo.TextArray(19) = gObject(objSel).oStaticVars
                            fxgEXEInfo.AddItem "Null3"
                            fxgEXEInfo.TextArray(21) = gObject(objSel).Null3
                            fxgEXEInfo.AddItem "ObjectType"
                            fxgEXEInfo.TextArray(23) = gObject(objSel).ObjectType
                            fxgEXEInfo.AddItem "ProcCount"
                            fxgEXEInfo.TextArray(25) = gObject(objSel).ProcCount


                        End If
                        If UBound(tblPath) = 5 Then
                            Dim objInfosel As Long
                            objInfosel = Val(tblPath(4))
                            fxgEXEInfo.ColWidth(0) = 2500
                            fxgEXEInfo.TextArray(2) = "Address of ConstantPool"
                            fxgEXEInfo.TextArray(3) = gObjectInfoHolder(objInfosel).aConstantPool
                            fxgEXEInfo.AddItem "Address of Object"
                            fxgEXEInfo.TextArray(5) = gObjectInfoHolder(objInfosel).aObject
                            fxgEXEInfo.AddItem "Address of ObjectTable"
                            fxgEXEInfo.TextArray(7) = gObjectInfoHolder(objInfosel).aObjectTable
                            fxgEXEInfo.AddItem "Address of ProcTable"
                            fxgEXEInfo.TextArray(9) = gObjectInfoHolder(objInfosel).aProcTable
                            fxgEXEInfo.AddItem "Address of SmallRecord"
                            fxgEXEInfo.TextArray(11) = gObjectInfoHolder(objInfosel).aSmallRecord
                            fxgEXEInfo.AddItem "Const1"
                            fxgEXEInfo.TextArray(13) = gObjectInfoHolder(objInfosel).Const1
                            fxgEXEInfo.AddItem "Flag1"
                            fxgEXEInfo.TextArray(15) = gObjectInfoHolder(objInfosel).Flag1
                            fxgEXEInfo.AddItem "iConstantsCount"
                            fxgEXEInfo.TextArray(17) = gObjectInfoHolder(objInfosel).iConstantsCount
                            fxgEXEInfo.AddItem "iMaxConstants"
                            fxgEXEInfo.TextArray(19) = gObjectInfoHolder(objInfosel).iMaxConstants
                            fxgEXEInfo.AddItem "Flag5"
                            fxgEXEInfo.TextArray(21) = gObjectInfoHolder(objInfosel).Flag5
                            fxgEXEInfo.AddItem "Flag6"
                            fxgEXEInfo.TextArray(23) = gObjectInfoHolder(objInfosel).Flag6
                            fxgEXEInfo.AddItem "Flag7"
                            fxgEXEInfo.TextArray(25) = gObjectInfoHolder(objInfosel).Flag7
                            fxgEXEInfo.AddItem "Null1"
                            fxgEXEInfo.TextArray(27) = gObjectInfoHolder(objInfosel).Null1
                            fxgEXEInfo.AddItem "Null2"
                            fxgEXEInfo.TextArray(29) = gObjectInfoHolder(objInfosel).Null2
                            fxgEXEInfo.AddItem "NumberOfProcs"
                            fxgEXEInfo.TextArray(31) = gObjectInfoHolder(objInfosel).NumberOfProcs
                            fxgEXEInfo.AddItem "ObjectIndex"
                            fxgEXEInfo.TextArray(33) = gObjectInfoHolder(objInfosel).ObjectIndex
                            fxgEXEInfo.AddItem "RunTimeLoaded"
                            fxgEXEInfo.TextArray(35) = gObjectInfoHolder(objInfosel).RunTimeLoaded
                        End If
                End Select
            Case "NETSTRUCT"
                sstViewFile.TabVisible(1) = True
                sstViewFile.TabVisible(0) = False
                sstViewFile.TabVisible(2) = False
            
                fxgEXEInfo.Visible = True
                fxgEXEInfo.ColAlignment(1) = 0
                fxgEXEInfo.Clear
                fxgEXEInfo.Rows = 2
                fxgEXEInfo.ColWidth(0) = 2000
                fxgEXEInfo.TextArray(0) = "Name"
                fxgEXEInfo.TextArray(1) = "Value"
                
                Select Case tblPath(2)
                        Case "", "NETHEADER"
                            Call modVBNET.ShowCLRHeader(fxgEXEInfo)
                        Case "METADATA"
                            Call modVBNET.ShowMetaDataHeader(fxgEXEInfo)
                        Case "STREAMS"
                            For i = 0 To UBound(modVBNET.gVBNetStreamHeaders)
                                If tblPath(3) = gVBNetStreamHeaders(i).rcName Or tblPath(3) = "" Then
                                    fxgEXEInfo.ColWidth(0) = 1500
                                    fxgEXEInfo.TextArray(2) = "Name"
                                    fxgEXEInfo.TextArray(3) = gVBNetStreamHeaders(i).rcName
                                    fxgEXEInfo.AddItem "Offset"
                                    fxgEXEInfo.TextArray(5) = gVBNetStreamHeaders(i).iOffset
                                    fxgEXEInfo.AddItem "Size"
                                    fxgEXEInfo.TextArray(7) = gVBNetStreamHeaders(i).iSize
                                    Exit For
                                End If
                            Next
                        Case "STRINGHEAP"
                            fxgEXEInfo.ColAlignment(1) = 0
                            fxgEXEInfo.Clear
                            fxgEXEInfo.Rows = 2
                            fxgEXEInfo.ColWidth(0) = 600
                            fxgEXEInfo.ColWidth(1) = 4000
                            fxgEXEInfo.TextArray(0) = "Offset"
                            fxgEXEInfo.TextArray(1) = "String"
                            Call modVBNET.ShowStrings(fxgEXEInfo)
                        Case "USHEAP"
                            fxgEXEInfo.ColAlignment(1) = 0
                            fxgEXEInfo.Clear
                            fxgEXEInfo.Rows = 2
                            fxgEXEInfo.ColWidth(0) = 600
                            fxgEXEInfo.ColWidth(1) = 4000
                            fxgEXEInfo.TextArray(0) = "Offset"
                            fxgEXEInfo.TextArray(1) = "String"
                            Call modVBNET.ShowUserStrings(fxgEXEInfo)
                        Case "GUIDHEAP"
                            fxgEXEInfo.ColAlignment(1) = 0
                            fxgEXEInfo.Clear
                            fxgEXEInfo.Rows = 2
                            fxgEXEInfo.ColWidth(0) = 600
                            fxgEXEInfo.ColWidth(1) = 4000
                            fxgEXEInfo.TextArray(0) = "Offset"
                            fxgEXEInfo.TextArray(1) = "String"
                            Call modVBNET.ShowGUIDHeap(fxgEXEInfo)
                        Case "BLOBHEAP"
                            fxgEXEInfo.ColAlignment(1) = 0
                            fxgEXEInfo.Clear
                            fxgEXEInfo.Rows = 2
                            fxgEXEInfo.ColWidth(0) = 600
                            fxgEXEInfo.ColWidth(1) = 4000
                            fxgEXEInfo.TextArray(0) = "Offset"
                            fxgEXEInfo.TextArray(1) = "Size"
                            Call modVBNET.ShowBlobHeap(fxgEXEInfo)
                                      
                End Select
            Case "EXEDATA"  '#####################################################'
                sstViewFile.TabVisible(1) = True
                sstViewFile.TabVisible(0) = False
                sstViewFile.TabVisible(2) = False
                fxgEXEInfo.Visible = True
                
                fxgEXEInfo.ColAlignment(1) = 0
                fxgEXEInfo.Clear
                fxgEXEInfo.Rows = 2
                fxgEXEInfo.ColWidth(0) = 2000
                fxgEXEInfo.TextArray(0) = "Name"
                fxgEXEInfo.TextArray(1) = "Value"
                
                Select Case tblPath(2)
                    Case "", "EXEHEADER"
                        fxgEXEInfo.ColWidth(0) = 1500
                        fxgEXEInfo.TextArray(2) = "Signature"
                        fxgEXEInfo.TextArray(3) = DosHeader.Magic
                        fxgEXEInfo.AddItem "Extra Bytes"
                        fxgEXEInfo.TextArray(5) = DosHeader.NumBytesLastPage
                        fxgEXEInfo.AddItem "Pages"
                        fxgEXEInfo.TextArray(7) = DosHeader.NumPages
                        fxgEXEInfo.AddItem "Reloc Items"
                        fxgEXEInfo.TextArray(9) = DosHeader.NumRelocates
                        fxgEXEInfo.AddItem "Header Size"
                        fxgEXEInfo.TextArray(11) = DosHeader.NumHeaderBlks
                        fxgEXEInfo.AddItem "Min Alloc"
                        fxgEXEInfo.TextArray(13) = DosHeader.ReservedW8
                        fxgEXEInfo.AddItem "Max Alloc"
                        fxgEXEInfo.TextArray(15) = DosHeader.ReservedW9
                        fxgEXEInfo.AddItem "Initial SS"
                        fxgEXEInfo.TextArray(17) = DosHeader.SSPointer
                        fxgEXEInfo.AddItem "Initial SP"
                        fxgEXEInfo.TextArray(19) = DosHeader.SPPointer
                        fxgEXEInfo.AddItem "Check Sum"
                        fxgEXEInfo.TextArray(21) = DosHeader.Checksum
                        fxgEXEInfo.AddItem "Initial IP"
                        fxgEXEInfo.TextArray(23) = DosHeader.IPPointer
                        fxgEXEInfo.AddItem "Initial CS"
                        fxgEXEInfo.TextArray(25) = DosHeader.CurrentSeg
                        fxgEXEInfo.AddItem "Reloc Table"
                        fxgEXEInfo.TextArray(27) = DosHeader.RelocTablePointer
                        fxgEXEInfo.AddItem "Overlay"
                        fxgEXEInfo.TextArray(29) = DosHeader.Overlay
                    Case "COFFHEADER"
                        fxgEXEInfo.ColWidth(0) = 2000
                        fxgEXEInfo.TextArray(2) = "Signature"
                        fxgEXEInfo.TextArray(3) = PEHeader.Magic
                        fxgEXEInfo.AddItem "Machine"
                        fxgEXEInfo.TextArray(5) = PEHeader.Machine
                        fxgEXEInfo.AddItem "Number Of Sections"
                        fxgEXEInfo.TextArray(7) = PEHeader.NumSections
                        fxgEXEInfo.AddItem "Time Date Stamp"
                        fxgEXEInfo.TextArray(9) = PEHeader.TimeDate
                        fxgEXEInfo.AddItem "Pointer To Symbol Table"
                        fxgEXEInfo.TextArray(11) = PEHeader.SymbolTablePointer
                        fxgEXEInfo.AddItem "Number Of Symbols"
                        fxgEXEInfo.TextArray(13) = PEHeader.NumSymbols
                        fxgEXEInfo.AddItem "Optional Header Size"
                        fxgEXEInfo.TextArray(15) = PEHeader.OptionalHdrSize
                        fxgEXEInfo.AddItem "Characteristics"
                        fxgEXEInfo.TextArray(17) = PEHeader.Properties
                    Case "NEHEADER"
                        fxgEXEInfo.ColWidth(0) = 2000
                        fxgEXEInfo.TextArray(2) = "Signature"
                        fxgEXEInfo.TextArray(3) = NEHeader.signature
                        fxgEXEInfo.AddItem "VersionLinker"
                        fxgEXEInfo.TextArray(5) = NEHeader.VersionLinker
                        fxgEXEInfo.AddItem "RevisionLinker"
                        fxgEXEInfo.TextArray(7) = NEHeader.RevisionLinker
                        fxgEXEInfo.AddItem "EntryTableOffset"
                        fxgEXEInfo.TextArray(9) = NEHeader.EntryTableOffset
                        fxgEXEInfo.AddItem "SizeOfEntryTable"
                        fxgEXEInfo.TextArray(11) = NEHeader.SizeOfEntryTable
                        fxgEXEInfo.AddItem "CRC"
                        fxgEXEInfo.TextArray(13) = NEHeader.CRC
                        fxgEXEInfo.AddItem "Flags"
                        fxgEXEInfo.TextArray(15) = NEHeader.flags
                        fxgEXEInfo.AddItem "SegmentNumberAutomaticDataSegment"
                        fxgEXEInfo.TextArray(17) = NEHeader.SegmentNumberAutomaticDataSegment
                        fxgEXEInfo.AddItem "InitialSizeHeap"
                        fxgEXEInfo.TextArray(19) = NEHeader.InitialSizeHeap
                        fxgEXEInfo.AddItem "InitialSizeStack"
                        fxgEXEInfo.TextArray(21) = NEHeader.InitialSizeStack
                        fxgEXEInfo.AddItem "SegmentNumberOffsetCS"
                        fxgEXEInfo.TextArray(23) = NEHeader.SegmentNumberOffsetCS
                        fxgEXEInfo.AddItem "SegmentNumberOffsetSS"
                        fxgEXEInfo.TextArray(25) = NEHeader.SegmentNumberOffsetSS
                        fxgEXEInfo.AddItem "NumberEntriesSegmentTable"
                        fxgEXEInfo.TextArray(27) = NEHeader.NumberEntriesSegmentTable
                        fxgEXEInfo.AddItem "NumberEntriesModuleReferenceTable"
                        fxgEXEInfo.TextArray(29) = NEHeader.NumberEntriesModuleReferenceTable
                        fxgEXEInfo.AddItem "SizeOfNonResidentNameTable"
                        fxgEXEInfo.TextArray(31) = NEHeader.SizeOfNonResidentNameTable
                        fxgEXEInfo.AddItem "SegmentTableOffset"
                        fxgEXEInfo.TextArray(33) = NEHeader.SegmentTableOffset
                        fxgEXEInfo.AddItem "ResourceTableFileOffset"
                        fxgEXEInfo.TextArray(35) = NEHeader.ResourceTableFileOffset
                        fxgEXEInfo.AddItem "ResidentNameTableOffset"
                        fxgEXEInfo.TextArray(37) = NEHeader.ResidentNameTableOffset
                        fxgEXEInfo.AddItem "ModuleReferenceTableOffset"
                        fxgEXEInfo.TextArray(39) = NEHeader.ModuleReferenceTableOffset
                        fxgEXEInfo.AddItem "ImportedNamesTableOffset"
                        fxgEXEInfo.TextArray(41) = NEHeader.ImportedNamesTableOffset
                        fxgEXEInfo.AddItem "NonResidentNameTableOffset"
                        fxgEXEInfo.TextArray(43) = NEHeader.NonResidentNameTableOffset
                        fxgEXEInfo.AddItem "NumberMovableEntriesInEntryTable"
                        fxgEXEInfo.TextArray(45) = NEHeader.NumberMovableEntriesInEntryTable
                        fxgEXEInfo.AddItem "LogicalSectorAlignmentShiftCount"
                        fxgEXEInfo.TextArray(47) = NEHeader.LogicalSectorAlignmentShiftCount
                        fxgEXEInfo.AddItem "NumberResourceEntries"
                        fxgEXEInfo.TextArray(49) = NEHeader.NumberResourceEntries
                        fxgEXEInfo.AddItem "ExecutableType"
                        fxgEXEInfo.TextArray(51) = NEHeader.ExecutableType
                    Case "OPTIONALHEADER"
                        
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "Magic"
                        fxgEXEInfo.TextArray(3) = modPeSkeleton.OptHeader.Magic
                        fxgEXEInfo.AddItem "Linker Major Version"
                        fxgEXEInfo.TextArray(5) = modPeSkeleton.OptHeader.MajLinkerVer
                        fxgEXEInfo.AddItem "Linker Minor Version"
                        fxgEXEInfo.TextArray(7) = modPeSkeleton.OptHeader.MinLinkerVer
                        fxgEXEInfo.AddItem "Size Of Code Section"
                        fxgEXEInfo.TextArray(9) = modPeSkeleton.OptHeader.CodeSize
                        fxgEXEInfo.AddItem "Initialized DataSize"
                        fxgEXEInfo.TextArray(11) = modPeSkeleton.OptHeader.InitDataSize
                        fxgEXEInfo.AddItem "Uninitialized DataSize"
                        fxgEXEInfo.TextArray(13) = modPeSkeleton.OptHeader.UninitDataSize
                        fxgEXEInfo.AddItem "Entry Point RVA"
                        fxgEXEInfo.TextArray(15) = modPeSkeleton.OptHeader.EntryPoint
                        fxgEXEInfo.AddItem "Base Of Code"
                        fxgEXEInfo.TextArray(17) = modPeSkeleton.OptHeader.CodeBase
                        fxgEXEInfo.AddItem "Base Of Data"
                        fxgEXEInfo.TextArray(19) = modPeSkeleton.OptHeader.DataBase
                        fxgEXEInfo.AddItem "Image Base"
                        fxgEXEInfo.TextArray(21) = modPeSkeleton.OptHeader.ImageBase
                        fxgEXEInfo.AddItem "Section Alignement"
                        fxgEXEInfo.TextArray(23) = modPeSkeleton.OptHeader.SectionAlignment
                        fxgEXEInfo.AddItem "File Alignement"
                        fxgEXEInfo.TextArray(25) = modPeSkeleton.OptHeader.FileAlignment
                        fxgEXEInfo.AddItem "OS Major Version"
                        fxgEXEInfo.TextArray(27) = modPeSkeleton.OptHeader.MajOSVer
                        fxgEXEInfo.AddItem "OS Minor Version"
                        fxgEXEInfo.TextArray(29) = modPeSkeleton.OptHeader.MinOSVer
                        fxgEXEInfo.AddItem "User Major Version" 'bad
                        fxgEXEInfo.TextArray(31) = modPeSkeleton.OptHeader.MajImageVer
                        fxgEXEInfo.AddItem "User Minor Version" 'bad
                        fxgEXEInfo.TextArray(33) = modPeSkeleton.OptHeader.MinImageVer
                        fxgEXEInfo.AddItem "Sub Sys Major Version"
                        fxgEXEInfo.TextArray(35) = modPeSkeleton.OptHeader.MajSSysVer
                        fxgEXEInfo.AddItem "Sub Sys Minor Version"
                        fxgEXEInfo.TextArray(37) = modPeSkeleton.OptHeader.MinSSysVer
                        fxgEXEInfo.AddItem "Reserved" 'bad
                        fxgEXEInfo.TextArray(39) = modPeSkeleton.OptHeader.SSizeRes
                        fxgEXEInfo.AddItem "Image Size"
                        fxgEXEInfo.TextArray(41) = modPeSkeleton.OptHeader.SizeImage
                        fxgEXEInfo.AddItem "Header Size"
                        fxgEXEInfo.TextArray(43) = modPeSkeleton.OptHeader.SizeHeader
                        fxgEXEInfo.AddItem "File Checksum"
                        fxgEXEInfo.TextArray(45) = modPeSkeleton.OptHeader.Checksum
                        fxgEXEInfo.AddItem "Sub System"
                        fxgEXEInfo.TextArray(47) = modPeSkeleton.OptHeader.SSystem
                        fxgEXEInfo.AddItem "DLL Flags" 'bad
                        fxgEXEInfo.TextArray(49) = modPeSkeleton.OptHeader.LFlags
                        fxgEXEInfo.AddItem "Stack Reserved Size"
                        fxgEXEInfo.TextArray(51) = modPeSkeleton.OptHeader.SSizeRes
                        fxgEXEInfo.AddItem "Stack Commit Size"
                        fxgEXEInfo.TextArray(53) = modPeSkeleton.OptHeader.SSizeCom
                        fxgEXEInfo.AddItem "Heap Reserved Size"
                        fxgEXEInfo.TextArray(55) = modPeSkeleton.OptHeader.HSizeRes
                        fxgEXEInfo.AddItem "Heap Commit Size"
                        fxgEXEInfo.TextArray(57) = modPeSkeleton.OptHeader.HSizeCom
                        fxgEXEInfo.AddItem "Loader Flags"
                        fxgEXEInfo.TextArray(59) = modPeSkeleton.OptHeader.LFlags
                        'Data Directory
                        Dim dd As Long, dta As Long
                        dta = 61
                        For dd = 0 To 15
                            fxgEXEInfo.AddItem modPeSkeleton.OptHeader.DataDirectory(dd).name & " Address:"
                            fxgEXEInfo.TextArray(dta) = modPeSkeleton.OptHeader.DataDirectory(dd).Address
                            dta = dta + 2
                            fxgEXEInfo.AddItem modPeSkeleton.OptHeader.DataDirectory(dd).name & " Size:"
                            fxgEXEInfo.TextArray(dta) = modPeSkeleton.OptHeader.DataDirectory(dd).Size
                            dta = dta + 2
                        Next
                    Case "SECTIONHEADER"
                        If tblPath(3) <> vbNullString Then
                            Dim SelSection As Long
                            SelSection = Val(tblPath(3))
                            
                            fxgEXEInfo.ColWidth(0) = 2000
                            fxgEXEInfo.TextArray(2) = "Section Name"
                            fxgEXEInfo.TextArray(3) = modPeSkeleton.SecHeader(SelSection).SecName
                           
                            fxgEXEInfo.AddItem "Virtual Size"
                            fxgEXEInfo.TextArray(5) = modPeSkeleton.SecHeader(SelSection).Properties
                            fxgEXEInfo.AddItem "RVA Offset"
                            fxgEXEInfo.TextArray(7) = modPeSkeleton.SecHeader(SelSection).Address
                            fxgEXEInfo.AddItem "Size Of Raw Data"
                            fxgEXEInfo.TextArray(9) = modPeSkeleton.SecHeader(SelSection).SizeRawData
                            fxgEXEInfo.AddItem "Pointer To Raw Data"
                            fxgEXEInfo.TextArray(11) = modPeSkeleton.SecHeader(SelSection).RawDataPointer
                            fxgEXEInfo.AddItem "Pointer To Relocs"
                            fxgEXEInfo.TextArray(13) = modPeSkeleton.SecHeader(SelSection).RelocationPointer
                            fxgEXEInfo.AddItem "Pointer To Line Numbers"
                            fxgEXEInfo.TextArray(15) = modPeSkeleton.SecHeader(SelSection).LineNumPointer
                            fxgEXEInfo.AddItem "Number Of Relocs"
                            fxgEXEInfo.TextArray(17) = modPeSkeleton.SecHeader(SelSection).NumRelocations
                            fxgEXEInfo.AddItem "Number Of Line Numbers"
                            fxgEXEInfo.TextArray(19) = modPeSkeleton.SecHeader(SelSection).NumLineNumbers
                            fxgEXEInfo.AddItem "Section Flags"
                            fxgEXEInfo.TextArray(21) = modPeSkeleton.SecHeader(SelSection).Misc
                        Else
              

                           For i = 1 To PEHeader.NumSections
                               fxgEXEInfo.AddItem " " & Left$(SecHeader(i).SecName, lstrlen(SecHeader(i).SecName)) ' ExtString(SecHeader(i).SecName)
                               fxgEXEInfo.TextArray(3 + i * 2) = Right$(String$(8, "0") & Hex$(SecHeader(i).Address), 8) 'AddChar(Hex(SecHeader(i).Address), 8)
                            Next i
                        End If
                End Select
            Case "PROJECT"  '#####################################################'
                sstViewFile.TabVisible(0) = True
                sstViewFile.TabVisible(1) = False
                sstViewFile.TabVisible(2) = False
                sstViewFile.TabVisible(3) = False
                fxgEXEInfo.Visible = False
                If gVB4App = False Then
                    Call modOutput.ShowVBPFile
                Else
                    Call modOutput.ShowVBPFileVB4
                End If
                    gUpdateText = True
                    txtCode_Change
                    gUpdateText = False

            Case "CODE"     '#####################################################'
                sstViewFile.TabVisible(0) = True
                sstViewFile.TabVisible(1) = False
                sstViewFile.TabVisible(2) = False
                sstViewFile.TabVisible(3) = False
                fxgEXEInfo.Visible = False
                Select Case tblPath(2)
                    Case "", "API"
                    sstViewFile.TabVisible(0) = True
                    sstViewFile.TabVisible(1) = False
                    sstViewFile.TabVisible(2) = False
                    sstViewFile.TabVisible(3) = False
                    fxgEXEInfo.Visible = False
                    Call modGlobals.WriteApiList
                    gUpdateText = True
                    txtCode_Change
                    gUpdateText = False
                    Case "", "PCODE"

                        sstViewFile.TabVisible(0) = True
                        sstViewFile.TabVisible(1) = False
                        sstViewFile.TabVisible(2) = False
                        sstViewFile.TabVisible(3) = False
                        fxgEXEInfo.Visible = False
                        txtCode.LoadFile (App.Path & "\dump\" & SFile & "\PcodeOut.txt")
                        frmMain.tag = "P"
                        txtCode.Font.Bold = True
                        gUpdateText = True
                        txtCode_Change
                        gUpdateText = False
                        
                        frmMain.tag = ""
                    Case "", "PCODETOVB"

                        sstViewFile.TabVisible(0) = True
                        sstViewFile.TabVisible(1) = False
                        sstViewFile.TabVisible(2) = False
                        sstViewFile.TabVisible(3) = False
                        fxgEXEInfo.Visible = False
                        txtCode.LoadFile (App.Path & "\dump\" & SFile & "\PcodeToVB.txt")
                        
                        '//gUpdateText = True
                        'txtCode_Change
                        'gUpdateText = False
                    Case "", "PCODESTRINGS"
                        sstViewFile.TabVisible(0) = True
                        sstViewFile.TabVisible(1) = False
                        sstViewFile.TabVisible(2) = False
                        sstViewFile.TabVisible(3) = False
                        fxgEXEInfo.Visible = False
                        txtCode.Text = modPCode.GetPCodeStringList
                        
                    Case "", "ASM"
                        If gVBHeader.aSubMain <> 0 Then
                            txtCode.Text = "Please refer to Native Procedure Decompile."
                        Else
                            MsgBox "Use Native Procedure Decompile under Tools Menu.", vbInformation
                        End If
                    
                End Select
                
            Case "FORMS"
                If tblPath(2) <> vbNullString Then
                    sstViewFile.TabVisible(0) = True
                    sstViewFile.TabVisible(3) = True
                    sstViewFile.TabVisible(1) = False
                    sstViewFile.TabVisible(2) = False
                    fxgEXEInfo.Visible = False
                    
                    For i = 0 To txtFinal.UBound
                        If UCase$(txtFinal(i).tag) = tblPath(2) Then
                            If gVB4App = False Then
                                strBuffer = "VERSION 5.00" & vbCrLf
                            Else
                                strBuffer = "VERSION 4.00" & vbCrLf

                            End If
                            'List Form Objects
                            If VBVersion <> 4 Then
                                Dim iGuid As Integer, iObject As Integer
                                For iObject = 0 To UBound(gExternalObjectHolder)
                                    If gExternalObjectHolder(iObject).strFormName = txtFinal(i).tag Then
                                        For iGuid = 0 To UBound(gOcxList)
                                            If gOcxList(iGuid).strLibname = gExternalObjectHolder(iObject).strLibname Then
                                                Dim strGuid As String
                                                Dim strVersion As String
                                                'TypeLib
                                                strGuid = modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & gOcxList(iObject).strGuid & "\TypeLib", "")
                                                'Version
                                                strVersion = modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & gOcxList(iObject).strGuid & "\Version", "")
                                                strBuffer = strBuffer & "Object = " & cQuote & "" & strGuid & "#" & strVersion & "#0" & cQuote & "; " & cQuote & gOcxList(iGuid).strocxName & cQuote & vbCrLf
                                                Exit For
                                            End If
                                        Next
                                      
                                    End If
                                Next
                            End If
                            'Show form code
                            strBuffer = strBuffer & txtFinal(i).Text & vbCrLf
                            strBuffer = strBuffer & "Option Explicit" & vbCrLf
                            strBuffer = strBuffer & "'Generated by Semi VB Decompiler - VisualBasicZone.com" & vbCrLf
                            If gProjectInfo.aNativeCode = 0 Then
                                strBuffer = strBuffer & "'This application is compiled to P-Code for code listing refer to Procedures section or P-Code Procedure Decompile under the Tools Menu." & vbCrLf
                            Else
                                strBuffer = strBuffer & "'This application is compiled to Native refer to Native Procedure Decompile under the Tools Menu" & vbCrLf
                            End If
                            
                            If VBVersion <> 4 Then
                                For nApi = 0 To UBound(gProcedureList)
                                    If UCase$(tblPath(2)) = UCase$(gProcedureList(nApi).strParent) And gProcedureList(nApi).strProcedureName <> "" Then
                                        If Right$(gProcedureList(nApi).strProcedureName, 1) = ")" Then
                                            strBuffer = strBuffer & "Private Sub " & gProcedureList(nApi).strProcedureName & vbCrLf
                                        Else
                                            strBuffer = strBuffer & "Private Sub " & gProcedureList(nApi).strProcedureName & "()" & vbCrLf
                                        End If
                                        strBuffer = strBuffer & "End Sub" & vbCrLf
                                    End If
                                Next
                            End If
                            txtCode.Text = strBuffer
                            Exit For
                            
        
                        End If
                    Next
                        
                        If tblPath(3) = vbNullString Then
                            For g = 0 To UBound(gObjectOffsetArray)
                                If UCase$(gObjectOffsetArray(g).ObjectName) = UCase$(tblPath(2)) Then
                                    sstViewFile.Tab = 3
                                    'MsgBox gObjectOffsetArray(g).Address
                                    lblHelpText.Caption = ""
                                    Call modControls.GetControlProperties(gObjectOffsetArray(g).Address)
                                    Exit For
                                End If
                            Next
                        Else
                        'Get Control Properties
                       ' MsgBox UCase(tblPath(3))
                        
                        
                            For g = 0 To UBound(gControlOffset)
                                If UCase$(tblPath(2)) = UCase$(gControlOffset(g).Owner) And UCase$(gControlOffset(g).ControlName) = UCase$(tblPath(3)) Then
                                    sstViewFile.Tab = 3
                                    lblHelpText.Caption = ""
                                    Call modControls.GetControlProperties(gControlOffset(g).offset)
                                    Exit For
                                End If
                                DoEvents
                            Next
                        End If
                   
                    gUpdateText = True
                    txtCode_Change
                    gUpdateText = False
                End If
            Case "USERCONTROL"
                If tblPath(2) <> vbNullString Then
                    sstViewFile.TabVisible(0) = True
                    sstViewFile.TabVisible(3) = True
                    sstViewFile.TabVisible(1) = False
                    sstViewFile.TabVisible(2) = False
                    fxgEXEInfo.Visible = False
                    strBuffer = vbNullString
                    For i = 0 To txtFinal.UBound
                        If UCase$(txtFinal(i).tag) = tblPath(2) Then
                            strBuffer = txtFinal(i).Text
                            strBuffer = strBuffer & "Option Explicit" & vbCrLf
                            strBuffer = strBuffer & "'Generated by Semi VB Decompiler - VisualBasicZone.com" & vbCrLf
                            For nApi = 0 To UBound(gProcedureList)
                                If UCase$(tblPath(2)) = UCase$(gProcedureList(nApi).strParent) And gProcedureList(nApi).strProcedureName <> "" Then
                                    If Right$(gProcedureList(nApi).strProcedureName, 1) = ")" Then
                                        strBuffer = strBuffer & "Private Sub " & gProcedureList(nApi).strProcedureName & vbCrLf
                                    Else
                                        strBuffer = strBuffer & "Private Sub " & gProcedureList(nApi).strProcedureName & "()" & vbCrLf
                                    End If
                                    strBuffer = strBuffer & "End Sub" & vbCrLf
                                End If
                            Next
                            txtCode.Text = strBuffer
                            Exit For
                            

                            End If
                         Next
                        
                   
                    gUpdateText = True
                    txtCode_Change
                    gUpdateText = False
                End If
                
                Case "MODS"
                If tblPath(2) <> vbNullString Then
                    sstViewFile.TabVisible(0) = True
                    sstViewFile.TabVisible(1) = False
                    sstViewFile.TabVisible(2) = False
                    sstViewFile.TabVisible(3) = False
                    fxgEXEInfo.Visible = False
                    txtCode.Text = vbNullString
                    strBuffer = ""
                    strBuffer = strBuffer & "Option Explicit" & vbCrLf
                    strBuffer = strBuffer & "'Generated by Semi VB Decompiler - VisualBasicZone.com" & vbCrLf
                    If gProjectInfo.aNativeCode = 0 Then
                        strBuffer = strBuffer & "'This application is compiled to P-Code for code listing refer to Procedures section or P-Code Procedure Decompile under the Tools Menu." & vbCrLf
                    Else
                        strBuffer = strBuffer & "'This application is compiled to Native refer to Native Procedure Decompile under the Tools Menu" & vbCrLf
                    End If
                    For nApi = 0 To UBound(gProcedureList)
                        If UCase$(tblPath(2)) = UCase$(gProcedureList(nApi).strParent) Then
                            If gProcedureList(nApi).strProcedureName <> "" Then
                                strBuffer = strBuffer & "Private Sub " & gProcedureList(nApi).strProcedureName & "()" & vbCrLf
                                strBuffer = strBuffer & "End Sub" & vbCrLf
                            End If
                        End If
                    Next
                    txtCode.Text = strBuffer
                    gUpdateText = True
                    txtCode_Change
                    gUpdateText = False
                  End If
                Case "CLASS"
                If tblPath(2) <> vbNullString Then
                    sstViewFile.TabVisible(0) = True
                    sstViewFile.TabVisible(1) = False
                    sstViewFile.TabVisible(2) = False
                    sstViewFile.TabVisible(3) = False
                    fxgEXEInfo.Visible = False
                    txtCode.Text = vbNullString
                    strBuffer = vbNullString
                    strBuffer = strBuffer & "Option Explicit" & vbCrLf
                    strBuffer = strBuffer & "'Generated by Semi VB Decompiler - VisualBasicZone.com" & vbCrLf
                    If gProjectInfo.aNativeCode = 0 Then
                        strBuffer = strBuffer & "'This application is compiled to P-Code for code listing refer to Procedures section or P-Code Procedure Decompile under the Tools Menu." & vbCrLf
                    Else
                        strBuffer = strBuffer & "'This application is compiled to Native refer to Native Procedure Decompile under the Tools Menu" & vbCrLf
                    End If
                    
                    For nApi = 0 To UBound(gProcedureList)
                        If UCase$(tblPath(2)) = UCase$(gProcedureList(nApi).strParent) Then
                            If gProcedureList(nApi).strProcedureName <> "" Then
                                strBuffer = strBuffer & "Private Sub " & gProcedureList(nApi).strProcedureName & "()" & vbCrLf
                                strBuffer = strBuffer & "End Sub" & vbCrLf
                            End If
                        End If
                    Next
                    txtCode.Text = strBuffer
                    gUpdateText = True
                    txtCode_Change
                    gUpdateText = False
                End If
                
                Case "USERDOC"
                If tblPath(2) <> vbNullString Then
                    sstViewFile.TabVisible(0) = True
                    sstViewFile.TabVisible(3) = True
                    sstViewFile.TabVisible(1) = False
                    sstViewFile.TabVisible(2) = False
                    fxgEXEInfo.Visible = False
                    strBuffer = vbNullString
                    For i = 0 To txtFinal.UBound
                        If UCase$(txtFinal(i).tag) = tblPath(2) Then
                            strBuffer = txtFinal(i).Text
                            strBuffer = strBuffer & "Option Explicit" & vbCrLf
                            strBuffer = strBuffer & "'Generated by Semi VB Decompiler - VisualBasicZone.com" & vbCrLf

                            txtCode.Text = strBuffer
                            Exit For
                        End If
                    Next
                        
                   
                    gUpdateText = True
                    txtCode_Change
                    gUpdateText = False
                End If
                
                Case "IMAGES"
                'Image Preview
                If tblPath(2) <> vbNullString Then
                sstViewFile.TabVisible(2) = True
                    sstViewFile.TabVisible(0) = False
                    sstViewFile.TabVisible(1) = False
                    sstViewFile.TabVisible(3) = False
                    
                    fxgEXEInfo.Visible = False
             
                    For i = 0 To UBound(FrxPreview) - 1
                        If UCase$(tblPath(2)) = UCase$(FrxPreview(i).strPath) Then
                            On Error Resume Next
                            picPreview.ToolTipText = "Size of Image = " & FileLen(App.Path & "\dump\" & SFile & "\" & FrxPreview(i).strPath & " (in bytes)")
                            picPreview.Picture = LoadPicture(App.Path & "\dump\" & SFile & "\" & FrxPreview(i).strPath)
                        End If
                    Next i
                End If
                  
        End Select
    
     
    End If
End Sub
Private Sub SetupTreeView()
'*****************************
'Purpose: Sets up all the nodes in the Treeview control
'*****************************
'On Error GoTo errHandle
On Error Resume Next
    Dim strParent As String, Filename As String
    Dim i As Long, g As Long
    Filename = SFile
    Dim lMainIcon As Long
    If bISVBNET = False Then
        lMainIcon = 34
    Else
        lMainIcon = 48
    End If
    
    Call tvProject.Nodes.Add(, , "ROOT/PROJECT/" & Filename, Mid$(Filename, InStrRev(Filename, "\") + 1), lMainIcon)
    
    tvProject.Nodes(1).Selected = True
    tvProject.Nodes(1).Expanded = True
    tvProject_NodeClick tvProject.Nodes(1)
    
    '####################   Information about the exe  ####################'
   
    If bNEFormat = False Then
        Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/EXEDATA/", Lang.strTreePEHEADER, 1)
        strParent = "ROOT/EXEDATA/"
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/EXEDATA/EXEHEADER/", "EXE Header", 2
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/EXEDATA/COFFHEADER/", "Coff Header", 2
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/EXEDATA/OPTIONALHEADER/", "Optional Header", 2
    Else
        Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/EXEDATA/", "NE Header", 1)
        strParent = "ROOT/EXEDATA/"
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/EXEDATA/EXEHEADER/", "EXE Header", 2
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/EXEDATA/NEHEADER/", "Ne Header", 2
    End If
        

    
    'IF PE application then
    If bNEFormat = False Then
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/EXEDATA/SECTIONHEADER/", "Section Header", 3
        For i = 1 To PEHeader.NumSections
         tvProject.Nodes.Add "ROOT/EXEDATA/SECTIONHEADER/", tvwChild, "ROOT/EXEDATA/SECTIONHEADER/" & i & "/", SecHeader(i).SecName, 2
        Next i
    End If

    'If VB.net show its CLR header
    If bISVBNET = True Then
        Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/NETSTRUCT/", ".Net Structures", 1)
        strParent = "ROOT/NETSTRUCT/"
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/NETSTRUCT/NETHEADER/", ".Net CLR Header", 2
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/NETSTRUCT/METADATA/", "MetaData Header", 2
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/NETSTRUCT/STREAMS/", "MetaData Streams", 2
        strParent = "ROOT/NETSTRUCT/STREAMS/"
        For i = 0 To UBound(gVBNetStreamHeaders)
            tvProject.Nodes.Add strParent, tvwChild, "ROOT/NETSTRUCT/STREAMS/" & gVBNetStreamHeaders(i).rcName & "/", gVBNetStreamHeaders(i).rcName, 2
        Next
        strParent = "ROOT/NETSTRUCT/"
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/NETSTRUCT/STRINGHEAP", "#String Heap", 2
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/NETSTRUCT/BLOBHEAP", "#Blob Heap", 2
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/NETSTRUCT/GUIDHEAP", "#GUID Heap", 2
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/NETSTRUCT/USHEAP", "User Strings Heap", 2
    End If
    If gVB4App = True Or gVB5App = True Or gVB6App = True Then
        '####################   VB Structures       ####################'
       
        Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/STRUCT/", Lang.strTreeVBStrucutres, 1)
        strParent = "ROOT/STRUCT/"
        
        If gVB4App = True Then
        'Show VB4 Structures
            tvProject.Nodes.Add strParent, tvwChild, "ROOT/STRUCT/VB4HEADER/", "VB4 Header", 2
        Else
            tvProject.Nodes.Add strParent, tvwChild, "ROOT/STRUCT/VBHEADER/", Lang.strTreeVBHEADER, 2
          
            tvProject.Nodes.Add strParent, tvwChild, "ROOT/STRUCT/VBPROJECTINFO/", Lang.strTreeVBProjectInformation, 2
           
            tvProject.Nodes.Add strParent, tvwChild, "ROOT/STRUCT/VBCOMREGDATA/", Lang.strTreeVBComRegistrationData, 2
           
            tvProject.Nodes.Add strParent, tvwChild, "ROOT/STRUCT/VBCOMREGINFO/", "VB COM Registration Info", 2
          
            tvProject.Nodes.Add strParent, tvwChild, "ROOT/STRUCT/VBOBJECTABLE/", Lang.strTreeVBObjectTable, 2
        
            tvProject.Nodes.Add strParent, tvwChild, "ROOT/STRUCT/VBOBJECTS/", Lang.strTreeVBObjects, 2
            For i = 0 To UBound(gObject)
             tvProject.Nodes.Add "ROOT/STRUCT/VBOBJECTS/", tvwChild, "ROOT/STRUCT/VBOBJECTS/" & i & "/", "Object: " & gObjectNameArray(i), 2
             tvProject.Nodes.Add "ROOT/STRUCT/VBOBJECTS/" & i & "/", tvwChild, "ROOT/STRUCT/VBOBJECTS/OBJINFO/" & i & "/", "ObjectInfo", 2
            Next
        End If
    End If
    If gVB4App = True Or gVB5App = True Or gVB6App = True Then
    '####################   VB Forms       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/FORMS/", Lang.strTreeForms, 1)
    strParent = "ROOT/FORMS/"
    If gVB4App = True Then
        For i = 0 To UBound(strVB4Forms)
            If strVB4Forms(i) <> "" Then
                tvProject.Nodes.Add strParent, tvwChild, "ROOT/FORMS/" & UCase$(strVB4Forms(i)) & "/", strVB4Forms(i), 10
            End If
        Next
    Else
    
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 1 Then
    
                    tvProject.Nodes.Add strParent, tvwChild, "ROOT/FORMS/" & UCase$(gObjectNameArray(i)) & "/", gObjectNameArray(i), 10
                    tvProject.Nodes.Add strParent & UCase$(gObjectNameArray(i)) & "/", 4, "ROOT/FORMS/" & UCase$(gObjectNameArray(i)) & "/SUBCOUNT/", "Methods Count: " & gObjectProcCountArray(i), 45
                End If
            Next g
        Next
        'For i = 1 To UBound(gControlNameArray)
        '    If gControlNameArray(i).strControlName <> vbNullString And gControlNameArray(i).strControlName <> "Form" Then
        '        On Error Resume Next
        '        tvProject.Nodes.Add "ROOT/FORMS/" & UCase(gControlNameArray(i).strParentForm) & "/", tvwChild, "ROOT/FORMS/" & UCase(gControlNameArray(i).strParentForm) & "/" & gControlNameArray(i).strControlName & "/", gControlNameArray(i).strControlName, 2
        '    End If
        'Next
        For i = 0 To UBound(gControlOffset)
            If gControlOffset(i).ControlName <> vbNullString Then
                On Error Resume Next
                tvProject.Nodes.Add "ROOT/FORMS/" & UCase$(gControlOffset(i).Owner) & "/", tvwChild, "ROOT/FORMS/" & UCase$(gControlOffset(i).Owner) & "/" & gControlOffset(i).ControlName & "/", gControlOffset(i).ControlName, GetControlPicture(gControlOffset(i).ControlType)
            End If
        Next
    End If
    End If
    If gVB5App = True Or gVB6App = True Then
    '####################   VB Modules       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/MODS/", Lang.strTreeModules, 1)
    strParent = "ROOT/MODS/"
    AppData.AppModuleCount = 0
    For i = 0 To UBound(gObject)
        For g = 0 To UBound(gObjectTypeList)
            If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 2 Then
                tvProject.Nodes.Add strParent, tvwChild, "ROOT/MODS/" & UCase$(gObjectNameArray(i)) & "/", gObjectNameArray(i), 40
                tvProject.Nodes.Add strParent & UCase$(gObjectNameArray(i)) & "/", 4, "ROOT/MODS/" & UCase$(gObjectNameArray(i)) & "/SUBCOUNT/", "Methods Count: " & gObjectProcCountArray(i), 45
                AppData.AppModuleCount = AppData.AppModuleCount + 1
            End If
        Next g
    Next
    '####################   VB Classes       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/CLASS/", Lang.strTreeClasses, 1)
    strParent = "ROOT/CLASS/"
    For i = 0 To UBound(gObject)
        'If gObject(i).ObjectType = 1146883 Then
        For g = 0 To UBound(gObjectTypeList)
            If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 3 Then
                tvProject.Nodes.Add strParent, tvwChild, "ROOT/CLASS/" & UCase$(gObjectNameArray(i)) & "/", gObjectNameArray(i), 41
                tvProject.Nodes.Add strParent & UCase$(gObjectNameArray(i)) & "/", 4, "ROOT/CLASSS/" & UCase$(gObjectNameArray(i)) & "/SUBCOUNT/", "Methods Count: " & gObjectProcCountArray(i), 45
            End If
        Next g
    Next
    '####################   User Controls       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/USERCONTROL/", Lang.strTreeUserControls, 1)
    strParent = "ROOT/USERCONTROL/"

    For i = 0 To UBound(gObject)
        For g = 0 To UBound(gObjectTypeList)
            If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 4 Then
                tvProject.Nodes.Add strParent, tvwChild, "ROOT/USERCONTROL/" & UCase$(gObjectNameArray(i)) & "/", gObjectNameArray(i), 43
                tvProject.Nodes.Add strParent & UCase$(gObjectNameArray(i)) & "/", 4, "ROOT/USERCONTROL/" & UCase$(gObjectNameArray(i)) & "/SUBCOUNT/", "Methods Count: " & gObjectProcCountArray(i), 45
            End If
        Next g
    Next
    '####################   Property Pages       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/PROPERTYPAGE/", Lang.strTreePropertyPages, 1)
    strParent = "ROOT/PROPERTYPAGE/"
    For i = 0 To UBound(gObject)
        For g = 0 To UBound(gObjectTypeList)
            If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 5 Then
            'If gObject(i).ObjectType = 1409027 Then
                tvProject.Nodes.Add strParent, tvwChild, "ROOT/PROPERTYPAGE/" & UCase$(gObjectNameArray(i)) & "/", gObjectNameArray(i), 42
            End If
        Next g
    Next
    '####################   User Documents     ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/USERDOC/", "User Documents", 1)
    strParent = "ROOT/USERDOC/"
    For i = 0 To UBound(gObject)
        For g = 0 To UBound(gObjectTypeList)
            If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 6 Then
                tvProject.Nodes.Add strParent, tvwChild, "ROOT/USERDOC/" & UCase$(gObjectNameArray(i)) & "/", gObjectNameArray(i), 46
            End If
        Next g
    Next
    '####################   Procedures - Code       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/CODE/", Lang.strTreeProceduresCode, 1)
    strParent = "ROOT/CODE/"
    'Add P-Code View
    If gProjectInfo.aNativeCode = 0 Then
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/CODE/" & "PCODE", "View P-Code", 4
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/CODE/" & "PCODETOVB", "View P-Code To VB Code (Beta)", 4
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/CODE/" & "PCODESTRINGS", "P-Code Stings List", 4
    End If
    
    tvProject.Nodes.Add strParent, tvwChild, "ROOT/CODE/" & "API", "API List", 4
    'If project is Native Code
    If gProjectInfo.aNativeCode <> 0 Then
        tvProject.Nodes.Add(strParent, tvwChild, "ROOT/CODE/" & "ASM", "Code assembly for Native", 4).tag = -2
    End If
    
    '####################   Images     ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/IMAGES/", Lang.strTreeImages, 1)
    strParent = "ROOT/IMAGES/"
    For i = 0 To UBound(FrxPreview) - 1
        tvProject.Nodes.Add strParent, tvwChild, "ROOT/IMAGES/" & UCase$(FrxPreview(i).strPath) & "/", FrxPreview(i).strPath, 6
    Next
   
    '####################   File Version Information    ####################'
     Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/VERSIONINFO/", Lang.strTreeFileVersionInformation, 1)
    End If 'End of vbinfo
    '####################   Other Information    ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & Filename, tvwChild, "ROOT/OTHER/", Lang.strTreeImportInformation, 1)
    strParent = "ROOT/OTHER/"
    
        Dim TDs As String, ouR As String
        J = 1
          
            tvProject.Nodes.Add strParent, 4, LCase$(ImportList(0).strName), ImportList(0).strName, 44
                k = UBound(exeIMPORT_APINAME())
                Do While J <= k
                
                    tvProject.Nodes.Add LCase$(ImportList(0).strName), 4, exeIMPORT_APINAME(J).ApiName, exeIMPORT_APINAME(J).ApiName, 44
                    tvProject.Nodes.Add exeIMPORT_APINAME(J).ApiName, 4, , "Offset " & Hex(exeIMPORT_APINAME(J).Address) & "h", 44
                    
                    If Left$(LCase$(ImportList(0).strName), 8) = "msvbvm60" Then
                    If Left$(exeIMPORT_APINAME(J).ApiName, 8) = "!ordinal" Then
                        'via ordinal
                        TDs = VBFunction_Description(Val(Mid$(exeIMPORT_APINAME(J).ApiName, 12)), vbNullString, ouR)
                        If TDs = "undef" Then
                            tvProject.Nodes.Add exeIMPORT_APINAME(J).ApiName, 4, , "Name : " & ouR, 18
                        Else
                            tvProject.Nodes.Add exeIMPORT_APINAME(J).ApiName, 4, , "Name: " & ouR, 18
                          
                            tvProject.Nodes.Add exeIMPORT_APINAME(J).ApiName, 4, , TDs, 19
                        End If
                    Else
                        'via directname
                        TDs = VBFunction_Description(0, exeIMPORT_APINAME(J).ApiName, ouR)
                        If TDs = "undef" Then
                        Else
                            tvProject.Nodes.Add exeIMPORT_APINAME(J).ApiName, 4, , TDs, 19
                        End If
                    End If
                    End If
                    J = J + 1
                Loop
            


    CurrentNode = 0

    tvProject_NodeClick tvProject.Nodes(1)
Exit Sub
errHandle:
    MsgBox "Error_frmMain_SetupTreeView: " & err.Number & " " & err.Description
End Sub

Private Function DumpObject(ByVal FileNum As Variant, ObjectName As String, length As Long, FileStart As Long, HeaderEnd As Long) As Long
'*****************************
'Purpose: Dumps a Gui Object
'*****************************
  On Error GoTo bad
    MakeDir (App.Path & "\dump")
    MakeDir (App.Path & "\dump\" & SFile)
    Dim bArray() As Byte
    ReDim bArray(length)
    'Get the ojbect information
    Seek FileNum, FileStart + 1
    
    Get FileNum, , bArray
    Dim fFileEnd As Long
    fFileEnd = Loc(FileNum)
    Seek FileNum, HeaderEnd
    'Save the information
    Dim F As Long
    F = FreeFile
    Open App.Path & "\dump\" & SFile & "\" & ObjectName & ".txt" For Binary Access Write Lock Write As #F
        Put #F, , bArray
    Close #F
    
    DumpObject = (fFileEnd + 1)
    Exit Function
bad:
    MsgBox "Error_frmMain_DumpObject: " & err.Number & " " & err.Description

    DumpObject = -1
Exit Function
End Function

Sub GetStdPicture(ByVal FileNum As Variant, ByVal length As Variant, ByVal strName As String, ByVal ParentForm As String, ByVal fAddress As Long)
'*****************************
'Purpose: To save an STD Picture and detect what kind of picture file it is.
'*****************************
   On Error Resume Next
    Dim picHeader As typePictureHeader
    Dim bPicArray() As Byte
   
    'Get Picture Header
    Get FileNum, , picHeader

    
    Dim strExt As String
    strExt = ".ico"
    length = length - 8
    'MsgBox "Length: " & Length & " " & strName
    If length > 5000000 Then Exit Sub
    If length < 0 Then Exit Sub
    
     ReDim bPicArray(length)
     Get FileNum, , bPicArray

    If bPicArray(0) = 66 And bPicArray(1) = 77 Then
        strExt = ".bmp"
    ElseIf bPicArray(0) = 71 And bPicArray(1) = 73 And bPicArray(2) = 70 Then
        strExt = ".gif"
    ElseIf bPicArray(0) = 0 And bPicArray(2) = 1 Then
        strExt = ".ico"
    ElseIf bPicArray(0) = 0 And bPicArray(2) = 2 Then
        strExt = ".cur"
    ElseIf bPicArray(0) = 255 And bPicArray(1) = 216 Then
        strExt = ".jpg"
    ElseIf bPicArray(0) = 215 And bPicArray(1) = 205 Then
        strExt = ".wmf"
    End If
    
    FrxPreview(UBound(FrxPreview)).strPath = strName & strExt
    FrxPreview(UBound(FrxPreview)).FRXAddress = fAddress
    FrxPreview(UBound(FrxPreview)).length = length
    FrxPreview(UBound(FrxPreview)).ParentForm = ParentForm
    Dim F As Long
    F = FreeFile
    Open App.Path & "\dump\" & SFile & "\" & strName & strExt For Binary Access Write Lock Write As #F
        
        Put #F, , bPicArray

    Close #F
    ReDim Preserve FrxPreview(UBound(FrxPreview) + 1)
End Sub


Sub ProccessControls(F As Variant)
'*****************************
'Purpose: Process Forms And Control Properties
'*****************************
    Dim bFormEndUsed As Boolean
    Dim strCurrentForm As String
    Dim fPos As Long 'Holds current location in the file used for controlheader
    Dim cListIndex As Integer ' Used for COM
    Dim cControlHeader As ControlHeader
    Dim lForm As Long
    Dim FRXAddress As Long
    Dim posNextControl As Long
    Dim MenuCount As Long
    ReDim gObjectOffsetArray(0)
    'Erase existing data
    bFormEndUsed = False


        'If not a VB4App get guitable
        If gVB4App = False Then
            If gVBHeader.FormCount = 0 Then Exit Sub
        
            Seek F, gVBHeader.aGuiTable + 1 - OptHeader.ImageBase
            
            'Get Form table
            If gVBHeader.FormCount > 0 Then
                ReDim gGuiTable(gVBHeader.FormCount - 1)
                'MsgBox Loc(F)
                Get #F, , gGuiTable
                  
            End If

        End If
        'Loop though each form...
        For lForm = 0 To UBound(gGuiTable)
        'MsgBox "FORM Pointer: " & gGuiTable(lForm).aFormPointer
       ' Debug.Print gGuiTable(lForm).aFormPointer
       ' Debug.Print gGuiTable(lForm).lObjectID
        If gVB4App = False Then
            Seek F, gGuiTable(lForm).aFormPointer + 94 - OptHeader.ImageBase
        Else
            Seek F, gGuiTable(lForm).aFormPointer + 10 - OptHeader.ImageBase
        End If
        FRXAddress = 0
        MenuCount = 0
       
       
'Loop from new child control
NewControl:
        fPos = Loc(F)

        Seek F, fPos + 1
        Dim tArray As ArrayTestType
        Get F, , tArray
        Seek F, fPos + 1

        If tArray.arrayflag <> 128 Then
            Get #F, , cControlHeader
    
        Else
            Dim cArrayHeader As ControlArrayHeader
            Get #F, , cArrayHeader
           
            cControlHeader.length = cArrayHeader.length
            cControlHeader.cName = cArrayHeader.cName
            cControlHeader.cType = cArrayHeader.cType
            cControlHeader.cId = cArrayHeader.cId
            
        End If

        posNextControl = fPos + cControlHeader.length + 2
    
       If gDumpData = True Then
        Dim fHeaderEnd As Long
        Dim fControlEnd As Long
        fHeaderEnd = Loc(F)
        'Store each object's information in a file
        fControlEnd = DumpObject(F, cControlHeader.cName, CLng(cControlHeader.length), fPos, fHeaderEnd)
        If gSkipCom = True Then
        'Get all dumps of the controls even though COM is off
            If fControlEnd <> -1 Then
                Seek F, fControlEnd
                GoTo NewControl
            End If
        End If
       End If
       
       
       If gSkipCom = False Then
        Dim tliTypeInfo As TypeInfo 'Used for COM to find information about the properties of the control
        Dim FileLen As Long 'Used to caculate how much father to go in the control
        'Select what type of control it is
        If gShowOffsets = True Then AddText "'Object Offset: " & fPos
       ' MsgBox cControlHeader.cName
        Select Case cControlHeader.cType
            
            Case vbPictureBox '= 0
                cListIndex = 22
               
                Call AddText("Begin VB.PictureBox " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)

            Case vbLabel '= 1
                cListIndex = 14
                Call AddText("Begin VB.Label " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbTextBox ' = 2
                cListIndex = 27
                Call AddText("Begin VB.TextBox " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbFrame '= 3
                cListIndex = 10
                Call AddText("Begin VB.Frame " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbCommandbutton '= 4
                cListIndex = 4
                 
                Call AddText("Begin VB.CommandButton " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbCheckbox '= 5
                cListIndex = 1
                Call AddText("Begin VB.Checkbox " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbOptionbutton     ' = 6
                cListIndex = 21
                Call AddText("Begin VB.Optionbutton " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbComboBox     ' = 7
                cListIndex = 3
                gComboBoxStyle = False
                Call AddText("Begin VB.Combobox " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbListbox     '= 8
                cListIndex = 17
                Call AddText("Begin VB.ListBox " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbHscroll     '= 9
                cListIndex = 12
                Call AddText("Begin VB.HScrollBar " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbVscroll     '= 10
                cListIndex = 32
                Call AddText("Begin VB.VScrollBar " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbTimer     '= 11
                cListIndex = 28
                Call AddText("Begin VB.Timer " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbform     '= 13
                'If gVB4App = True Then
                 '   cListIndex = 28
                'Else
                    cListIndex = 9
                'End If
                strCurrentForm = cControlHeader.cName
                Call modGlobals.LoadNewFormHolder(cControlHeader.cName)
                gIdentSpaces = 0
                Call AddText("Begin VB.Form " & cControlHeader.cName)
                gObjectOffsetArray(UBound(gObjectOffsetArray)).ObjectName = cControlHeader.cName
                gObjectOffsetArray(UBound(gObjectOffsetArray)).Address = fPos
                ReDim Preserve gObjectOffsetArray(UBound(gObjectOffsetArray) + 1)
                txtStatus.Text = txtStatus.Text & "Processing Form:" & strCurrentForm & vbCrLf
                FrameStatus.Refresh
                txtStatus.SelStart = Len(txtStatus)
                txtStatus.Refresh
                gIdentSpaces = 1
                If VBVersion = 4 Then
                    strVB4Forms(UBound(strVB4Forms)) = strCurrentForm
                    ReDim Preserve strVB4Forms(UBound(strVB4Forms) + 1)
                End If
            Case vbDriveListbox     '= 16
                cListIndex = 7
                Call AddText("Begin VB.DriveListbox " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbDirectoryListbox     '= 17
                cListIndex = 6
                Call AddText("Begin VB.DirListBox " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbFileListbox     '= 18
                cListIndex = 8
                Call AddText("Begin VB.FileListBox " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbmenu     '= 19
                cListIndex = 19
                Call AddText("Begin VB.Menu " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbMDIForm     '= 20
                cListIndex = 18
                gIdentSpaces = 0
                strCurrentForm = cControlHeader.cName
                Call modGlobals.LoadNewFormHolder(cControlHeader.cName)
                txtStatus.Text = txtStatus.Text & "Processing MDIForm:" & cControlHeader.cName & vbCrLf
                Call AddText("Begin VB.MDIForm " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1

                 gObjectOffsetArray(UBound(gObjectOffsetArray)).ObjectName = cControlHeader.cName
                 gObjectOffsetArray(UBound(gObjectOffsetArray)).Address = fPos
                ReDim Preserve gObjectOffsetArray(UBound(gObjectOffsetArray) + 1)
                
                 
            Case vbShape     '= 22
                cListIndex = 26
                Call AddText("Begin VB.Shape " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbLine     '= 23
                cListIndex = 16
                Call AddText("Begin VB.Line " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbImage     '= 24
                cListIndex = 13
                Call AddText("Begin VB.Image " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbData     '= 37
                cListIndex = 5
                Call AddText("Begin VB.Data " & cControlHeader.cName)
                 gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbOLE     '= 38
                cListIndex = 20
                Call AddText("Begin VB.OLE " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1
                gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
            Case vbUserControl     '= 40
                cListIndex = 29
                gIdentSpaces = 0
                strCurrentForm = cControlHeader.cName
                Call modGlobals.LoadNewFormHolder(cControlHeader.cName)
                txtStatus.Text = txtStatus.Text & "Processing UserControl:" & cControlHeader.cName & vbCrLf
                Call AddText("Begin VB.UserControl " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1
                gObjectOffsetArray(UBound(gObjectOffsetArray)).ObjectName = cControlHeader.cName
                gObjectOffsetArray(UBound(gObjectOffsetArray)).Address = fPos
                ReDim Preserve gObjectOffsetArray(UBound(gObjectOffsetArray) + 1)
                
            
            Case vbPropertyPage     '= 41
                cListIndex = 24
                strCurrentForm = cControlHeader.cName
                Call modGlobals.LoadNewFormHolder(cControlHeader.cName)
                txtStatus.Text = txtStatus.Text & "Processing PropertyPage:" & strCurrentForm & vbCrLf
                Call AddText("Begin VB.PropertyPage " & cControlHeader.cName)
                gObjectOffsetArray(UBound(gObjectOffsetArray)).ObjectName = cControlHeader.cName
                gObjectOffsetArray(UBound(gObjectOffsetArray)).Address = fPos
                ReDim Preserve gObjectOffsetArray(UBound(gObjectOffsetArray) + 1)
            Case vbUserDocument     '= 42
                cListIndex = 30
                strCurrentForm = cControlHeader.cName
                Call modGlobals.LoadNewFormHolder(cControlHeader.cName)
                Call AddText("Begin VB.UserDocument " & cControlHeader.cName)
                gIdentSpaces = gIdentSpaces + 1
                gObjectOffsetArray(UBound(gObjectOffsetArray)).ObjectName = cControlHeader.cName
                gObjectOffsetArray(UBound(gObjectOffsetArray)).Address = fPos
                ReDim Preserve gObjectOffsetArray(UBound(gObjectOffsetArray) + 1)
                txtStatus.Text = txtStatus.Text & "Processing UserDocument:" & strCurrentForm & vbCrLf
            Case 255 'external control
                Dim strExternObject As String
                strExternObject = GetAllString(F)
                Call AddText("Begin " & strExternObject & " " & cControlHeader.cName & " 'Length:" & cControlHeader.length)
                Call AddExternalObject(strCurrentForm, strExternObject)
                'Load the control view COM if its on the computer
                'Dim iGuid As Integer
                'For iGuid = 0 To UBound(gOcxList)
                '    If gOcxList(iGuid).strLibName = strExternObject Then
                '        Dim strGuid As String
                '        Dim strVersion As String
                '        Dim sVerTemp
                '        'TypeLib
                '        strGuid = modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & gOcxList(iGuid).strGuid & "\TypeLib", "")
                '        'Version
                '        strVersion = modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & gOcxList(iGuid).strGuid & "\Version", "")
                '        sVerTemp = Split(strVersion, ".")
                '        'MsgBox gOcxList(iGuid).strGuid
                '        tliTypeLibInfo.LoadRegTypeLib gOcxList(iGuid).strGuid, sVerTemp(0), sVerTemp(1), 9
                '        Call ProcessTypeLibrary
                '        Exit For
                '    End If
                'Next
             gControlOffset(UBound(gControlOffset)).ControlName = cControlHeader.cName
                gControlOffset(UBound(gControlOffset)).offset = fPos
                gControlOffset(UBound(gControlOffset)).Owner = strCurrentForm
                gControlOffset(UBound(gControlOffset)).ControlType = cControlHeader.cType
                ReDim Preserve gControlOffset(UBound(gControlOffset) + 1)
                  gIdentSpaces = gIdentSpaces + 1
                Seek F, fPos + cControlHeader.length
                GoTo EndLabel
        End Select
         
        Set tliTypeInfo = tliTypeLibInfo.GetTypeInfo(Replace(Replace(lstTypeInfos.List(cListIndex), "<", ""), ">", ""))
        'Use the ItemData in lstTypeInfos to set the SearchData for lstMembers
        tliTypeLibInfo.GetMembersDirect lstTypeInfos.ItemData(cListIndex), lstMembers.hwnd, , , True
        
        FileLen = Loc(F) - fPos
        FileLen = cControlHeader.length - FileLen
        
        Dim bCode As Byte 'holds gui opcode
        Dim varHold As Variant 'Holds the different data types
        Dim strHold As String 'holds the string
        Dim strReturnType As String 'holds the return type

        Do While Loc(F) < (fPos + cControlHeader.length - 2)
       
        'Do Until Loc(f) >= (fPos + cControlHeader.Length - 1)
     
         Get #F, , bCode
         
         FileLen = FileLen - 1
        
         Dim g As Long

        For g = 0 To lstMembers.ListCount - 1
            
    'Process special events
            Dim iFileLength As Long
        Select Case cControlHeader.cType
                Case vbPictureBox: '0
                    iFileLength = modControls.ProccessPictureBox(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbLabel: ' = 1
                    iFileLength = modControls.ProccessLabel(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbTextBox: ' = 2
                    iFileLength = modControls.ProccessTextBox(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbFrame: ' = 3
                     iFileLength = modControls.ProccessFrame(F, bCode)
                     If iFileLength <> -1 Then Exit For
                Case vbCommandbutton: ' = 4
                     iFileLength = modControls.ProccessCommandButton(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbCheckbox: '= 5
                    iFileLength = modControls.ProccessCheckBox(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbOptionbutton: ' = 6
                    iFileLength = modControls.ProccessOption(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbComboBox: ' = 7
                    iFileLength = modControls.ProccessComboBox(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbListbox: ' = 8
                    iFileLength = modControls.ProccessListBox(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbHscroll: ' = 9
                    iFileLength = modControls.ProccessHscroll(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbVscroll: ' = 10
                    iFileLength = modControls.ProccessVscroll(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbTimer: ' = 11
                    iFileLength = modControls.ProccessTimer(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbform: ' = 13
                    iFileLength = modControls.ProccessForm(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbDriveListbox: ' = 16
                    iFileLength = modControls.ProccessDriveListBox(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbDirectoryListbox: ' = 17
                    iFileLength = modControls.ProccessDirListBox(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbFileListbox: ' = 18
                    iFileLength = modControls.ProccessFileListBox(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbmenu: ' = 19
                    iFileLength = modControls.ProccessMenu(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbMDIForm: ' = 20
                    iFileLength = modControls.ProccessForm(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbShape: ' = 22
                    iFileLength = modControls.ProccessShape(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbLine: ' = 23
                   iFileLength = modControls.ProccessLine(F, bCode)
                   If iFileLength <> -1 Then Exit For
                Case vbImage: ' = 24
                    iFileLength = modControls.ProccessImage(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbData: ' = 37
                    iFileLength = modControls.ProccessDataControl(F, bCode)
                    If iFileLength <> -1 Then Exit For
                    
                Case vbOLE: ' = 38
                    iFileLength = modControls.ProccessOLE(F, bCode)
                    If iFileLength <> -1 Then Exit For
                Case vbUserControl: ' = 40
                    
                    iFileLength = modControls.ProccessUserControl(F, bCode)
                   If iFileLength <> -1 Then Exit For
                Case vbPropertyPage: ' = 41
                Case vbUserDocument: ' = 42
                
            End Select
            strReturnType = "n"
            If ReturnGuiOpcode(lstTypeInfos.ItemData(cListIndex), tliInvokeKinds, lstMembers.List(g)) = bCode Then
              Dim strExtraInfo As String
         
                strReturnType = Trim$(ReturnDataType(lstTypeInfos.ItemData(cListIndex), tliInvokeKinds, lstMembers.List(g)))
         
                If gShowOffsets = True Then
                    strExtraInfo = "  ' GuiOpcode: " & bCode & " Offset Dec: " & Loc(F)
                End If


                If InStr(1, strReturnType, "Byte") Then
                    varHold = GetByte2(F)
                    Call AddText(lstMembers.List(g) & " = " & varHold & strExtraInfo)
                    FileLen = FileLen - 1
                    Exit For
                End If
                If InStr(1, strReturnType, "Boolean") Then
                    varHold = GetBoolean(F)
                    If varHold = True Then
                        Call AddText(lstMembers.List(g) & " = " & 0 & strExtraInfo)
                    Else
                        Call AddText(lstMembers.List(g) & " = " & -1 & strExtraInfo)
                    End If
                    Seek F, Loc(F)
                    FileLen = FileLen - 2
                    Exit For
                End If
                If InStr(1, strReturnType, "Integer") Then
                    varHold = gVBFile.GetInteger(Loc(F))
                    'varHold = GetInteger(f)
                    Call AddText(lstMembers.List(g) & " = " & varHold & strExtraInfo)
                    FileLen = FileLen - 2
                    Exit For
                End If
                If InStr(1, strReturnType, "Long") Then
                    varHold = GetLong(F)
                    Call AddText(lstMembers.List(g) & " = " & varHold & strExtraInfo)
                    FileLen = FileLen - 4
                    Exit For
                End If
                
                If InStr(1, strReturnType, "Single") Then
                    varHold = GetSingle(F)
                    Call AddText(lstMembers.List(g) & " = " & varHold & strExtraInfo)
                    FileLen = FileLen - 4
                    Exit For
                End If

                If InStr(1, strReturnType, "String") Then
                    strHold = GetAllString(F)
                    Call AddText(lstMembers.List(g) & " = " & cQuote & strHold & cQuote & strExtraInfo)
                    FileLen = FileLen - Len(strHold) - 3
                    Exit For
                End If

              
                If InStr(1, strReturnType, "stdole.Picture") Then
                    
                    varHold = GetLong(F)
                   
                    If varHold <> -1 Then
                    'MsgBox "Loc:" & Loc(f) & " " & varHold
                        If cControlHeader.cName <> strCurrentForm Then
                            Call GetStdPicture(F, varHold, strCurrentForm & "." & cControlHeader.cName, strCurrentForm, FRXAddress)
                        Else
                            Call GetStdPicture(F, varHold, cControlHeader.cName, strCurrentForm, FRXAddress)
                        End If
                        
                        
                        Call AddText(lstMembers.List(g) & "=" & cQuote & strCurrentForm & ".frx" & cQuote & ":" & PadHex(Hex(FRXAddress), 4) & strExtraInfo)
                        Seek F, Loc(F)
                     On Error GoTo NextFormDec:
                        If varHold < 0 Then
                            
                        varHold = 0
                        GoTo NextFormDec
                        End If
                        If varHold > 10000000 Then
                        varHold = 0
                            GoTo NextFormDec
                        End If
                       FRXAddress = FRXAddress + varHold + 12
                        'Exit Sub
                        'FileLen = FileLen - varHold + 1 ' - 18
                        FileLen = FileLen - 12
                        'MsgBox varHold
                        'MsgBox "FileLen:" & FileLen

                    Else
                        FileLen = FileLen - 4
                    End If
                    Exit For

                End If

                Exit For

            End If
            

         Next
         'Check For Unknown opcode
         If strReturnType = "n" Then
            'UNKOWN Opcode
            Call modGlobals.AddToErrorLog("Unknown OpCode: " & cControlHeader.cName & " Bcode:" & bCode & " CType: " & cControlHeader.cType & " LOC: " & Loc(F))
            Seek F, fPos + cControlHeader.length
            GoTo EndLabel
         End If
         
         
         'Exit the Process controls in case it hangs on a property
         If CancelDecompile = True Then Exit Sub
         DoEvents
        Loop
        
EndLabel:
        'Get the seperator type for the end of the control
        Dim cControlEnd As Integer
        Dim bCheckEnd As Byte
        
        cControlEnd = GetInteger(F)
        'MsgBox cControlEnd & " Loc:" & Loc(F)
       ' MsgBox cControlHeader.cName & " " & cControlEnd & " Pos: " & Loc(f)
        If cControlEnd = vbFormEnd Then
            bFormEndUsed = True
            gIdentSpaces = gIdentSpaces - 1
            Call AddText("End")
         
        End If
        If cControlEnd = vbFormNewChildControl Then
            Seek F, posNextControl
            GoTo NewControl
            
        End If
        If cControlEnd = vbFormChildControl Then 'FF03
          ' MsgBox "Aga" & cControlHeader.cType & " " & cControlHeader.cName
            If cControlHeader.cType <> vbmenu Then
                gIdentSpaces = gIdentSpaces - 1
                Call AddText("End")
                Do
                    Get F, , bCheckEnd
                    If bCheckEnd = 2 And cControlHeader.cType <> vbmenu Then
                        gIdentSpaces = gIdentSpaces - 1
                        Call AddText("End")
                    End If
                    If bCheckEnd = 3 And cControlHeader.cType <> vbmenu Then
                        gIdentSpaces = gIdentSpaces - 1
                        Call AddText("End")
                    End If
                Loop Until bCheckEnd > 3 Or bCheckEnd = 0
                 
                If bCheckEnd <> 4 Then
                    'MsgBox cControlHeader.cName & " " & cControlEnd & " Pos: " & Loc(f) & "next: " & posNextControl
                    Seek F, Loc(F)
                    GoTo NewControl
                End If
            Else
              ' MsgBox MenuCount & " " & cControlHeader.cName
              ' gIdentSpaces = gIdentSpaces - 1
            '  Call AddText("End")
                Dim iMenu As Long
                For iMenu = 0 To MenuCount - 1
                    gIdentSpaces = gIdentSpaces - 1
                    Call AddText("End")
                Next
                MenuCount = 0
                Do
                    Get F, , bCheckEnd
                    If bCheckEnd = 2 Then

                        If IdentNextMenu = False Then
                            MenuCount = MenuCount + 1
                           ' Call AddText("'MenuCount= " & MenuCount)

                        Else
                            gIdentSpaces = gIdentSpaces - 1
                            Call AddText("End")
                            
                            MenuCount = MenuCount + 1
                            gIdentSpaces = gIdentSpaces + 1
                            IdentNextMenu = False
                        End If
                    End If
                Loop Until bCheckEnd > 3 Or bCheckEnd = 0
                If bCheckEnd <> 4 Then
                    'MsgBox cControlHeader.cName & " " & cControlEnd & " Pos: " & Loc(f) & "next: " & posNextControl
                    Seek F, Loc(F)
                    GoTo NewControl
                End If
            End If
            
        End If
        
        If cControlEnd = vbFormExistingChildControl Then 'FF02
            
            If cControlHeader.cType <> vbmenu Then
                 gIdentSpaces = gIdentSpaces - 1
                 Call AddText("End")
                 Do
                     Get F, , bCheckEnd
                     If bCheckEnd = 2 Then
                         gIdentSpaces = gIdentSpaces - 1
                         Call AddText("End")
                     End If
                     If bCheckEnd = 3 Then
                         gIdentSpaces = gIdentSpaces - 1
                         Call AddText("End")
                     End If
                 Loop Until bCheckEnd >= 3 Or bCheckEnd = 0
                     If bCheckEnd = 0 Or bCheckEnd > 5 Then
                        Seek F, Loc(F)
                     End If
                     If bCheckEnd <> 4 Then
                         GoTo NewControl
                     End If
            Else
                If IdentNextMenu = True Then
                    
                    MenuCount = MenuCount + 1
                    'gIdentSpaces = gIdentSpaces + 1
                    IdentNextMenu = False
                Else
                    gIdentSpaces = gIdentSpaces - 1
                    Call AddText("End")
                    
                   ' MenuCount = MenuCount + 1
                    ''Call AddText("'MenuCount= " & MenuCount)
                End If
                GoTo NewControl
            End If
            
        End If
        If cControlEnd = vbFormMenu Then
          'Seek f, posNextControl
            MenuCount = 0 'MenuCount + 2
            'Call AddText("'MenuCount= " & MenuCount)
            GoTo NewControl
            
        End If
        If bFormEndUsed = False Then
            gIdentSpaces = gIdentSpaces - 1
            Call AddText("End")
        End If
    End If 'For gSkipCom

NextFormDec:
    
    Next lForm 'Main Form Loop
'##########################################
'End of Form/Control Properties Loop
'##########################################
End Sub


Private Sub tvProject_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ext As String
    Dim strFile As String
    strFile = Data.Files(1)
    ext = LCase$(Right$(strFile, 3))
    
    If ext <> "dll" And ext <> "exe" And ext <> "ocx" Then
        MsgBox "Only Accepts dlls, exes, and ocxs", vbInformation
        Exit Sub
    Else
        Dim Buffer As String
        Buffer = String$(255, 0)
        GetFileTitle strFile, Buffer, Len(Buffer)
        Buffer = Left$(Buffer, InStr(1, Buffer, Chr$(0)) - 1)
        Call OpenVBExe(strFile, Buffer)
    End If
End Sub

Private Sub txtCode_Change()
'*****************************
'Purpose: Color Coding for the Syntax
'*****************************
    If gUpdateText = False Then Exit Sub
    If gShowColors = False Then Exit Sub
    
    If frmMain.tag = "P" Then
        Call ColorPCode
        Exit Sub
    End If

        
        Dim Texte As String
        Dim SelStart As Long
        Dim CharCursor As FIRSTCHAR_INFO
        Dim StartFind As Long
        Dim LengthFind As Long
        txtBuffer.Text = txtCode.Text
        
        SelStart = txtBuffer.SelStart
        txtBuffer.MousePointer = rtfHourglass

        Texte = txtBuffer.Text
        
        '======================    SelectionChanged    ========================='

        Dim NewsLines() As String
        Dim i As Long, tStartComp As Long, tEndComp As Long
        Dim TrueRtfText As String, BuffRTFText As String, BuffLines() As String
        TrueRtfText = txtBuffer.TextRTF
        
      
        NewsLines = Split(TrueRtfText, vbCrLf)
        '#########################################################################'
        BuffLines = NewsLines
        
        For i = 0 To UBound(NewsLines)
            If Not i > UBound(LinesCheck) Then
                If NewsLines(i) <> LinesCheck(i) Then
                    Exit For
                End If
            Else
                Exit For
            End If
        Next i
        If i - 1 >= 0 Then
            ReDim Preserve BuffLines(0 To i - 1)
        Else
            ReDim Preserve BuffLines(0 To 0)
        End If
        BuffRTFText = Join(BuffLines, vbCrLf)
        buffCodeAv.TextRTF = BuffRTFText & "}"
        tStartComp = Len(buffCodeAv.Text)
        
        BuffRTFText = vbNullString
        For i = 0 To UBound(NewsLines)
            If UBound(NewsLines) - i >= 0 Then
                If NewsLines(UBound(NewsLines) - i) <> LinesCheck(UBound(LinesCheck) - i) Then
                    Exit For
                End If
                BuffRTFText = NewsLines(UBound(NewsLines) - i) & BuffRTFText
            Else
                Exit For
            End If
        Next i
        buffCodeAp.TextRTF = NewsLines(0) & BuffRTFText
        tEndComp = Len(buffCodeAp.Text)
        
        If Len(Texte) - tEndComp - tStartComp > 0 Then
            Text1 = Mid$(txtBuffer.Text, tStartComp + 1, Len(Texte) - tEndComp - tStartComp)
            StartFind = tStartComp
            LengthFind = Len(Texte) - tEndComp - tStartComp
        Else

            StartFind = InStrRev(Texte, vbCrLf, IIf(SelStart = 0, 1, SelStart)) + 1
            If StartFind = 1 Then
                StartFind = 0
            End If
                
            If InStr(SelStart + 1, Texte, vbCrLf) = 0 Then
                LengthFind = Len(Texte)
            Else
                LengthFind = InStr(SelStart + 1, Texte, vbCrLf) - 1
            End If
            
            LengthFind = LengthFind - StartFind
        End If

        '======================= Section of Colors ========================='
        If LengthFind > 0 Then

            
            txtBuffer.SelStart = StartFind
            txtBuffer.SelLength = LengthFind
            txtBuffer.SelColor = vbBlack
            txtBuffer.SelBold = False
            txtBuffer.SelItalic = False
            txtBuffer.SelUnderline = False
            
            'KeyWords to highlight
            ColorWord "beginproperty", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "endproperty", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "begin", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "end", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "sub", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "public sub", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "private sub", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "public function", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "private function", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "function", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "end sub", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "end function", InstrColor, txtBuffer, "/B/I/", StartFind + 1, LengthFind
            ColorWord "dim", InstrColor, txtBuffer, , StartFind + 1, LengthFind
           ' ColorWord "if", InstrColor, txtBuffer, "/I/", StartFind + 1, LengthFind
           'ColorWord "else", InstrColor, txtBuffer, "/I/", StartFind + 1, LengthFind
           ' ColorWord "elseif", InstrColor, txtBuffer, "/I/", StartFind + 1, LengthFind
           ' ColorWord "then", InstrColor, txtBuffer, "/I/", StartFind + 1, LengthFind
           ' ColorWord "end if", InstrColor, txtBuffer, "/I/", StartFind + 1, LengthFind
           ' ColorWord "goto", InstrColor, txtBuffer, , StartFind + 1, LengthFind
           ' ColorWord "while", InstrColor, txtBuffer, , StartFind + 1, LengthFind
           ' ColorWord "wend", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            'ColorWord "for", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            'ColorWord "next", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            'ColorWord "not", InstrColor, txtBuffer, , StartFind + 1, LengthFind
            'ColorWord "print", FuncColor, txtBuffer, , StartFind + 1, LengthFind
            
            txtBuffer.SelStart = StartFind
            CharCursor = GetFirstChar(txtBuffer.SelStart + 1, txtBuffer, """'")
   
            While CharCursor.lCursor < StartFind + LengthFind And CharCursor.lCursor > 0
                

                    Select Case CharCursor.sChar
                        Case """"
                            Dim InStrFind As Long
                            InStrFind = txtBuffer.Find("""", CharCursor.lCursor) + 1
                            
                            InStrFind = IIf(InStrFind < txtBuffer.Find(vbCrLf, CharCursor.lCursor) + 1, InStrFind, txtBuffer.Find(vbCrLf, CharCursor.lCursor) + 1)
                            txtBuffer.SelStart = CharCursor.lCursor - 1
                            If InStrFind > 0 Then
                                txtBuffer.SelLength = InStrFind - CharCursor.lCursor + 1
                                CharCursor.lCursor = InStrFind
                            Else
                                txtBuffer.SelLength = Len(Texte) - CharCursor.lCursor + 1
                                CharCursor.lCursor = Len(Texte) + 1
                            End If
                            txtBuffer.SelBold = False
                            txtBuffer.SelItalic = False
                            txtBuffer.SelUnderline = False
                            txtBuffer.SelColor = StringColor
                        Case "'"
                            txtBuffer.SelStart = CharCursor.lCursor - 1
                                If InStr(CharCursor.lCursor + 1, Texte, vbCrLf) > 0 Then
                                Dim Buff As String, Counter As Long, TheLen As Long
                                TheLen = InStr(CharCursor.lCursor + 1, Texte, vbCrLf) - CharCursor.lCursor + 1
                                Buff = Mid$(txtBuffer.Text, txtBuffer.SelStart + 1)
                                Counter = 1
                                While Mid$(Trim$(GetPart(Buff, Counter, vbCrLf)), 1, 1) = "'"
                                    TheLen = TheLen + Len(GetPart(Buff, Counter, vbCrLf)) '+ 2
                                    Counter = Counter + 1
                                Wend
                                txtBuffer.SelLength = TheLen
                                CharCursor.lCursor = InStr(CharCursor.lCursor + 1, Texte, vbCrLf) + IIf(Counter > 1, TheLen, 0) + 1
                            Else
                                txtBuffer.SelLength = Len(Texte) - CharCursor.lCursor + 1
                                CharCursor.lCursor = Len(Texte)
                            End If
                            txtBuffer.SelBold = False
                            txtBuffer.SelItalic = True
                            txtBuffer.SelUnderline = False
                            txtBuffer.SelColor = CommentColor
                                                
                        End Select
                    CharCursor = GetFirstChar(CharCursor.lCursor + 1, txtBuffer, """'")
              
            Wend
        End If

        LinesCheck = Split(txtBuffer.TextRTF, vbCrLf)
        
        txtBuffer.SelStart = SelStart
        txtBuffer.Refresh
        txtBuffer.MousePointer = rtfArrow
        DoEvents
        
        txtCode.TextRTF = txtBuffer.TextRTF
      
End Sub

Public Sub ColorWord(ByVal Word As String, ByVal Color As Long, txtBox As RichTextBox, Optional Style As String, Optional ByVal lCursor As Long, Optional ByVal length As Long)
'*****************************
'Purpose: To color a Keyword
'*****************************
    Dim Cursor As Long
    Cursor = lCursor
    Cursor = txtBox.Find(Word, Cursor - 1, , rtfWholeWord) '- 1
    While IIf(length > 0, (Cursor < lCursor + length) And (Cursor > -1), Cursor > -1)
        txtBox.SelColor = Color
        txtBox.SelBold = IIf(UCase$(Style) Like "*/B/*", True, False)
        txtBox.SelItalic = IIf(UCase$(Style) Like "*/I/*", True, False)
        txtBox.SelUnderline = IIf(UCase$(Style) Like "*/U/*", True, False)
        Cursor = txtBox.Find(Word, Cursor + 1, , rtfWholeWord)
    Wend
End Sub
Sub GetControlSize(ByVal F As Variant)
'*****************************
'Purpose: Get the control size type
'*****************************
    Dim cPosition As typeStandardControlSize
    Dim fPos As Long
    fPos = Loc(F) + 1
    Get F, , cPosition
    If cPosition.cLeft <> -32768 Then
    
        Call AddText("Left = " & cPosition.cLeft)
        Call AddText("Top = " & cPosition.cTop)
        Call AddText("Height = " & cPosition.cHeight)
        Call AddText("Width = " & cPosition.cWidth)
                    
    Else
     Dim cPosition2 As typeStandardControlSize2
     Get F, fPos + 2, cPosition2
     Call AddText("Left = " & cPosition2.cLeft)
     Call AddText("Top = " & cPosition2.cTop)
     Call AddText("Height = " & cPosition2.cHeight)
     Call AddText("Width = " & cPosition2.cWidth)
    
    End If
                
End Sub
Sub GetFontProperty(F As Variant)
'*****************************
'Purpose: Get the font property type.
'*****************************
    Dim cFont As FontType
    Dim bItalic As Boolean, bUnderLine As Boolean, bStrike As Boolean
    bItalic = False
    bUnderLine = False
    bStrike = False
    gIdentSpaces = gIdentSpaces + 1
    Call AddText("BeginProperty Font")
    gIdentSpaces = gIdentSpaces + 1
    Get F, , cFont
                'MsgBox cFont.Weight
    Call AddText("Name = " & cQuote & gVBFile.GetString(Loc(F), cFont.FontLen) & cQuote)
               ' FileLen = FileLen - Len(cFont)
    Call AddText("Size = " & (cFont.Size / 10000))
    Call AddText("Charset = " & cFont.un2)
    Call AddText("Weight = " & cFont.Weight)
                'FileLen = FileLen - cFont.FontLen
            'Font Property Opcodes
            'action=2 italic
            'action4=underline
            'action6 underline+italic=6
            'action10=italic+strickough
            'action=8=strikethough
            '12=underline +strikethough
            '14 =italic+underline+strkethough
                If cFont.action = 2 Then
                    bItalic = True
                End If
                If cFont.action = 4 Then
                    bUnderLine = True
                End If
                If cFont.action = 6 Then
                    bUnderLine = True
                    bItalic = True
                End If
                If cFont.action = 8 Then
                    bStrike = True
                End If
                If cFont.action = 10 Then
                    bItalic = True
                    bStrike = True
                End If
                If cFont.action = 12 Then
                    bUnderLine = True
                    bStrike = True
                End If
                If cFont.action = 14 Then
                    bItalic = True
                    bUnderLine = True
                    bStrike = True
                End If
                
                
                If bItalic = True Then
                    Call AddText("Italic = -1")
                Else
                    Call AddText("Italic = 0")
                End If
                If bUnderLine = True Then
                    Call AddText("Underline = -1")
                Else
                    Call AddText("Underline = 0")
                End If
                If bStrike = True Then
                    Call AddText("Strikethrough = -1")
                Else
                    Call AddText("Strikethrough = 0")
                End If
                gIdentSpaces = gIdentSpaces - 1
                Call AddText("EndProperty")
                gIdentSpaces = gIdentSpaces - 1
              '  Seek f, Loc(f) - 1 ' - 2
                'Seek f, Loc(f) + 3
               ' MsgBox Loc(f)
               If gShowOffsets = True Then
               ' Call AddText("'Offset Font End: " & Loc(f) & " " & GetByte2(f))
               End If
End Sub


Private Sub txtEditArray_Change(index As Integer)
'*****************************
'Purpose: Used to detect changes to the exe from the form editor
'*****************************

    Dim i As Integer
    Dim bUsed As Boolean
    Dim strSpacer As String
    bUsed = False
    If frmMain.txtEditArray(index).tag = "Single" Then
         If IsNumeric(txtEditArray(index).Text) = False Then
            txtEditArray(index).Text = 0
         End If
        For i = 0 To UBound(SingleChange)
            If lblArrayEdit(index).tag = SingleChange(i).offset Then
                SingleChange(i).sSingle = txtEditArray(index).Text
                bUsed = True
            End If
        Next
        If bUsed = False Then
            ReDim Preserve SingleChange(UBound(SingleChange) + 1)
            SingleChange(UBound(SingleChange)).offset = lblArrayEdit(index).tag
            SingleChange(UBound(SingleChange)).sSingle = txtEditArray(index).Text
        End If
         
    End If
    If frmMain.txtEditArray(index).tag = "String" Then
        For i = 0 To UBound(StringChange)
            If lblArrayEdit(index).tag = StringChange(i).offset Then
                strSpacer = ""
                If Len(txtEditArray(index).Text) < txtEditArray(index).MaxLength Then
                strSpacer = Space$(txtEditArray(index).MaxLength - Len(txtEditArray(index).Text))
                End If
                StringChange(i).sString = txtEditArray(index).Text & strSpacer
                bUsed = True
            End If
        Next
        If bUsed = False Then
            ReDim Preserve StringChange(UBound(StringChange) + 1)
            StringChange(UBound(StringChange)).offset = lblArrayEdit(index).tag
            strSpacer = ""
            If Len(txtEditArray(index).Text) < txtEditArray(index).MaxLength Then
            strSpacer = Space$(txtEditArray(index).MaxLength - Len(txtEditArray(index).Text))
            End If
            
            StringChange(UBound(StringChange)).sString = txtEditArray(index).Text & strSpacer
        End If
        
    End If
    If frmMain.txtEditArray(index).tag = "Boolean" Then
        For i = 0 To UBound(BooleanChange)
            If lblArrayEdit(index).tag = BooleanChange(i).offset Then
            
                'BooleanChange(i).bBool = txtEditArray(index).Text
                If LCase$(txtEditArray(index).Text) = "true" Then
                    BooleanChange(UBound(BooleanChange)).bBool = True
                End If
                If LCase$(txtEditArray(index).Text) = "false" Then
                    BooleanChange(UBound(BooleanChange)).bBool = False
                End If
                bUsed = True
            End If
        Next
        If bUsed = False Then
            ReDim Preserve BooleanChange(UBound(BooleanChange) + 1)
            BooleanChange(UBound(BooleanChange)).offset = lblArrayEdit(index).tag
            
            If LCase$(txtEditArray(index).Text) = "true" Then
                BooleanChange(UBound(BooleanChange)).bBool = True
            End If
            If LCase$(txtEditArray(index).Text) = "false" Then
                BooleanChange(UBound(BooleanChange)).bBool = False
            End If
            
        End If
    End If
    If frmMain.txtEditArray(index).tag = "Long" Then
         If IsNumeric(txtEditArray(index).Text) = False Then
            txtEditArray(index).Text = 0
         End If
        For i = 0 To UBound(LongChange)
            If lblArrayEdit(index).tag = LongChange(i).offset Then
                LongChange(i).lLong = txtEditArray(index).Text
                bUsed = True
            End If
        Next
        If bUsed = False Then
            ReDim Preserve LongChange(UBound(LongChange) + 1)
            LongChange(UBound(LongChange)).offset = lblArrayEdit(index).tag
            LongChange(UBound(LongChange)).lLong = txtEditArray(index).Text
        End If
         
    End If
    If frmMain.txtEditArray(index).tag = "Integer" Then
         If IsNumeric(txtEditArray(index).Text) = False Then
            txtEditArray(index).Text = 0
        Else
            If txtEditArray(index).Text < -32000 Or txtEditArray(index).Text > 32000 Then
                txtEditArray(index).Text = 0
            End If
        End If
        For i = 0 To UBound(IntegerChange)
            If lblArrayEdit(index).tag = IntegerChange(i).offset Then
                IntegerChange(i).iInt = txtEditArray(index).Text
                bUsed = True
            End If
        Next
        If bUsed = False Then
            ReDim Preserve IntegerChange(UBound(IntegerChange) + 1)
            IntegerChange(UBound(IntegerChange)).offset = lblArrayEdit(index).tag
            IntegerChange(UBound(IntegerChange)).iInt = txtEditArray(index).Text
        End If

    End If
    If frmMain.txtEditArray(index).tag = "Byte" Then
        If IsNumeric(txtEditArray(index).Text) = False Then
            txtEditArray(index).Text = 0
        Else
            If txtEditArray(index).Text < 0 Or txtEditArray(index).Text > 255 Then
                txtEditArray(index).Text = 0
            End If
        End If

        For i = 0 To UBound(ByteChange)
            If lblArrayEdit(index).tag = ByteChange(i).offset Then
                ByteChange(i).bByte = txtEditArray(index).Text
                bUsed = True
            End If
        Next
        If bUsed = False Then
            ReDim Preserve ByteChange(UBound(ByteChange) + 1)
            ByteChange(UBound(ByteChange)).offset = lblArrayEdit(index).tag
            ByteChange(UBound(ByteChange)).bByte = txtEditArray(index).Text
        End If
    
    End If

    mnuFileSaveExe.Enabled = True
End Sub

Sub LoadLanguageList()
On Error GoTo nofile:
    Dim F As Long
    F = FreeFile
    Dim strData As String
    ReDim gLanguageList(0)
    Open App.Path & "\lang\lang.ini" For Input As #F
        Do While Not EOF(F)
            Line Input #F, strData
            If Left$(UCase$(strData), 7) = "DEFAULT=" Then
                gDefaultLanguage = Right$(strData, Len(strData) - 7)
            End If
            If Left$(UCase$(strData), 5) = "LANG=" Then
                gLanguageList(UBound(gLanguageList)) = Right$(strData, Len(strData) - 5)
                ReDim Preserve gLanguageList(UBound(gLanguageList) + 1)
            End If
        Loop
    Close #F
    ReDim Preserve gLanguageList(UBound(gLanguageList) - 1)
Exit Sub
nofile:
    MsgBox "Error_frmMain_LoadLanguageList: " & err.Number & " " & err.Description
Exit Sub
End Sub
Private Sub LoadLanguage(ByVal strName As String)
On Error GoTo nofile:
    Dim F As Long
    F = FreeFile
    Dim strData As String
    Open App.Path & "\lang\" & strName & ".ini" For Input As #F
        Line Input #F, strData
        Lang.Title = strData
        Me.Caption = Lang.Title & " " & Version
        Line Input #F, strData
        Line Input #F, strData
        mnuFile.Caption = strData
        Line Input #F, strData
        mnuFileOpen.Caption = strData
        Line Input #F, strData
        mnuFileGenerate.Caption = strData
        Line Input #F, strData
        mnuFileSaveExe.Caption = strData
        Line Input #F, strData
        mnuFileExportMemoryMap.Caption = strData
        Line Input #F, strData
        Me.mnuFileLanguage.Caption = strData
        Line Input #F, strData
        Me.mnuFileAntiDecompiler.Caption = strData
        Line Input #F, strData
        Me.mnuFileExit.Caption = strData
        Line Input #F, strData
        Me.mnuOptions.Caption = strData
        Line Input #F, strData
        Me.mnuTools.Caption = strData
        Line Input #F, strData
        Me.mnuToolsPCodeProcedure.Caption = strData
        Line Input #F, strData
        Me.mnuHelp.Caption = strData
        Line Input #F, strData
        Me.mnuHelpReportBug.Caption = strData
        Line Input #F, strData
        Me.mnuHelpAbout.Caption = strData
        Line Input #F, strData
        Line Input #F, strData
        sstViewFile.TabCaption(0) = strData
        Line Input #F, strData
        sstViewFile.TabCaption(1) = strData
        Line Input #F, strData
        sstViewFile.TabCaption(2) = strData
        Line Input #F, strData
        sstViewFile.TabCaption(3) = strData
        'Get spacer
        Line Input #F, strData
        Line Input #F, strData
        Lang.strTreePEHEADER = strData
        Line Input #F, strData
        Lang.strTreeVBStrucutres = strData
        Line Input #F, strData
        Lang.strTreeVBHEADER = strData
        Line Input #F, strData
        Lang.strTreeVBProjectInformation = strData
        Line Input #F, strData
        Lang.strTreeVBComRegistrationData = strData
        Line Input #F, strData
        Lang.strTreeVBObjectTable = strData
        Line Input #F, strData
        Lang.strTreeVBObjects = strData
        Line Input #F, strData
        Lang.strTreeForms = strData
        Line Input #F, strData
        Lang.strTreeModules = strData
        Line Input #F, strData
        Lang.strTreeClasses = strData
        Line Input #F, strData
        Lang.strTreeUserControls = strData
        Line Input #F, strData
        Lang.strTreePropertyPages = strData
        Line Input #F, strData
        Lang.strTreeProceduresCode = strData
        Line Input #F, strData
        Lang.strTreeImages = strData
        Line Input #F, strData
        Lang.strTreeFileVersionInformation = strData
        Line Input #F, strData
        Lang.strTreeImportInformation = strData
    Close #F
   
Exit Sub
nofile:
    MsgBox "Error_frmMain_LoadLanguage: " & err.Number & " " & err.Description
Exit Sub
End Sub
Public Function GetListType(ByVal F As Variant) As Long
    On Error GoTo badlist
    Dim strData As String
    Dim TempDoulbe As Double
    Dim TempInt As Integer
                'Get number of items
                TempDouble# = GetWordByFile(CInt(F))
                
                'Any items?
                If TempDouble# <> 0 Then
                
                     'Skip 2 bytes
                     TempInt% = GetByteByFile(CInt(F))
                     TempInt% = GetByteByFile(CInt(F))
                     
                     'Loop through "TempDouble" items...
                     Do
                     
                         'Get length of item
                         TempInt% = GetWordByFile(CInt(F))
                                   
                         'Loop through "TempInt" chars
                         Do
                         
                             'Get a char
                            strData = strData & Chr$(GetByteByFile(CInt(F)))
                             
                             'Reset text length
                             TempInt% = TempInt% - 1
                             
                         Loop While TempInt% <> 0
                         
                         'Add in a CR
                        strData = strData & Chr$(13)
                         
                         'Reset item count
                         TempDouble# = TempDouble# - 1
                             
                     Loop While TempDouble# > 0
                    
                End If
            'Debug.Print strData
    Exit Function
    Dim iNumberOfItems As Integer
    Dim iNum As Integer
    Dim iLength As Integer
    Dim i As Integer
    Dim bArray() As Byte
    Dim b As Byte
    Dim Counter As Integer
    Counter = 0
    Get F, , iNumberOfItems
    Get F, , iNum
    Counter = 4
    If iNumberOfItems = 0 Then
        Get F, , b
        Counter = Counter + 1
        GetListType = Counter
        Exit Function
    End If
    'Now get list items
    'Dim strData As String
    For i = 1 To iNumberOfItems
        Get F, , iLength
        Counter = Counter + 2
        ReDim bArray(iLength - 1)
        Counter = Counter + iLength
        Get F, , bArray
        
    Next
    Get F, , b
    Get F, , iNum
    Get F, , iNum
    Counter = Counter + 5
   ' MsgBox iNum
    'ItemDatas
   ' MsgBox "LOC: " & Loc(f)
    For i = 1 To iNumberOfItems
        Get F, , iLength
        ReDim bArray(iLength - 1)
        Get F, , bArray
        Counter = Counter + iLength + 2
    Next
'MsgBox "LOC: " & Loc(f)
    GetListType = Counter
    Exit Function
badlist:
    MsgBox "Error_frmMain_GetListType: " & err.Source & " " & err.Description
    Exit Function
End Function

Sub VB3Decompile(ByVal strFileName As String)
On Error GoTo nofile:
    Dim sPath As String
    Dim structFolder As BROWSEINFO
    Dim iNull As Integer
    Dim ret As Long
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
    sPath = sPath & "\"
    'VB3DecompilerOcx1.SetDataPath = App.Path & "\data\vb3\"
    'VB3DecompilerOcx1.OutputFolder = sPath
    'VB3DecompilerOcx1.Filename = strFileName
    'VB3DecompilerOcx1.DecompileFile
Exit Sub
nofile:
    MsgBox "This file is probably protected and can not be decompiled!", vbCritical
    Exit Sub
End Sub
Private Sub ShowVB4Header()
    fxgEXEInfo.ColWidth(0) = 2500
    fxgEXEInfo.TextArray(2) = "Signature"
    fxgEXEInfo.TextArray(3) = VB4Header.sig
    fxgEXEInfo.AddItem "CompilerFileVersion"
    fxgEXEInfo.TextArray(5) = VB4Header.CompilerFileVersion
    fxgEXEInfo.AddItem "i1"
    fxgEXEInfo.TextArray(7) = VB4Header.int1
    fxgEXEInfo.AddItem "i2"
    fxgEXEInfo.TextArray(9) = VB4Header.int2
    fxgEXEInfo.AddItem "i3"
    fxgEXEInfo.TextArray(11) = VB4Header.int3
    fxgEXEInfo.AddItem "i4"
    fxgEXEInfo.TextArray(13) = VB4Header.int4
    fxgEXEInfo.AddItem "i5"
    fxgEXEInfo.TextArray(15) = VB4Header.int5
    fxgEXEInfo.AddItem "i6"
    fxgEXEInfo.TextArray(17) = VB4Header.int6
    fxgEXEInfo.AddItem "i7"
    fxgEXEInfo.TextArray(19) = VB4Header.int7
    fxgEXEInfo.AddItem "i8"
    fxgEXEInfo.TextArray(21) = VB4Header.int8
    fxgEXEInfo.AddItem "i9"
    fxgEXEInfo.TextArray(23) = VB4Header.int9
    fxgEXEInfo.AddItem "i10"
    fxgEXEInfo.TextArray(25) = VB4Header.int10
    fxgEXEInfo.AddItem "i11"
    fxgEXEInfo.TextArray(27) = VB4Header.int11
    fxgEXEInfo.AddItem "i12"
    fxgEXEInfo.TextArray(29) = VB4Header.int12
    fxgEXEInfo.AddItem "i13"
    fxgEXEInfo.TextArray(31) = VB4Header.int13
    fxgEXEInfo.AddItem "i14"
    fxgEXEInfo.TextArray(33) = VB4Header.int14
    fxgEXEInfo.AddItem "i15"
    fxgEXEInfo.TextArray(35) = VB4Header.int15
    fxgEXEInfo.AddItem "LangId"
    fxgEXEInfo.TextArray(37) = VB4Header.LangID
    fxgEXEInfo.AddItem "i16"
    fxgEXEInfo.TextArray(39) = VB4Header.int16
    fxgEXEInfo.AddItem "i17"
    fxgEXEInfo.TextArray(41) = VB4Header.int17
    fxgEXEInfo.AddItem "i18"
    fxgEXEInfo.TextArray(43) = VB4Header.int18
    fxgEXEInfo.AddItem "Address of SubMain"
    fxgEXEInfo.TextArray(45) = VB4Header.aSubMain
    fxgEXEInfo.AddItem "Address"
    fxgEXEInfo.TextArray(47) = VB4Header.Address2
    fxgEXEInfo.AddItem "i1"
    fxgEXEInfo.TextArray(49) = VB4Header.i1
    fxgEXEInfo.AddItem "i2"
    fxgEXEInfo.TextArray(51) = VB4Header.i2
    fxgEXEInfo.AddItem "i3"
    fxgEXEInfo.TextArray(53) = VB4Header.i3
    fxgEXEInfo.AddItem "i4"
    fxgEXEInfo.TextArray(55) = VB4Header.i4
    fxgEXEInfo.AddItem "i5"
    fxgEXEInfo.TextArray(57) = VB4Header.i5
    fxgEXEInfo.AddItem "i6"
    fxgEXEInfo.TextArray(59) = VB4Header.i6
    fxgEXEInfo.AddItem "iExeNameLength"
    fxgEXEInfo.TextArray(61) = VB4Header.iExeNameLength
    fxgEXEInfo.AddItem "iProjectSavedNameLength"
    fxgEXEInfo.TextArray(63) = VB4Header.iProjectSavedNameLength
    fxgEXEInfo.AddItem "iHelpFileLength"
    fxgEXEInfo.TextArray(65) = VB4Header.iHelpFileLength
    fxgEXEInfo.AddItem "iProjectNameLength"
    fxgEXEInfo.TextArray(67) = VB4Header.iProjectNameLength
    fxgEXEInfo.AddItem "FormCount"
    fxgEXEInfo.TextArray(69) = VB4Header.FormCount
    fxgEXEInfo.AddItem "int19"
    fxgEXEInfo.TextArray(71) = VB4Header.int19
    fxgEXEInfo.AddItem "NumberOfExternalComponets"
    fxgEXEInfo.TextArray(73) = VB4Header.NumberOfExternalComponets
    fxgEXEInfo.AddItem "int20"
    fxgEXEInfo.TextArray(75) = VB4Header.int20
    fxgEXEInfo.AddItem "Address of GUI Table"
    fxgEXEInfo.TextArray(77) = VB4Header.aGuiTable
    fxgEXEInfo.AddItem "Address4"
    fxgEXEInfo.TextArray(79) = VB4Header.Address4
    fxgEXEInfo.AddItem "aExternalComponetTable"
    fxgEXEInfo.TextArray(81) = VB4Header.aExternalComponetTable
    fxgEXEInfo.AddItem "Address of Project Info2"
    fxgEXEInfo.TextArray(83) = VB4Header.aProjectInfo2

End Sub

Private Sub ColorPCode()

        Dim Texte As String
        Dim SelStart As Long
        Dim CharCursor As FIRSTCHAR_INFO
        Dim StartFind As Long
        Dim LengthFind As Long
        txtBuffer.Text = txtCode.Text
        
        SelStart = txtBuffer.SelStart
        txtBuffer.MousePointer = rtfHourglass

        Texte = txtBuffer.Text
        
        '======================    SelectionChanged    ========================='
        Dim NewsLines() As String
        Dim i As Long, tStartComp As Long, tEndComp As Long
        Dim TrueRtfText As String, BuffRTFText As String, BuffLines() As String
        TrueRtfText = txtBuffer.TextRTF
        
        '#########################################################################'
      
        NewsLines = Split(TrueRtfText, vbCrLf)      '<<<<<   absolument   <<<<<<<#'
        '#########################################################################'
        BuffLines = NewsLines
        
        For i = 0 To UBound(NewsLines)
            If Not i > UBound(LinesCheck) Then
                If NewsLines(i) <> LinesCheck(i) Then
                    Exit For
                End If
            Else
                Exit For
            End If
        Next i
        If i - 1 >= 0 Then
            ReDim Preserve BuffLines(0 To i - 1)
        Else
            ReDim Preserve BuffLines(0 To 0)
        End If
        BuffRTFText = Join(BuffLines, vbCrLf)
        buffCodeAv.TextRTF = BuffRTFText & "}"
        tStartComp = Len(buffCodeAv.Text)
        
        BuffRTFText = vbNullString
        For i = 0 To UBound(NewsLines)
            If UBound(NewsLines) - i >= 0 Then
                If NewsLines(UBound(NewsLines) - i) <> LinesCheck(UBound(LinesCheck) - i) Then
                    Exit For
                End If
                BuffRTFText = NewsLines(UBound(NewsLines) - i) & BuffRTFText
            Else
                Exit For
            End If
        Next i
        buffCodeAp.TextRTF = NewsLines(0) & BuffRTFText
        tEndComp = Len(buffCodeAp.Text)
        
        If Len(Texte) - tEndComp - tStartComp > 0 Then
            Text1 = Mid$(txtBuffer.Text, tStartComp + 1, Len(Texte) - tEndComp - tStartComp)
            StartFind = tStartComp
            LengthFind = Len(Texte) - tEndComp - tStartComp
        Else
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        

            StartFind = InStrRev(Texte, vbCrLf, IIf(SelStart = 0, 1, SelStart)) + 1
            If StartFind = 1 Then
                StartFind = 0
            End If
                
            If InStr(SelStart + 1, Texte, vbCrLf) = 0 Then
                LengthFind = Len(Texte)
            Else
                LengthFind = InStr(SelStart + 1, Texte, vbCrLf) - 1
            End If
            
            LengthFind = LengthFind - StartFind
        End If
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        '======================================================================='
        
        '======================= Section of Colors ========================='
        If LengthFind > 0 Then

            
            txtBuffer.SelStart = StartFind
            txtBuffer.SelLength = LengthFind 'Len(Texte)
            txtBuffer.SelColor = vbBlack
            txtBuffer.SelBold = True
            txtBuffer.SelItalic = False
            txtBuffer.SelUnderline = False
 
            'KeyWords to highlight
            ColorWord "Branch", vbRed, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "BranchF", vbRed, txtBuffer, , StartFind + 1, LengthFind
            ColorWord "Unknown", vbRed, txtBuffer, , StartFind + 1, LengthFind
    
            'VCallHresult
            ColorWord "VCallHresult", RGB(80, 80, 40), txtBuffer, , StartFind + 1, LengthFind
            'Orange
            ColorWord "ImpAdCallFPR4", RGB(255, 128, 0), txtBuffer, , StartFind + 1, LengthFind
            ColorWord "ThisVCallHresult", RGB(255, 128, 0), txtBuffer, , StartFind + 1, LengthFind
            ColorWord "ImpAdCallI4", RGB(255, 128, 0), txtBuffer, , StartFind + 1, LengthFind
            

            txtBuffer.SelStart = StartFind
            CharCursor = GetFirstChar(txtBuffer.SelStart + 1, txtBuffer, """'")
            
            While CharCursor.lCursor < StartFind + LengthFind And CharCursor.lCursor > 0
            Exit Sub

                    Select Case CharCursor.sChar
                        Case """"
                            Dim InStrFind As Long
                            InStrFind = txtBuffer.Find("""", CharCursor.lCursor) + 1
                            
                            InStrFind = IIf(InStrFind < txtBuffer.Find(vbCrLf, CharCursor.lCursor) + 1, InStrFind, txtBuffer.Find(vbCrLf, CharCursor.lCursor) + 1)
                            txtBuffer.SelStart = CharCursor.lCursor - 1
                            If InStrFind > 0 Then
                                txtBuffer.SelLength = InStrFind - CharCursor.lCursor + 1
                                CharCursor.lCursor = InStrFind
                            Else
                                txtBuffer.SelLength = Len(Texte) - CharCursor.lCursor + 1
                                CharCursor.lCursor = Len(Texte) + 1
                            End If
                            txtBuffer.SelBold = False
                            txtBuffer.SelItalic = False
                            txtBuffer.SelUnderline = False
                            txtBuffer.SelColor = StringColor
                        Case "'"
                            txtBuffer.SelStart = CharCursor.lCursor - 1
                                If InStr(CharCursor.lCursor + 1, Texte, vbCrLf) > 0 Then
                                Dim Buff As String, Counter As Long, TheLen As Long
                                TheLen = InStr(CharCursor.lCursor + 1, Texte, vbCrLf) - CharCursor.lCursor + 1
                                Buff = Mid$(txtBuffer.Text, txtBuffer.SelStart + 1)
                                Counter = 1
                                While Mid$(Trim$(GetPart(Buff, Counter, vbCrLf)), 1, 1) = "'"
                                    TheLen = TheLen + Len(GetPart(Buff, Counter, vbCrLf)) '+ 2
                                    Counter = Counter + 1
                                Wend
                                txtBuffer.SelLength = TheLen
                                CharCursor.lCursor = InStr(CharCursor.lCursor + 1, Texte, vbCrLf) + IIf(Counter > 1, TheLen, 0) + 1
                            Else
                                txtBuffer.SelLength = Len(Texte) - CharCursor.lCursor + 1
                                CharCursor.lCursor = Len(Texte)
                            End If
                            txtBuffer.SelBold = False
                            txtBuffer.SelItalic = True
                            txtBuffer.SelUnderline = False
                            txtBuffer.SelColor = CommentColor
                    End Select
                    CharCursor = GetFirstChar(CharCursor.lCursor + 1, txtBuffer, """'")
              
            Wend
        End If
        '======================================================================='
        
        LinesCheck = Split(txtBuffer.TextRTF, vbCrLf)
        
        txtBuffer.SelStart = SelStart
        txtBuffer.Refresh
        txtBuffer.MousePointer = rtfArrow
        DoEvents
        

        txtCode.TextRTF = txtBuffer.TextRTF
End Sub
'Private Sub pSetIcon( _
'        ByVal sIconKey As String, _
'        ByVal sMenuKey As String _
'    )
'Dim lIconIndex As Long
'    lIconIndex = plGetIconIndex(sIconKey)
'    ctlPopMenu.ItemIcon(sMenuKey) = lIconIndex
'End Sub

'Private Function plGetIconIndex( _
'        ByVal sKey As String _
'    ) As Long
'    plGetIconIndex = ilsIcons.ListImages.Item(sKey).Index - 1
'End Function
Private Function GetCodeStart(Address As Double, F As Integer)

    Dim CodeLength As Double
    Dim TempDouble As Double
    'First get code length....
        
    'Go to offset
    Seek F, Address# + 1
        
    'Move ahead to Length data
    Seek F, Seek(F) + &H8
    
    'Get length
    TempDouble# = GetWordByFile(F)
    
    'Save it
    CodeLength# = TempDouble#
    
    'Get start position
    GetCodeStart = Address# - CodeLength#
    
End Function
Private Function GetProcAddress(Number As Long, F As Integer, iMax As Integer)
                 'Get ProcAddr
    Seek F, gObjectInfo.aProcTable + 1 - OptHeader.ImageBase
                 For addr = 0 To iMax
                    Dim dProc As Double
                     dProc = GetDWordByFile(F)
                    
                       If ((dProc > DecLoadOffset#) And (dProc < ((LOF(F) + DecLoadOffset#)))) Then
                        dProc = dProc - DecLoadOffset#
                        Dim oProcPos As Long
                        oProcPos = Seek(F)
                       ' Debug.Print GetCodeStart(dProc, f) & " #" & addr
                        If addr = Number Then
                            'MsgBox GetCodeStart(dProc, f)
                            GetProcAddress = (GetCodeStart(dProc, F) + OptHeader.ImageBase)
                        End If
                        Seek F, oProcPos
                        
                       End If
                 Next
        
End Function
Private Function GetControlPicture(ByVal bId As Byte) As Long
    
    Select Case bId
        Case vbPictureBox
            GetControlPicture = 21
        Case vbLabel
            GetControlPicture = 15
        Case vbTextBox
            GetControlPicture = 24
        Case vbFrame
            GetControlPicture = 11
        Case vbCommandbutton
            GetControlPicture = 8
        Case vbCheckbox
            GetControlPicture = 7
        Case vbOptionbutton
            GetControlPicture = 20
        Case vbComboBox
            GetControlPicture = 33
        Case vbListbox
            GetControlPicture = 17
        Case vbHscroll
            GetControlPicture = 51
        Case vbVscroll
            GetControlPicture = 28
        Case vbTimer
            GetControlPicture = 25
        Case vbDriveListbox
            GetControlPicture = 30
        Case vbDirectoryListbox
            GetControlPicture = 29
        Case vbFileListbox
            GetControlPicture = 31
        Case vbmenu
            GetControlPicture = 19
        Case vbShape
            GetControlPicture = 32
        Case vbLine
            GetControlPicture = 16
        Case vbImage
            GetControlPicture = 12
        Case vbData
            GetControlPicture = 49
        Case vbOLE
            GetControlPicture = 50
        Case 255
            GetControlPicture = 39
        Case Else
            GetControlPicture = 2
    End Select
End Function




