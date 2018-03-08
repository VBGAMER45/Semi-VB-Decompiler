VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Semi VB Decompiler Api Add"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   6480
      TabIndex        =   11
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add New Function"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   5
      Top             =   2760
      Width           =   8055
      Begin VB.TextBox txtAddFunctionName 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtAddFunctionDef 
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   480
         Width           =   5055
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Function Name"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Function Full Declare"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
   End
   Begin RichTextLib.RichTextBox txtFunctionName 
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3413
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin RichTextLib.RichTextBox txtApiList 
      Height          =   1935
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3413
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0082
   End
   Begin VB.Label Label3 
      Caption         =   "Api Add: Allows you to add more api's to Semi VB Decompiler's API Database"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label lblFull 
      Caption         =   "Full Function Declare"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label lblFunctionName 
      Caption         =   "Function Name:"
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
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblNumber 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Api List
Private Type APIRecord
    Function As String
    Definition As String
End Type
Const APIRecLen As Integer = 536
'Const APIRecTotal As Integer = 1543
Dim APIRecTotal As Integer


Private Sub cmdAdd_Click()
    Dim APIFile As Integer
    Dim iAPILoop As Long
    Dim tAPIRec As APIRecord
    
    If txtAddFunctionName.Text = "" Then
        MsgBox "You need to add a function name!", vbInformation
        Exit Sub
    End If
    If txtAddFunctionDef.Text = "" Then
        MsgBox "You need to add a function declare!", vbInformation
        Exit Sub
    End If
    
    APIFile = FreeFile
    
    
    Open App.Path & "\data\winapi.dat" For Random As #APIFile Len = APIRecLen

            tAPIRec.Function = txtAddFunctionName.Text
            tAPIRec.Definition = txtAddFunctionDef.Text
            APIRecTotal = APIRecTotal + 1
            Put #APIFile, APIRecTotal, tAPIRec

        txtFunctionName.Text = txtFunctionName.Text & Trim$(tAPIRec.Function) & vbCrLf
        txtApiList.Text = txtApiList.Text & Trim$(tAPIRec.Definition) & vbCrLf


    Close #APIFile

    lblNumber.Caption = APIRecTotal
    MsgBox "Record Added", vbInformation
    txtAddFunctionName.Text = ""
    txtAddFunctionDef.Text = ""
    
End Sub



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    APIRecTotal = 1545
    lblNumber.Caption = APIRecTotal
    Me.Caption = "Semi VB Decompiler Api Add - Loading"
    Me.Show
    Call LoadApi
    lblNumber.Caption = "Total API: " & APIRecTotal
    Me.Caption = "Semi VB Decompiler Api Add"
End Sub
Sub LoadApi()
On Error GoTo errHandle
    Dim APIFile As Integer
    Dim iAPILoop As Long
    Dim tAPIRec As APIRecord
    APIFile = FreeFile
    txtFunctionName.Text = ""
    txtApiList.Text = ""
   Dim strBuffer As String, strBuffer2 As String
   Dim Length As Long
   Open App.Path & "\data\winapi.dat" For Random As #APIFile Len = APIRecLen
        Length = LOF(APIFile)
        
        APIRecTotal = Length \ APIRecLen
        APIRecTotal = APIRecTotal + 1
        For iAPILoop = 1 To APIRecTotal
            Get #APIFile, iAPILoop, tAPIRec
            strBuffer = strBuffer & Trim$(tAPIRec.Function) & vbCrLf
            strBuffer2 = strBuffer2 & Trim$(tAPIRec.Definition) & vbCrLf

        Next iAPILoop
        txtFunctionName.Text = strBuffer
        txtApiList.Text = strBuffer2
    Close #APIFile
Exit Sub
errHandle:
MsgBox "LoadApi: " & Err.Description
End Sub


