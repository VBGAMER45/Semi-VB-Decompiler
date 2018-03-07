VERSION 5.00
Object = "{586D49AA-00F4-4C06-B9DA-47784EC936AC}#14.0#0"; "prjVB3Decompiler.ocx"
Begin VB.Form frmTester 
   Caption         =   "VB3 Decompiler"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin prjVB3Decompiler.VB3DecompilerOcx VB3DecompilerOcx1 
      Left            =   720
      Top             =   1440
      _ExtentX        =   3149
      _ExtentY        =   1614
   End
End
Attribute VB_Name = "frmTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Example Decompiling a vb3 application
    VB3DecompilerOcx1.SetDataPath = App.Path & "\vb3\"
    VB3DecompilerOcx1.OutputFolder = App.Path & "\output\"
    VB3DecompilerOcx1.FileName = App.Path & "\vb3.exe"
    VB3DecompilerOcx1.DecompileFile
    
    'Example Decompling a vb2 application
    VB3DecompilerOcx1.SetDataPath = App.Path & "\vb3\"
    VB3DecompilerOcx1.OutputFolder = App.Path & "\output2\"
    VB3DecompilerOcx1.FileName = App.Path & "\vb2.exe"
    VB3DecompilerOcx1.DecompileFile
End Sub

