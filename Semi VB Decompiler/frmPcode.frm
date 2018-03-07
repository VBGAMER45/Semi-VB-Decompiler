VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPcode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P-Code Procedure Decompile View"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddAddress 
      Caption         =   "&Add Address"
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   3120
      TabIndex        =   9
      Text            =   "0"
      ToolTipText     =   "Needs to be a memory address. If you have a file offset use the File Offset to Address convertor."
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Item"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdExportList 
      Caption         =   "E&xport List"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.OptionButton optPCodeToVB 
      Caption         =   "P-Code To VB"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      ToolTipText     =   "Decompiles the selected procedure and attempts to conver the P-Code Tokens to VB Code"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.OptionButton optPCode 
      Caption         =   "P-Code"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      ToolTipText     =   "Decompiles a procedure and shows the P-Code Tokens"
      Top             =   4680
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox txtView 
      Height          =   3570
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   360
      Width           =   4935
   End
   Begin VB.ListBox lstProcedures 
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   360
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSetting 
      Caption         =   "Current Setting:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Just click on a procedure in the list to decompile it."
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblTitle 
      Caption         =   "Procedure List:"
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
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmPcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************
'frmPCode
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************
Option Explicit

Private Sub cmdAddAddress_Click()
    Me.lstProcedures.AddItem txtAddress.Text
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExportList_Click()
    CD1.Filename = ""
    CD1.DefaultExt = ".txt"
    CD1.Filter = "Text Files(*.txt)|*.txt"
    CD1.DialogTitle = "Export Procedure List"
    CD1.ShowSave
    
    If CD1.Filename = "" Then Exit Sub
    
    Dim F As Integer
    F = FreeFile
    Open CD1.Filename For Output As #F
        Print #F, "Semi VB Decompiler - P-Code Procedure List"
        Print #F, "Filename: " & SFilePath
        Print #F, ""
        Print #F, "Procedure Memory Offsets:"
        Dim i As Integer
        For i = 0 To lstProcedures.ListCount - 1
            Print #F, lstProcedures.List(i)
        Next
        
    Close #F
End Sub

Private Sub cmdRemove_Click()
    If lstProcedures.ListIndex = -1 Then
        MsgBox "Please select an event or procedure.", vbInformation
        Exit Sub
    End If
    
    Dim iResponse As Integer
    iResponse = MsgBox("Are you sure you want to remove address: " & lstProcedures.List(lstProcedures.ListIndex), vbYesNo + vbInformation, "Remove Item?")
    If iResponse = vbYes Then
        lstProcedures.RemoveItem (lstProcedures.ListIndex)
    End If
End Sub

Private Sub Form_Load()
'*****************************
'Purpose: To load all events into the listbox
'*****************************
On Error GoTo errHandle
    Dim ProcAddr() As Long
    Dim g As Integer, i As Integer
    Close 'Close any files
    Call modPCode.LoadPE2(SFilePath)
    Dim F As Integer
    F = FreeFile
    'Get all procedures
    lstProcedures.Clear
    Open SFilePath For Binary Access Read As #F
    For i = 0 To UBound(gObjectInfoHolder)
        If gObjectInfoHolder(i).NumberOfProcs > 0 Then
        ReDim ProcAddr(gObjectInfoHolder(i).NumberOfProcs - 1)
        Seek #F, gObjectInfoHolder(i).aProcTable + 1 - OptHeader.ImageBase
        Get #F, , ProcAddr
        For g = 0 To UBound(ProcAddr)
            If ProcAddr(g) <> 0 And ProcAddr(g) <> -1 Then
                If ProcAddr(g) < UBound(SubName) And ProcAddr(g) > LBound(SubName) Then
                    SubName(ProcAddr(g)) = gObjectNameArray(i) & ".Proc" & ProcAddr(g)
                    lstProcedures.AddItem ProcAddr(g)
                End If
            End If
        Next
        End If
    Next
        Dim addrSubMain As Long
        If gVBHeader.aSubMain <> 0 Then
            Seek #F, gVBHeader.aSubMain + 2 - OptHeader.ImageBase
            Get #F, , addrSubMain
          Dim sTemp
            sTemp = Split(SubName(addrSubMain), ".")
            SubName(addrSubMain) = sTemp(0) & ".Sub Main"
        End If
    Close #F
    'Add Event ProcLists
    For i = 0 To UBound(EventProcList) - 1
        If EventProcList(i) <> 0 Then
            lstProcedures.AddItem EventProcList(i)
        End If
    Next
    For i = 0 To UBound(SubNamelist) - 1
        If SubNamelist(i).offset < UBound(SubName) Then
            SubName(SubNamelist(i).offset) = SubNamelist(i).strName
        End If
    Next

Exit Sub
errHandle:
    MsgBox "Error_frmPcode_form_load: " & err.Number & " " & err.Description
End Sub

Private Sub lstProcedures_Click()
On Error GoTo errHandle
    If lstProcedures.Text <> "" Then
        bEndOfProcedure = False
        If optPCode.value = True Then
            txtView.Text = modPCode.DecompileProc(lstProcedures.Text)
        Else
            txtView.Text = modPCode.DecompileProcToVB(lstProcedures.Text)
        End If
    End If
Exit Sub
errHandle:
MsgBox "Error_frmPCode_lstProcedures_Click: " & err.Number & " " & err.Description
End Sub

Private Sub optPCode_Click()
    lstProcedures_Click
End Sub

Private Sub optPCodeToVB_Click()
    lstProcedures_Click
End Sub

Private Sub txtAddress_Change()
   If IsNumeric(txtAddress.Text) = False Then txtAddress.Text = 0
End Sub


