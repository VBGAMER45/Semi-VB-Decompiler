VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmNativeDecompile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Native Procedure Decompile (Beta)"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExportList 
      Caption         =   "E&xport List"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      ToolTipText     =   "Export the Native Procedure address list."
      Top             =   3960
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   840
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Item"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   3120
      TabIndex        =   8
      Text            =   "0"
      ToolTipText     =   "Needs to be a memory address. If you have a file offset use the File Offset to Address convertor."
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdAddAddress 
      Caption         =   "&Add Address"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      ToolTipText     =   "Adds an address to the list to be decompiled."
      Top             =   3960
      Width           =   1575
   End
   Begin VB.ListBox lstProcedures 
      Height          =   3570
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtView 
      Height          =   3570
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.OptionButton optPCode 
      Caption         =   "Native Asm"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Decompiles a procedure and shows the P-Code Tokens"
      Top             =   4560
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optNativeToVB 
      Caption         =   "Native Asm To VB"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      ToolTipText     =   "Decompiles the selected procedure and attempts to conver the P-Code Tokens to VB Code"
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblSetting 
      Caption         =   "Current Setting:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
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
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Just click on a procedure in the list to disassemble it."
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmNativeDecompile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************
'frmNativeDecompile
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************
Option Explicit
Dim dsm As New CDisassembler
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
        Print #F, "Semi VB Decompiler - Native Procedure List"
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
    Me.lstProcedures.Clear
    If gProjectInfo.aNativeCode <> 0 Then
        If gVBHeader.aSubMain <> 0 Then
            Me.lstProcedures.AddItem gVBHeader.aSubMain
        End If
    End If
    Dim i As Integer
    For i = 0 To UBound(gNativeProcArray) - 1
        Me.lstProcedures.AddItem gNativeProcArray(i).offset
    Next i

    
    'Me.lstProcedures.AddItem "4247518"
    'Me.lstProcedures.AddItem gProjectInfo.aNativeCode
End Sub

Private Sub lstProcedures_Click()
On Error GoTo errHandle
    If lstProcedures.ListIndex = -1 Then Exit Sub

    Dim fp As Integer, g As Long
     
    txtView.Text = ""

    If lstProcedures.List(lstProcedures.ListIndex) = gVBHeader.aSubMain Then
        txtView.Text = txtView.Text & "Disassembly of SubMain()" & vbCrLf
    Else
        For g = 0 To UBound(gNativeProcArray)
            If gNativeProcArray(g).offset = lstProcedures.List(lstProcedures.ListIndex) Then
                txtView.Text = txtView.Text & "Disassembly of " & gNativeProcArray(g).sName & "()" & vbCrLf
            End If
        Next g
    End If
    Dim b(5000) As Byte
    fp = FreeFile
    'MsgBox "ag"
    'Close
    Open SFilePath For Binary Access Read As #fp
       Get #fp, lstProcedures.List(lstProcedures.ListIndex) + 1 - OptHeader.ImageBase, b
    Dim va As Long
    Dim col As Collection
    Dim inst As CInstruction
     va = lstProcedures.List(lstProcedures.ListIndex)
    
    Set col = dsm.DisasmBlock(b(), va)
    Dim strBuffer As String
    strBuffer = txtView.Text
    If NativeShowOffsets = True And NativeShowHexInformation = True Then
        For Each inst In col
           'Set li = lvDisasm.ListItems.Add(, , inst.offset)
           'li.SubItems(1) = inst.dump
           'li.SubItems(2) = inst.command
           
           strBuffer = strBuffer & inst.offset & " " & inst.dump & " " & inst.command & vbCrLf
            
            If inst.command = "RETN" Then
                Exit For
            End If
        Next
        
    ElseIf NativeShowOffsets = True And NativeShowHexInformation = False Then
        For Each inst In col
           'Set li = lvDisasm.ListItems.Add(, , inst.offset)
           'li.SubItems(1) = inst.dump
           'li.SubItems(2) = inst.command
           
           strBuffer = strBuffer & inst.offset & " " & inst.command & vbCrLf
            
            If inst.command = "RETN" Then
                Exit For
            End If
        Next
    ElseIf NativeShowOffsets = False And NativeShowHexInformation = True Then
        For Each inst In col
           'Set li = lvDisasm.ListItems.Add(, , inst.offset)
           'li.SubItems(1) = inst.dump
           'li.SubItems(2) = inst.command
           
           strBuffer = strBuffer & inst.dump & " " & inst.command & vbCrLf
            
            If inst.command = "RETN" Then
                Exit For
            End If
        Next
    Else
        For Each inst In col
           'Set li = lvDisasm.ListItems.Add(, , inst.offset)
           'li.SubItems(1) = inst.dump
           'li.SubItems(2) = inst.command
           
           strBuffer = strBuffer & inst.command & vbCrLf
            
            If inst.command = "RETN" Then
                Exit For
            End If
        Next
    
    End If
    txtView.Text = strBuffer
    
    Close #fp
Exit Sub
errHandle:
MsgBox "Error_frmNativeDecompile_lstProcedures_Click: " & err.Number & " " & err.Description
End Sub

Private Sub optNativeToVB_Click()
    MsgBox "Cheater :)"
End Sub

Private Sub txtAddress_Change()
    If IsNumeric(txtAddress.Text) = False Then txtAddress.Text = 0
End Sub




