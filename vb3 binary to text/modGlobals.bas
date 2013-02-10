Attribute VB_Name = "modGlobals"
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'***********************************
Option Explicit

Public Const Version As String = "0.01"  'Version Number

Public Enum ControlType
    vbPictureBox = 0
    vbLabel = 1
    vbTextBox = 2
    vbFrame = 3
    vbCommandbutton = 4
    vbCheckbox = 5
    vbOptionbutton = 6
    vbComboBox = 7
    vbListbox = 8
    vbHscroll = 9
    vbVscroll = 10
    vbTimer = 11
    vbForm = 13
    vbDriveListbox = 16
    vbDirectoryListbox = 17
    vbFileListbox = 18
    vbMenu = 19
    vbMDIForm = 20
    vbShape = 22
    vbLine = 23
    vbImage = 24
    vbData = 37
    vbOLE = 38
    vbUserControl = 40
    vbPropertyPage = 41
    vbUserDocument = 42
End Enum

'Defualt Control Header for non control arrays
Public Type ControlHeader
    Length As Integer
    B1 As Byte
    IsArray As Byte
    ControlID As Byte
    NameLength As Byte
End Type

'Array Header Type
Public Type ArrayControlHeader
    Length As Integer
    B1 As Byte
    IsArray As Integer
    ControlID As Byte
    un1 As Byte
    NameLength As Byte
End Type

'The form header
Public Type FormHeader
    i1 As Integer
    un1 As Integer
    NumberOfControls As Byte 'not including form
    un2 As Byte
    un3 As Integer
    un4 As Byte
End Type

'Control Sepeartor Constatns
Public Const vbFormNewChildControl = 511 'FF01
Public Const vbFormExistingChildControl = 767 'FF02
Public Const vbFormChildControl = 1023 'FF03
Public Const vbFormEnd = 1279 'FF04
Public Const vbFormMenu = 1535 'FF05

'Show the Gui Offsets?
Global gShowOffsets As Boolean
'Extract Images?
Global gExtractImages As Boolean
'Process VBX Controls
Global gProcessVBX As Boolean
'How many spaces to ident?
Global gIdentSpaces As Integer

Global strBuffer As String 'Holds the form data
Global gFormDone As Boolean
Global gVBFormFile As clsFile

'Hold the paths to each VBX Control
Global gVBXControlPath() As String

Global bFirstFF As Boolean
Public Sub AddError(ByVal strText As String)
'**********************************
'Purpose: Adds an error to the ErrorLog
'**********************************
    frmMain.txtErrorLog.Text = frmMain.txtErrorLog.Text & strText & vbCrLf
End Sub
Public Sub LoadNewFormHolder(ByVal FormName As String)
'*****************************
'Purpose:To load a new textbox to hold each form's information
'*****************************
    Dim i As Integer
    
    i = frmMain.txtStorage.UBound + 1
    Load frmMain.txtStorage(i)
    With frmMain.txtStorage(i)
        .Tag = FormName
        
    End With
    frmMain.txtStorage(i - 1).Text = strBuffer
    strBuffer = ""
End Sub
Sub DoFinalFormBuffer()
'**********************************
'Purpose: Stores the final buffer into the txtarray then the memory buffer is cleared
'**********************************
    Dim i As Integer
    i = frmMain.txtStorage.UBound
    frmMain.txtStorage(i).Text = strBuffer
    strBuffer = ""
End Sub
Public Sub GetControlSize(ByVal F As Integer)
'**********************************
'Purpose: Gets control size property for controls.
'**********************************
    Dim cTop As Long, cLeft As Long, cHeight As Long, cWidth As Long
    cLeft = GetWordByFile(F)
    cTop = GetWordByFile(F)
    cWidth = GetWordByFile(F)
    cHeight = GetWordByFile(F)
    
    Call frmMain.AddText("Left = " & cLeft)
    Call frmMain.AddText("Top = " & cTop)
    Call frmMain.AddText("Width = " & cWidth)
    Call frmMain.AddText("Height = " & cHeight)
End Sub
Public Function GetTrueFalse(ByVal F As Integer) As Integer
'**********************************
'Purpose: Gets True False from file.
'**********************************
    Dim b As Byte
    Get #F, , b
    If b = 0 Then
        GetTrueFalse = -1
    Else
        GetTrueFalse = 0
    End If
End Function
Public Sub PrintReadMe()
'**********************************
'Purpose: Prints a ReadMe each time the program is run, just in case it is not included
'**********************************
On Error GoTo nofile:
    Dim F As Integer
    F = FreeFile
    Open App.Path & "\readmebinarytotext.txt" For Output As #F
        Print #F, "************************************"
        Print #F, "VB 1/2/3 Binary Form to Text Converter by vbgamer45"
        Print #F, "Version : " & App.Major & "." & App.Minor & "." & App.Revision
        Print #F, "http://www.visualbasiczone.com"
        Print #F, "************************************"
        Print #F, ""
        Print #F, "Contents"
        Print #F, " 1. Features"
        Print #F, " 2. Bugs"
        Print #F, " 3. Contact"
        Print #F, ""
        Print #F, "1. Features"
        Print #F, "   Converts VB 1/2/3 Binary forms to text."
        Print #F, "   You can open a whole project and convert each form to text format."
        Print #F, "   Does not require VB 3/4 to be installed!"
        Print #F, ""
        Print #F, "2. Bugs"
        Print #F, "   If you find a bug please email me and give as much information as possible."
        Print #F, ""
        Print #F, "3. Contact"
        Print #F, "   You can reach me at support@visualbasiczone.com"
    Close #F
Exit Sub
nofile:
    Exit Sub
End Sub
Public Function PadHex(ByVal sHex As String, Optional Pad As Integer = 8) As String
'*****************************
'Purpose: To add extra zero's to a hexadecimal string
'*****************************
    If Len(sHex) > Pad Then
        PadHex = sHex
    Else
        PadHex = String(Pad - Len(sHex), 48) & sHex
    End If
End Function
Public Sub GetFontProperty(ByVal F As Integer)
'*****************************
'Purpose: Gets the Font property from a file
'*****************************
        Dim bData As Byte
        Get #F, , bData
        Call frmMain.AddText("FontName = " & Chr$(34) & gVBFormFile.GetString(Loc(F), bData, False) & Chr$(34))
        Call frmMain.AddText("FontSize = " & gVBFormFile.GetSingle(Loc(F)))
        Get #F, , bData
        If bData And 1 Then
            Call frmMain.AddText("FontBold = -1")
        Else
            Call frmMain.AddText("FontBold = 0")
        End If
        If bData And 2 Then
            Call frmMain.AddText("FontItalic = -1")
        Else
            Call frmMain.AddText("FontItalic = 0")
        End If
        If bData And 8 Then
            Call frmMain.AddText("FontStrikethru = -1")
        Else
            Call frmMain.AddText("FontStrikethru = 0")
        End If
        If bData And 4 Then
            Call frmMain.AddText("FontUnderline = -1")
        Else
            Call frmMain.AddText("FontUnderline = 0")
        End If
End Sub


