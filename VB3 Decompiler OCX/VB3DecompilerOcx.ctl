VERSION 5.00
Begin VB.UserControl VB3DecompilerOcx 
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1785
   InvisibleAtRuntime=   -1  'True
   Picture         =   "VB3DecompilerOcx.ctx":0000
   ScaleHeight     =   915
   ScaleWidth      =   1785
   Begin VB.Label lblUrl 
      BackStyle       =   0  'Transparent
      Caption         =   "VisualBasicZone.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "VB3DecompilerOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
'Const m_def_About = 0
Const m_def_SetDataPath = ""
Const m_def_Filename = ""
Const m_def_OutputFolder = ""
'Property Variables:

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get SetDataPath() As String
Attribute SetDataPath.VB_Description = "Sets the data path for the VB3 decompiler files."
    SetDataPath = m_SetDataPath
End Property

Public Property Let SetDataPath(ByVal New_SetDataPath As String)
    m_SetDataPath = New_SetDataPath
    PropertyChanged "SetDataPath"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Filename() As String
Attribute Filename.VB_Description = "Sets the Filename of the VB3 file to be decompiled."
    Filename = m_Filename
End Property

Public Property Let Filename(ByVal New_Filename As String)
    m_Filename = New_Filename
    PropertyChanged "Filename"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get OutputFolder() As String
Attribute OutputFolder.VB_Description = "Sets the Output folder path."
    OutputFolder = m_OutputFolder
End Property

Public Property Let OutputFolder(ByVal New_OutputFolder As String)
    m_OutputFolder = New_OutputFolder
    PropertyChanged "OutputFolder"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_SetDataPath = m_def_SetDataPath
    m_Filename = m_def_Filename
    m_OutputFolder = m_def_OutputFolder
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_SetDataPath = PropBag.ReadProperty("SetDataPath", m_def_SetDataPath)
    m_Filename = PropBag.ReadProperty("Filename", m_def_Filename)
    m_OutputFolder = PropBag.ReadProperty("OutputFolder", m_def_OutputFolder)
End Sub

Private Sub UserControl_Resize()
    Width = 1785
    Height = 915
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("SetDataPath", m_SetDataPath, m_def_SetDataPath)
    Call PropBag.WriteProperty("Filename", m_Filename, m_def_Filename)
    Call PropBag.WriteProperty("OutputFolder", m_OutputFolder, m_def_OutputFolder)
'    Call PropBag.WriteProperty("About", m_About, m_def_About)
End Sub
Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_Description = "Shows the About Box"
Attribute ShowAboutBox.VB_UserMemId = -552
Attribute ShowAboutBox.VB_MemberFlags = "40"
    frmAbout.Show vbModal, Me
End Sub
Public Function DecompileFile() As String
Attribute DecompileFile.VB_Description = "Decompiles a VB2 or VB3 file as long as the folder, data, and filename are set"
    If m_Filename = "" Then
        MsgBox "No file specified to decompile!", vbCritical
        Exit Function
    End If
    If m_OutputFolder = "" Then
        MsgBox "No output folder set!", vbCritical
        Exit Function
    End If
    If m_SetDataPath = "" Then
        MsgBox "No data path set!", vbCritical
        Exit Function
    End If
    frmMain.Hide
    frmMain.Show
    frmMain.Hide
    Call frmMain.DoDecompile
    
End Function
