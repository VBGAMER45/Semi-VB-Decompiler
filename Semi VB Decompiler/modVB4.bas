Attribute VB_Name = "modVB4"
'*********************************************
'modVB4
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************
Option Explicit

Private Type oldVB4HeaderType
    b1 As Byte '68 JmpOpcode
    lAddress1 As Long
    b2 As Byte 'E8 Jmp Opcode
    ThunderRTMain As Long 'JMP VB40032 #100
    b3 As Byte 'NULL
    b4 As Byte 'NULL
    LangID As Integer
    b5 As Byte 'NULL
    b6 As Byte
    b7 As Byte 'Number
    b8 As Byte 'NULL
    b9 As Byte 'NULL
    b10 As Byte 'NULL
    b11 As Byte 'NULL
    b12 As Byte 'NULL
    FormCount As Byte
    b13 As Byte
    lAddress2 As Long 'Sometimes Null
    lAddress3 As Long
    lAddress4 As Long 'ThunProject1
    lAddress5 As Long
    bA(19) As Byte
    lAddress6 As Long 'End of Code?
End Type

Private Type VB4HEADERType
    sig As Long 'SIG 129 53 84 182
    CompilerFileVersion As Integer
    int1 As Integer
    int2 As Integer
    int3 As Integer
    int4 As Integer
    int5 As Integer
    int6 As Integer
    int7 As Integer
    int8 As Integer
    int9 As Integer
    int10 As Integer
    int11 As Integer
    int12 As Integer
    int13 As Integer
    int14 As Integer
    int15 As Integer
    LangID As Integer
    int16 As Integer
    int17 As Integer
    int18 As Integer
    aSubMain As Long
    Address2 As Long
    i1 As Integer
    i2 As Integer
    i3 As Integer
    i4 As Integer
    i5 As Integer
    i6 As Integer
    iExeNameLength As Integer
    iProjectSavedNameLength As Integer
    iHelpFileLength As Integer
    iProjectNameLength As Integer
    FormCount As Integer
    int19 As Integer
    NumberOfExternalComponets As Integer
    int20 As Integer  'The same in each file 176d
    aGuiTable As Long  'GUI Pointer
    Address4 As Long
    aExternalComponetTable As Long '??Not a 100% sure
    aProjectInfo2 As Long  'Project Info2?
    
End Type

Private Type VB4GuiTableType
    uuid(15) As Byte
    bArray(27) As Byte
    aFormPointer As Long
End Type
Private Type ProjectInfo
    LangID As Integer
    i1 As Integer
    i2 As Integer
    i3 As Integer
    i4 As Integer
    i5 As Integer
    Address1 As Long 'Points to address
    Address2 As Long 'Points to ProjectInfoAddress2Link
    Address3 As Long
    Address4 As Long
End Type

Private Type ProjectInfoAddress2Link
    Address1 As Long
    Address2 As Long
    Address3 As Long
    Adddres4 As Long
End Type
Private Type ProjectInfo2Type
    l1 As Long
    l2 As Long
    l3 As Long
    oProjectName As Long 'Project Name
    oVBPath As Long 'VB3 Path
    oAppDescription As Long 'APP Description
    guiduuid(15) As Byte
End Type

Global VB4Header As VB4HEADERType
Dim VB4GuiTables() As VB4GuiTableType
Dim vb4Projectinfo2 As ProjectInfo2Type
Global strVB4Forms() As String
Sub OpenVB4EXE(strFileName As String, lOffset As Long)
    Dim F As Integer
    Dim b As Byte
    Dim lHeaderOffset As Long
    Dim i As Integer
    Dim strBuffer As String
    Dim bDebugBuild As Boolean
    bDebugBuild = False
    F = FreeFile
    ReDim strVB4Forms(0)
    Open strFileName For Binary Access Read As #F
        Get F, lOffset + 1, b
        Get F, , lHeaderOffset
        
        Set gVBFile = New clsFile
        Call gVBFile.Setup(strFileName)
        F = gVBFile.FileNumber
       ' MsgBox "OFFSET!:" & lHeaderOffset - OptHeader.ImageBase + 1
        Seek F, lHeaderOffset - OptHeader.ImageBase + 1

        
        Get F, , VB4Header
        'Get VB EXE Name
        If VB4Header.iExeNameLength <> 0 Then
            ProjectExename = modGlobals.GetUntilNull(F)
        End If
        'Get Saved Project Name
        If VB4Header.LangID <> 0 Then
            ProjectName = modGlobals.GetUntilNull(F)
        End If
        'Get Help File
        If VB4Header.iHelpFileLength <> 0 Then
            HelpFile = modGlobals.GetUntilNull(F)
        End If
        'Get Project Name
        If VB4Header.iProjectNameLength <> 0 Then
            ProjectTitle = modGlobals.GetUntilNull(F)
        End If
        
        If VB4Header.FormCount > 0 Then
            ReDim VB4GuiTables(VB4Header.FormCount - 1)
            ReDim gGuiTable(VB4Header.FormCount - 1)
            Seek F, VB4Header.aGuiTable - OptHeader.ImageBase + 1
            Get F, , VB4GuiTables
            For i = 0 To UBound(gGuiTable)
                gGuiTable(i).aFormPointer = VB4GuiTables(i).aFormPointer
            Next
        End If
        'MsgBox "OFFSET!:" & VB4Header.aProjectInfo2 - OptHeader.ImageBase + 1
        Seek F, VB4Header.aProjectInfo2 - OptHeader.ImageBase + 1
        Get F, , vb4Projectinfo2
        If vb4Projectinfo2.oProjectName <> 0 Then
            Seek F, VB4Header.aProjectInfo2 + vb4Projectinfo2.oProjectName - OptHeader.ImageBase + 1
            strBuffer = modGlobals.GetUntilNull(F)
        End If
        If vb4Projectinfo2.oVBPath <> 0 Then
            Seek F, VB4Header.aProjectInfo2 + vb4Projectinfo2.oVBPath - OptHeader.ImageBase + 1
            strBuffer = modGlobals.GetUntilNull(F)
        End If
        If vb4Projectinfo2.oAppDescription <> 0 Then
            Seek F, VB4Header.aProjectInfo2 + vb4Projectinfo2.oAppDescription - OptHeader.ImageBase + 1
            ProjectDescription = modGlobals.GetUntilNull(F)
        End If


        'Main Loop to Get all Form's Properties
        frmMain.FrameStatus.Visible = True
        frmMain.txtStatus.Text = vbNullString
        Call frmMain.ProccessControls(F)
        Call modGlobals.DoFinalFormBuffer
    Close F


End Sub

