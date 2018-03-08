VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Visual Basic Obfuscator for  VB 5/6"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6480
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstControlNames 
      Height          =   2205
      Left            =   3480
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
   End
   Begin VB.ListBox lstObjectNames 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblList 
      Caption         =   "Here is a list of names that will be obfuscated."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label lblControlNames 
      Caption         =   "Control and Form names"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblObjectNames 
      Caption         =   "Object Names:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblNote 
      Caption         =   $"frmMain.frx":030A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************
'VB OBfuscator part of Semi VB Decompiler
'Copyright 2005 VisualBasicZone.com
'Developed by Jonathan Valentin
'***************************************
Option Explicit

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileOpen_Click()
    CD.FileName = ""
    CD.DialogTitle = "Select VB 5/6 Program"
    CD.Filter = "VB Files(*.exe,*.ocx,*.dll)|*.exe;*.ocx;*.dll|All Files(*.*)|*.*;"
    CD.ShowOpen
    If CD.FileName = "" Then Exit Sub
    
    lstControlNames.Clear
    lstObjectNames.Clear
    ReDim ObjectSaveInfo(0)
    ReDim ControlSaveInfo(0)
    
    Call OpenVBExe(CD.FileName)
    
End Sub

Private Sub mnuFileSave_Click()
On Error GoTo errHandle
    CD.FileName = ""
    CD.DialogTitle = "Save VB 5/6 Program"
    CD.Filter = "VB Files(*.exe,*.ocx,*.dll)|*.exe;*.ocx;*.dll|All Files(*.*)|*.*;"
    CD.ShowSave
    If CD.FileName = "" Then Exit Sub
    FileCopy SFilePath, App.Path & "\file.exe"
    Dim F As Long
    F = FreeFile
    Open App.Path & "\file.exe" For Binary As #F
        Dim i As Long
        For i = 0 To UBound(ObjectSaveInfo)
            If ObjectSaveInfo(i) <> 0 Then
                Dim sName As String
                Seek #F, ObjectSaveInfo(i)
                sName = GetUntilNull(F)
                Put #F, ObjectSaveInfo(i), MakeString(Len(sName))
                
            End If
        Next i
        For i = 0 To UBound(ControlSaveInfo)
            If ControlSaveInfo(i).Offset <> 0 Then
                Seek #F, ControlSaveInfo(i).Offset
                Dim cControlHeader As ControlHeader
                Dim tArray As ArrayTestType
                Get F, , tArray
                Seek F, ControlSaveInfo(i).Offset
                Dim oldPos As Long
                'cControlHeader.cName = MakeString(ControlSaveInfo(i).Length)
                If tArray.arrayflag <> 128 Then
                    
                    oldPos = Loc(F)
                    Get #F, , cControlHeader
                                        
                    cControlHeader.cName = MakeString(ControlSaveInfo(i).Length)
                    Put #F, oldPos + 1, cControlHeader
            
                Else
                    Dim cArrayHeader As ControlArrayHeader
                    oldPos = Loc(F)
                    Get #F, , cArrayHeader
                    cArrayHeader.cName = MakeString(ControlSaveInfo(i).Length)
                    Put #F, oldPos + 1, cArrayHeader
                   
                    
                End If
            End If
        Next
    Close #F
    FileCopy App.Path & "\file.exe", CD.FileName
    Kill App.Path & "\file.exe"
    MsgBox "File Saved", vbInformation
    Exit Sub
errHandle:
    MsgBox "Error_FileSave: " & Err.Description
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
    
End Sub
Sub OpenVBExe(ByVal FilePath As String)
'################################################
'Purpose: Main function that gets all VB Sturtures
'#################################################
    Dim i As Long 'Loop Var
    Dim addr As Long 'Loop Var
    Dim StartOffset As Long 'Holds Address of first VB Struture
    Dim F As Long
    ReDim gControlNameArray(0) 'Treeveiw control list
    ReDim gControlOffset(0)
    ReDim gProcedureList(0)
    ReDim gObjectNameArray(0)

    Close
    'clear the nodes

    'Save name and path
    SFilePath = FilePath

    
    'Reset the error flag
    ErrorFlag = False
    'Clear the error log

    gVB4App = False
    gVB5App = False
    gVB6App = False

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

    Else
        If VBVersion = 1 Then
            MsgBox "This Program is VB Version 1.0 and is not supported.", vbCritical
            Close #InFileNumber
            Exit Sub
        End If
        
       If VBVersion = 4 Then
            MsgBox "This Program is VB Version 4.0 and is not supported.", vbCritical
            Close #InFileNumber
            Exit Sub
        End If
        
        If VBVersion = 2 Or VBVersion = 3 Then
            MsgBox "This program is VB Version: " & VBVersion & " and is not supported."
            Close #InFileNumber
            Exit Sub
        End If


            MsgBox "Not a VB 5/6 file.", vbOKOnly Or vbCritical Or vbApplicationModal, "Bad file!"
            gVB5App = False
            gVB4App = False
            gVB6App = False
            

        Close #InFileNumber
        Exit Sub
        
    End If
    


        OptHeader.ImageBase = mImageBaseAlign
        StartOffset = VBStartHeader.PushStartAddress - OptHeader.ImageBase


    
    F = FreeFile
    Open SFilePath For Binary As #F
        'Goto begining of vb header
        Seek F, StartOffset + 1
        'Get the vb header
        Get #F, , gVBHeader

        'GetHelpFile
        'Seek #F, StartOffset + 1 + gVBHeader.oHelpFile
        'HelpFile = GetUntilNull(F)
        'Get Project Name
        'Seek #F, StartOffset + 1 + gVBHeader.oProjectName
        'ProjectName = GetUntilNull(F)
        'Project Title
        'Seek #F, StartOffset + 1 + gVBHeader.oProjectTitle
        'ProjectTitle = GetUntilNull(F)
        'ExeName
        'Seek #F, StartOffset + 1 + gVBHeader.oProjectExename
        'ProjectExename = GetUntilNull(F)
        'Get ComRegisterData
        'Seek #F, gVBHeader.aComRegisterData + 1 - OptHeader.ImageBase
        'Get #F, , gCOMRegData
        'Get ProjectDescription
        'Seek #F, gVBHeader.aComRegisterData + 1 + gCOMRegData.oNTSProjectDescription - OptHeader.ImageBase
        'ProjectDescription = GetUntilNull(F)


        
        'Get Project Info Table
        Seek F, gVBHeader.aProjectInfo + 1 - OptHeader.ImageBase
        Get #F, , gProjectInfo


    
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
        ObjectSaveInfo(UBound(ObjectSaveInfo)) = gObject(loopC).aObjectName + 1 - OptHeader.ImageBase
        ReDim Preserve ObjectSaveInfo(UBound(ObjectSaveInfo) + 1)
        gObjectNameArray(loopC) = GetUntilNull(F)
        lstObjectNames.AddItem gObjectNameArray(loopC)
        

        
         'Get Optional Object Info
        Seek F, gObject(loopC).aObjectInfo + 57 - OptHeader.ImageBase
        



        
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
 

                 'Save the control information for the treeview
                 ReDim Preserve gControlNameArray(UBound(gControlNameArray) + 1)
                 gControlNameArray(UBound(gControlNameArray)).strControlName = ControlName
                 gControlNameArray(UBound(gControlNameArray)).strParentForm = gObjectNameArray(loopC)

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


                            ReDim Preserve gProcedureList(UBound(gProcedureList) + 1)
                        End If
                    End If
                Next
         
            End If

        
        End If
        Next loopC

        'Main Loop to Get all Form's Properties
        Call ProccessControls(F)

    Close F

    mnuFileSave.Enabled = True


  Exit Sub
    
AnalyzeError:
    MsgBox "Analyze error " & Err.Description, vbCritical Or vbOKOnly, "Source file error"
    
    Close
End Sub
Sub ProccessControls(F As Long)
'*****************************
'Purpose: Process Forms And Control Properties
'*****************************

    Dim fPos As Long 'Holds current location in the file used for controlheader

    Dim cControlHeader As ControlHeader
    Dim lForm As Long

    Dim posNextControl As Long




            If gVBHeader.FormCount = 0 Then Exit Sub
        
            Seek F, gVBHeader.aGuiTable + 1 - OptHeader.ImageBase
            
            'Get Form table
            If gVBHeader.FormCount > 0 Then
                ReDim gGuiTable(gVBHeader.FormCount - 1)
                'MsgBox Loc(F)
                Get #F, , gGuiTable
                  
            End If


        'Loop though each form...
        For lForm = 0 To UBound(gGuiTable)

            Seek F, gGuiTable(lForm).aFormPointer + 94 - OptHeader.ImageBase
   
        


       
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
           
            cControlHeader.Length = cArrayHeader.Length
            cControlHeader.cName = cArrayHeader.cName
            cControlHeader.cType = cArrayHeader.cType
            cControlHeader.cId = cArrayHeader.cId
            
        End If
            ControlSaveInfo(UBound(ControlSaveInfo)).Length = Len(cControlHeader.cName)
            ControlSaveInfo(UBound(ControlSaveInfo)).Offset = fPos + 1
            ReDim Preserve ControlSaveInfo(UBound(ControlSaveInfo) + 1)
            lstControlNames.AddItem cControlHeader.cName
        posNextControl = fPos + cControlHeader.Length + 2
    

       
       

      

     
       ' MsgBox cControlHeader.cName
        Seek F, fPos + cControlHeader.Length
        GoTo EndLabel
        
EndLabel:
        'Get the seperator type for the end of the control
        Dim cControlEnd As Integer
        Dim bCheckEnd As Byte
        
        Get #F, , cControlEnd ' = GetInteger(F)

        If cControlEnd = vbFormNewChildControl Then
            Seek F, posNextControl
            GoTo NewControl
            
        End If
        If cControlEnd = vbFormChildControl Then 'FF03
            If cControlHeader.cType <> vbmenu Then

                Do
                    Get F, , bCheckEnd
  
                Loop Until bCheckEnd > 3 Or bCheckEnd = 0
                 
                If bCheckEnd <> 4 Then
                    'MsgBox cControlHeader.cName & " " & cControlEnd & " Pos: " & Loc(f) & "next: " & posNextControl
                    Seek F, Loc(F)
                    GoTo NewControl
                End If
            Else


                Do
                    Get F, , bCheckEnd

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

                 Do
                     Get F, , bCheckEnd

                 Loop Until bCheckEnd >= 3 Or bCheckEnd = 0
                     If bCheckEnd = 0 Or bCheckEnd > 5 Then
                        Seek F, Loc(F)
                     End If
                     If bCheckEnd <> 4 Then
                         GoTo NewControl
                     End If
            Else

                GoTo NewControl
            End If
            
        End If
        If cControlEnd = vbFormMenu Then
          'Seek f, posNextControl

            GoTo NewControl
            
        End If


NextFormDec:
    
    Next lForm 'Main Form Loop

End Sub
Public Function MakeString(Length As Long) As String
    Dim strDATA As String
    Dim i As Long
    If Length > 1 Then
        strDATA = Chr(97 + Int(Rnd * 25))
        For i = 2 To Length
            If Int(Rnd * 2) = 1 Then
                strDATA = strDATA & Chr(48 + Int(Rnd * 9))
            Else
                strDATA = strDATA & Chr(97 + Int(Rnd * 25))
            End If
        Next
    Else
        strDATA = Chr(97 + Int(Rnd * 25))
    End If
    
    MakeString = strDATA
End Function

