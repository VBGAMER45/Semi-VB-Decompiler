Attribute VB_Name = "modOutput"
'*********************************************
'modOutput
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************
Option Explicit
'Types and Variables for exe changes
Private Type byteChangeType
    bByte As Byte
    offset As Long
End Type
Private Type BooleanChangeType
    bBool As Boolean
    offset As Long
End Type
Private Type IntegerChangeType
    iInt As Integer
    offset As Long
End Type
Private Type LongChangeType
    lLong As Long
    offset As Long
End Type
Private Type SingleChangeType
    sSingle As Single
    offset As Long
End Type
Private Type StringChangeType
    sString As String
    offset As Long
End Type
Private Type PictureChangeType
    length As Long
    offset As Long
End Type

Global PictureChange() As PictureChangeType
Global ByteChange() As byteChangeType
Global BooleanChange() As BooleanChangeType
Global IntegerChange() As IntegerChangeType
Global LongChange() As LongChangeType
Global SingleChange() As SingleChangeType
Global StringChange() As StringChangeType

'Holds Control offsets.
Private Type ControlOffsetType
    offset As Long
    Owner As String
    ControlName As String
    ControlType As Byte
End Type
Global gControlOffset() As ControlOffsetType
'Strutures holding external objects and form name they are connected with.
Private Type ExternalObjectType
    strFormName As String
    strLibname As String
End Type
Global gExternalObjectHolder() As ExternalObjectType
Sub DumpVBExeInfo(ByVal strFileName As String, ByVal FileTitle As String)
'*****************************
'Purpose: Prints a report about the Exe that was decompiled
'*****************************
Dim i As Long
Close 'in case files are open
    Open strFileName For Output As #1
        Print #1, "----------------------------------"
        Print #1, FileTitle
        Print #1, "Output made by Semi VB Decompiler by VisualBasicZone.com"
        Print #1, "----------------------------------"
        Print #1, "EXE Header / MS-DOS stub"
        Print #1, "----------------------------------"
        Print #1, "Magic= " & DosHeader.Magic
        Print #1, "NumBytesLastPage= " & DosHeader.NumBytesLastPage
        Print #1, "NumPages= " & DosHeader.NumPages
        Print #1, "NumBytesLastPage= " & DosHeader.NumBytesLastPage
        Print #1, "NumRelocates= " & DosHeader.NumRelocates
        Print #1, "NumHeaderBlks= " & DosHeader.NumHeaderBlks
        Print #1, "NumMinBlks= " & DosHeader.NumMinBlks
        Print #1, "NumMaxBlks= " & DosHeader.NumMaxBlks
        Print #1, "SSPointer= " & DosHeader.SSPointer
        Print #1, "SPPointer= " & DosHeader.SPPointer
        Print #1, "Checksum= " & DosHeader.Checksum
        Print #1, "IPPointer= " & DosHeader.IPPointer
        Print #1, "CurrentSeg= " & DosHeader.CurrentSeg
        Print #1, "RelocTablePointer= " & DosHeader.RelocTablePointer
        Print #1, "Overlay= " & DosHeader.Overlay
        Print #1, "ReservedW1= " & DosHeader.ReservedW1
        Print #1, "ReservedW2= " & DosHeader.ReservedW2
        Print #1, "ReservedW3= " & DosHeader.ReservedW3
        Print #1, "ReservedW4= " & DosHeader.ReservedW4
        Print #1, "OEMType= " & DosHeader.OEMType
        Print #1, "OEMData= " & DosHeader.OEMData
        Print #1, "ReservedW5= " & DosHeader.ReservedW5
        Print #1, "ReservedW6= " & DosHeader.ReservedW6
        Print #1, "ReservedW7= " & DosHeader.ReservedW7
        Print #1, "ReservedW8= " & DosHeader.ReservedW8
        Print #1, "ReservedW9= " & DosHeader.ReservedW9
        Print #1, "ReservedW10= " & DosHeader.ReservedW10
        Print #1, "ReservedW11= " & DosHeader.ReservedW11
        Print #1, "ReservedW12= " & DosHeader.ReservedW12
        Print #1, "ReservedW13= " & DosHeader.ReservedW13
        Print #1, "ReservedW14= " & DosHeader.ReservedW14
        Print #1, "ExeHeaderPointer= " & DosHeader.ExeHeaderPointer
        
        If bNEFormat = False Then
            Print #1, "----------------------------------"
            Print #1, "Coff Header"
            Print #1, "----------------------------------"
            Print #1, "Magic= " & PEHeader.Magic
            Print #1, "Machine= " & PEHeader.Machine
            Print #1, "NumberOfSections= " & PEHeader.NumSections
            Print #1, "TimeDateStamp= " & PEHeader.TimeDate
            Print #1, "PointerToSymbolTable= " & PEHeader.SymbolTablePointer
            Print #1, "NumberOfSymbols= " & PEHeader.NumSymbols
            Print #1, "SizeOfOptionalHeader= " & PEHeader.OptionalHdrSize
            Print #1, "Characteristics= " & PEHeader.Properties
            Print #1, "----------------------------------"
            Print #1, "Optional Header"
            Print #1, "----------------------------------"
            Print #1, "Magic= " & OptHeader.Magic
            Print #1, "MajorLinkerVersion= " & OptHeader.MajLinkerVer
            Print #1, "MinorLinkerVersion= " & OptHeader.MinLinkerVer
            Print #1, "SizeOfCode= " & OptHeader.CodeSize
            Print #1, "SizeOfInitializedData= " & OptHeader.InitDataSize
            Print #1, "SizeOfUninitializedData= " & OptHeader.UninitDataSize
            Print #1, "AddressOfEntryPoint= " & OptHeader.EntryPoint
            Print #1, "BaseOfCode= " & OptHeader.CodeBase
            Print #1, "BaseOfData= " & OptHeader.DataBase
            Print #1, "ImageBase= " & OptHeader.ImageBase
            Print #1, "SectionAlignment= " & OptHeader.SectionAlignment
            Print #1, "FileAlignment= " & OptHeader.FileAlignment
            Print #1, "MajorOperatingSystemVersion= " & OptHeader.MajOSVer
            Print #1, "MinorOperatingSystemVersion= " & OptHeader.MinOSVer
            Print #1, "MajorImageVersion= " & OptHeader.MajImageVer
            Print #1, "MinorImageVersion= " & OptHeader.MinImageVer
            Print #1, "MajorSubsystemVersion= " & OptHeader.MajSSysVer
            Print #1, "MinorSubsystemVersion= " & OptHeader.MinSSysVer
            Print #1, "Win32VersionValue= " & OptHeader.Win32Ver
            Print #1, "SizeOfImage= " & OptHeader.SizeImage
            Print #1, "SizeOfHeaders= " & OptHeader.SizeHeader
            Print #1, "CheckSum= " & OptHeader.Checksum
            Print #1, "Subsystem= " & OptHeader.SSystem
            Print #1, "DllCharacteristics= " & OptHeader.DLLProperties
            Print #1, "SizeOfStackReserve= " & OptHeader.SSizeRes
            Print #1, "SizeOfStackCommit= " & OptHeader.SSizeCom
            Print #1, "SizeOfHeapReserve= " & OptHeader.HSizeRes
            Print #1, "SizeOfHeapCommit= " & OptHeader.HSizeCom
            Print #1, "LoaderFlags= " & OptHeader.LFlags
            Print #1, "NumberOfRvaAndSizes= " & OptHeader.NumRVA_Sizes
            Print #1, "----------------------------------"
            Print #1, "Section Headers"
            Print #1, "----------------------------------"
            For i = 0 To PEHeader.NumSections - 1
                Print #1, "Name= " & SecHeader(i).SecName
                Print #1, "Misc= " & SecHeader(i).Misc
                Print #1, "VirtualAddress= " & SecHeader(i).Address
                Print #1, "SizeOfRawData= " & SecHeader(i).SizeRawData
                Print #1, "PointerToRawData= " & SecHeader(i).RawDataPointer
                Print #1, "PointerToRelocations= " & SecHeader(i).RelocationPointer
                Print #1, "PointerToLinenumbers= " & SecHeader(i).LineNumPointer
                Print #1, "NumberOfRelocations= " & SecHeader(i).NumRelocations
                Print #1, "NumberOfLinenumbers= " & SecHeader(i).NumLineNumbers
                Print #1, "Characteristics= " & SecHeader(i).Properties
                
            Next
        Else
        'Is NE format for VB Versions 1,2,3, and 4 16bit
            Print #1, "----------------------------------"
            Print #1, "SEGMENTED EXE HEADER"
            Print #1, "----------------------------------"
            Print #1, "Signature= " & NEHeader.signature
            Print #1, "VersionNumberofLinker= " & NEHeader.VersionLinker
            Print #1, "RevisionNumberofLinker= " & NEHeader.RevisionLinker
            Print #1, "EntryTableFileOffset= " & NEHeader.EntryTableOffset
            Print #1, "SizeOfEntryTable= " & NEHeader.SizeOfEntryTable
            Print #1, "CRC= " & NEHeader.CRC
            Print #1, "Flags= " & NEHeader.flags
            Print #1, "SegmentNumberAutomaticDataSegment= " & NEHeader.SegmentNumberAutomaticDataSegment
            Print #1, "InitialHeapSize= " & NEHeader.InitialSizeHeap
            Print #1, "InitialStackSize= " & NEHeader.InitialSizeStack
            Print #1, "OffsetOfCS= " & NEHeader.SegmentNumberOffsetCS
            Print #1, "OffsetOfSS= " & NEHeader.SegmentNumberOffsetSS
            Print #1, "NumberOfEntriesSegmentTable= " & NEHeader.NumberEntriesSegmentTable
            Print #1, "NumberOfEntriesModuleReferenceTable= " & NEHeader.NumberEntriesModuleReferenceTable
            Print #1, "SizeOfNon-ResidentNameTable= " & NEHeader.SizeOfNonResidentNameTable
            Print #1, "SegmentTableFileOffset= " & NEHeader.SegmentTableOffset
            Print #1, "ResourceTableFileOffset= " & NEHeader.ResourceTableFileOffset
            Print #1, "ResidentNameTableFileOffset= " & NEHeader.ResidentNameTableOffset
            Print #1, "ModuleReferenceTableFileOffset= " & NEHeader.ModuleReferenceTableOffset
            Print #1, "ImportedNamesTableFileOffset= " & NEHeader.ImportedNamesTableOffset
            Print #1, "NonResidentNameTableOffset= " & NEHeader.NonResidentNameTableOffset
            Print #1, "NumberMovableEntriesEntryTable= " & NEHeader.NumberMovableEntriesInEntryTable
            Print #1, "LogicalSectorAlignmentShiftCount= " & NEHeader.LogicalSectorAlignmentShiftCount
            Print #1, "NumberResourceEntries= " & NEHeader.NumberResourceEntries
            Print #1, "ExecutableType= " & NEHeader.ExecutableType
            
        End If
        
        If bISVBNET = True Then
            'Print Clr information
            Print #1, "----------------------------------"
            Print #1, "Common Language Runtime Header"
            Print #1, "----------------------------------"
            Print #1, "CB= " & gVBNETHeader.CB
            Print #1, "MajorRuntimeVersion= " & gVBNETHeader.MajorRuntimeVersion
            Print #1, "MinorRuntimeVersion= " & gVBNETHeader.MinorRuntimeVersion
            Print #1, "MetaDataSize= " & gVBNETHeader.MetaData.Size
            Print #1, "MetaDataVirtualAddress= " & gVBNETHeader.MetaData.VirtualAddress
            Print #1, "Flags= " & gVBNETHeader.flags
            Print #1, "EntryPointToken= " & gVBNETHeader.EntryPointToken
            Print #1, "ResourcesSize= " & gVBNETHeader.Resources.Size
            Print #1, "ResourcesVirtualAddress= " & gVBNETHeader.Resources.VirtualAddress
            Print #1, "StrongNameSignatureSize= " & gVBNETHeader.StrongNameSignature.Size
            Print #1, "StrongNameSignatureVirtualAddress= " & gVBNETHeader.StrongNameSignature.VirtualAddress
            Print #1, "CodeManagerTableSize= " & gVBNETHeader.CodeManagerTable.Size
            Print #1, "CodeManagerTableVirtualAddress= " & gVBNETHeader.CodeManagerTable.VirtualAddress
            Print #1, "VTableFixupsSize= " & gVBNETHeader.VTableFixups.Size
            Print #1, "VTableFixupsVirtualAddress= " & gVBNETHeader.VTableFixups.VirtualAddress
            Print #1, "ExportAddressTableJumpsSize= " & gVBNETHeader.ExportAddressTableJumps.Size
            Print #1, "ExportAddressTableJumpsVirtualAddress= " & gVBNETHeader.ExportAddressTableJumps.VirtualAddress
            Print #1, "ManagedNativeHeaderSize= " & gVBNETHeader.ManagedNativeHeader.Size
            Print #1, "ManagedNativeHeaderVirtualAddress= " & gVBNETHeader.ManagedNativeHeader.VirtualAddress
            
        End If
    
        
        If VBVersion <> 0 Then
            Print #1, "----------------------------------"
            Print #1, "VB Exe Info"
            Print #1, "----------------------------------"
            Print #1, "VB Version: " & VBVersion
        End If
        'Don't print VB information if not VB exe
        If gVB6App = True Or gVB5App = True Then
        Print #1, "VBStartOffset= " & AppData.VBStartOffset
        Print #1, "FormCount= " & gVBHeader.FormCount
        Print #1, "ModuleCount= " & AppData.AppModuleCount
        Print #1, "CompileType= " & AppData.CompileType

        
        Print #1, "----------------------------------"
        Print #1, "VB Project Infomation"
        Print #1, "----------------------------------"
        Print #1, "ProjectTitle= " & ProjectTitle
        Print #1, "ProjectName= " & ProjectName
        Print #1, "ExeName= " & ProjectExename
        Print #1, "HelpFile= " & HelpFile
        If gVBHeader.aSubMain <> 0 Then
            Print #1, "SubMain Address= " & gVBHeader.aSubMain + 1 - OptHeader.ImageBase
        End If
        Print #1, "ExternalComponentCount= " & gVBHeader.ExternalComponentCount
        Print #1, "----------------------------------"
        Print #1, "Object List"
        Print #1, "----------------------------------"
        For i = 0 To UBound(gObjectNameArray)
            Print #1, gObjectNameArray(i)
        Next i
     
        Print #1, "----------------------------------"
        Print #1, "External Ocx List"
        Print #1, "----------------------------------"
        If UBound(gOcxList) > 0 Then
            For i = 0 To UBound(gOcxList) - 1
                Print #1, gOcxList(i).strocxName
                Print #1, gOcxList(i).strName
                Print #1, gOcxList(i).strLibname
                Print #1, gOcxList(i).strGuid
                Print #1, ""
            Next i
        End If
        Print #1, "----------------------------------"
        Print #1, "Api List"
        Print #1, "----------------------------------"
        
        For i = 0 To UBound(gApiList) - 1
           'Print #1, "Declare " & gApiList(i).strFunctionName & " Lib " & cQuote & gApiList(i).strLibraryName & cQuote
            If Left$(gApiList(i).strFunctionName, 7) = "Declare" Then
                Print #1, gApiList(i).strFunctionName
            Else
                Print #1, "Declare " & gApiList(i).strFunctionName & " Lib " & cQuote & gApiList(i).strLibraryName & cQuote
            End If
        Next i
        Print #1, "----------------------------------"
        Print #1, "Controls Guids"
        Print #1, "----------------------------------"
        Print #1, "Parent Form, Control Name, GUID"
        For i = 0 To UBound(gControlNameArray)
            If gControlNameArray(i).strParentForm <> vbNullString Then
                Print #1, gControlNameArray(i).strParentForm & " , " & gControlNameArray(i).strControlName & " , " & gControlNameArray(i).strGuid
            End If
        Next i
        End If
    Close #1
    
    
End Sub
Sub WriteVBP(Filename As String)
'*****************************
'Purpose: Writes the visual basic project file
'*****************************
    Dim i As Integer
    Dim g As Integer
    Dim F As Long
    F = FreeFile
    Open Filename For Output As #F
        
        If gDllProject = False Then
            Print #F, "Type=Exe"
        Else
            If LCase$(Right$(SFilePath, 3)) = "dll" Then
                Print #F, "Type=OleDll"
            End If
            If LCase$(Right$(SFilePath, 3)) = "ocx" Then
                Print #F, "Type=Control"
            End If
        End If
        'If gVBHeader.ExternalComponentCount > 1 Then
           ' Print #3, "Reference="
           ' Print #3, "Object="
       ' End If
       
       If gVBHeader.ThreadFlags And 1 Then
        'ApartmentModel
       End If
       If gVBHeader.ThreadFlags And 2 Then
        'RequireLicense
        Print #F, "RequireLicenseKey=1"
       End If
       If gVBHeader.ThreadFlags And 4 Then
        'Unattended
        Print #F, "Unattended=-1"
       End If
       If gDllProject = True And gVBHeader.ThreadFlags And 8 Then
        'SingleThreaded
        Print #F, "ThreadingModel=0"
       End If
       If gVBHeader.ThreadFlags And &H10 And gVB6App = True Then
        'Retained
            Print #F, "Retained=1"
       End If
       
        
        'Form Loop
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 1 Then
                    Print #F, "Form=" & gObjectNameArray(i) & ".frm"
                    Exit For
                End If
            Next g
        Next i
        'Module Loop
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 2 Then
                    Print #F, "Module=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".bas"
                    Exit For
                End If
            Next g
        Next i
        'Class Loop
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 3 Then
                    Print #F, "Class=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".cls"
                    Exit For
                End If
            Next g
        Next i
        'UserControl Loop
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 4 Then
                   Print #F, "UserControl=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".ctl"
                    Exit For
                End If
            Next g
        Next i
        'Property Page Loop
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 5 Then
                    Print #F, "PropertyPage=" & gObjectNameArray(i) & ".pag"
                    Exit For
                End If
            Next g
        Next i
        'User Document Loop
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 6 Then
                    Print #F, "UserDocument=" & gObjectNameArray(i) & ".dob"
                    Exit For
                End If
            Next g
        Next i
        
        Dim strBuffer As String
        If gVBHeader.ExternalComponentCount > 0 Then
            For i = 0 To UBound(gOcxList) - 1
                If gOcxList(i).strGuid <> vbNullString Then
                    Dim strGuid As String
                    Dim strVersion As String
                    
            
                    'TypeLib
                    strGuid = modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & gOcxList(i).strGuid & "\TypeLib", "")
                    
                    'Version
                    strVersion = modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & gOcxList(i).strGuid & "\Version", "")
                    '
                    If InStr(1, strBuffer, "Object=" & strGuid & "#" & strVersion & "#0; " & gOcxList(i).strocxName) = False Then
                        strBuffer = strBuffer & "Object=" & strGuid & "#" & strVersion & "#0; " & gOcxList(i).strocxName & vbCrLf
                    'Print #F, "Object={" & strGuid & "}#" & strVersion & "#0; " & gOcxList(i).strocxName
                    End If
                End If
            Next i
            Print #F, strBuffer
        End If
        
        'Print #3, "IconForm=" & cQuote & DATAHERE & cQuote
        If gVBHeader.aSubMain = 0 Then
            Print #F, "Startup=" & cQuote & AppData.StartUpName & cQuote
        End If
        'Set Compile type
        If gProjectInfo.aNativeCode = 0 Then
            Print #F, "CompilationType=-1"
        Else
            Print #F, "CompilationType=0"
        End If
        Print #F, "Description=" & cQuote & ProjectDescription & cQuote
        Print #F, "HelpFile=" & cQuote & HelpFile & cQuote
        Print #F, "Name=" & cQuote & ProjectName & cQuote
        Print #F, "Title=" & cQuote & ProjectTitle & cQuote
        Print #F, "ExeName32=" & cQuote & ProjectExename & cQuote
       
        'Print version information
        If Len(gFileInfo.CompanyName) <> 0 Then
            Print #F, "VersionCompanyName=" & cQuote & gFileInfo.CompanyName & cQuote
        End If
        If Len(gFileInfo.FileDescription) <> 0 Then
            Print #F, "VersionFileDescription=" & cQuote & gFileInfo.FileDescription & cQuote
        End If
        If Len(gFileInfo.LegalCopyright) <> 0 Then
            Print #F, "VersionLegalCopyright=" & cQuote & gFileInfo.LegalCopyright & cQuote
        End If
        If Len(gFileInfo.Comments) <> 0 Then
            Print #F, "VersionComments=" & cQuote & gFileInfo.Comments & cQuote
        End If
        If Len(gFileInfo.LegalTradeMark) <> 0 Then
            Print #F, "VersionLegalTrademarks=" & cQuote & gFileInfo.LegalTradeMark & cQuote
        End If
        If Len(gFileInfo.ProductName) <> 0 Then
            Print #F, "VersionProductName=" & cQuote & gFileInfo.ProductName & cQuote
        End If
        
    Close #F
    
End Sub
Sub WriteVBPVB4(Filename As String)
'*****************************
'Purpose: Writes the visual basic project file
'*****************************
    Dim i As Integer
    Dim F As Long
    F = FreeFile
    Open Filename For Output As #F
        
        If gDllProject = False Then
            Print #F, "Type=Exe"
        Else
            If LCase$(Right$(SFilePath, 3)) = "dll" Then
                Print #F, "Type=OleDll"
            End If
            If LCase$(Right$(SFilePath, 3)) = "ocx" Then
                Print #F, "Type=Control"
            End If
        End If

        For i = 0 To UBound(strVB4Forms)
            If strVB4Forms(i) <> "" Then
                Print #F, "Form=" & strVB4Forms(i) & ".frm"
            End If
        Next

        If VB4Header.aSubMain = 0 Then
            Print #F, "Startup=" & cQuote & strVB4Forms(0) & cQuote
        End If
        'Set Compile type
        'If gProjectInfo.aNativeCode = 0 Then
            'Print #f, "CompilationType=-1"
        'Else
            'Print #f, "CompilationType=0"
        'End If
        Print #F, "Description=" & cQuote & ProjectDescription & cQuote
        Print #F, "HelpFile=" & cQuote & HelpFile & cQuote
        Print #F, "Name=" & cQuote & ProjectName & cQuote
        Print #F, "Title=" & cQuote & ProjectTitle & cQuote
        Print #F, "ExeName32=" & cQuote & ProjectExename & cQuote
       
        'Print version information
        If Len(gFileInfo.CompanyName) <> 0 Then
            Print #F, "VersionCompanyName=" & cQuote & gFileInfo.CompanyName & cQuote
        End If
        If Len(gFileInfo.FileDescription) <> 0 Then
            Print #F, "VersionFileDescription=" & cQuote & gFileInfo.FileDescription & cQuote
        End If
        If Len(gFileInfo.LegalCopyright) <> 0 Then
            Print #F, "VersionLegalCopyright=" & cQuote & gFileInfo.LegalCopyright & cQuote
        End If
        If Len(gFileInfo.Comments) <> 0 Then
            Print #F, "VersionComments=" & cQuote & gFileInfo.Comments & cQuote
        End If
        If Len(gFileInfo.LegalTradeMark) <> 0 Then
            Print #F, "VersionLegalTrademarks=" & cQuote & gFileInfo.LegalTradeMark & cQuote
        End If
        If Len(gFileInfo.ProductName) <> 0 Then
            Print #F, "VersionProductName=" & cQuote & gFileInfo.ProductName & cQuote
        End If
        
    Close #F
    
End Sub
Sub ShowVBPFileVB4()
'*****************************
'Purpose: To Show the VBP File in the textbox
'*****************************
    Dim i As Integer
    Dim strBuffer As String
    frmMain.txtCode.Text = vbNullString
    strBuffer = vbNullString
 
    If gDllProject = False Then
        strBuffer = strBuffer & "Type=Exe" & vbCrLf
    Else
        If LCase$(Right$(SFilePath, 3)) = "dll" Then
            strBuffer = strBuffer & "Type=OleDll" & vbCrLf
        End If
        If LCase$(Right$(SFilePath, 3)) = "ocx" Then
            strBuffer = strBuffer & "Type=Control" & vbCrLf
        End If
    End If

    For i = 0 To UBound(strVB4Forms)
        If strVB4Forms(i) <> "" Then
            strBuffer = strBuffer & "Form=" & strVB4Forms(i) & ".frm" & vbCrLf
        End If
    Next

        
    If VB4Header.aSubMain = 0 Then
        strBuffer = strBuffer & "Startup=" & cQuote & strVB4Forms(0) & cQuote & vbCrLf
    Else
        strBuffer = strBuffer & "Startup=" & cQuote & "Sub Main" & cQuote & vbCrLf
    
    End If
    'Set Compile type
   '' If gProjectInfo.aNativeCode = 0 Then
    '    strBuffer = strBuffer & "CompilationType=-1" & vbCrLf
   ' Else
   '     strBuffer = strBuffer & "CompilationType=0" & vbCrLf
   ' End If'
    
    strBuffer = strBuffer & "Description=" & cQuote & ProjectDescription & cQuote & vbCrLf
    strBuffer = strBuffer & "HelpFile=" & cQuote & HelpFile & cQuote & vbCrLf
    strBuffer = strBuffer & "Name=" & cQuote & ProjectName & cQuote & vbCrLf
    strBuffer = strBuffer & "Title=" & cQuote & ProjectTitle & cQuote & vbCrLf
    strBuffer = strBuffer & "ExeName32=" & cQuote & ProjectExename & cQuote & vbCrLf
    
    'Print version information
    If Len(gFileInfo.CompanyName) <> 0 Then
        strBuffer = strBuffer & "VersionCompanyName=" & cQuote & gFileInfo.CompanyName & cQuote & vbCrLf
    End If
    If Len(gFileInfo.FileDescription) <> 0 Then
        strBuffer = strBuffer & "VersionFileDescription=" & cQuote & gFileInfo.FileDescription & cQuote & vbCrLf
    End If
    If Len(gFileInfo.LegalCopyright) <> 0 Then
        strBuffer = strBuffer & "VersionLegalCopyright=" & cQuote & gFileInfo.LegalCopyright & cQuote & vbCrLf
    End If
    If Len(gFileInfo.Comments) <> 0 Then
        strBuffer = strBuffer & "VersionComments=" & cQuote & gFileInfo.Comments & cQuote & vbCrLf
    End If
    If Len(gFileInfo.LegalTradeMark) <> 0 Then
        strBuffer = strBuffer & "VersionLegalTrademarks=" & cQuote & gFileInfo.LegalTradeMark & cQuote & vbCrLf
    End If
    If Len(gFileInfo.ProductName) <> 0 Then
        strBuffer = strBuffer & "VersionProductName=" & cQuote & gFileInfo.ProductName & cQuote & vbCrLf
    End If
    'Show it now
    
    frmMain.txtCode.Text = strBuffer
End Sub

Sub ShowVBPFile()
'*****************************
'Purpose: To Show the VBP File in the textbox
'*****************************
  Dim i As Integer
If gVB4App = False And gVB5App = False And gVB6App = False Then
    frmMain.txtCode.Text = "Not a VB 4/5/6 file" & vbCrLf
    If bISVBNET = True Then
        frmMain.txtCode.Text = "This is a .Net application" & vbCrLf
    End If
    If VBVersion = 1 Then
        frmMain.txtCode.Text = "This is a VB 1 application not supported" & vbCrLf
    End If
    If VBVersion = 2 Then
        frmMain.txtCode.Text = "This is a VB 2 application not supported" & vbCrLf
    End If
    If VBVersion = 3 Then
        frmMain.txtCode.Text = "This is a VB 3 application not supported" & vbCrLf
    End If
    For i = 0 To UBound(SecHeader)
        If Left$(SecHeader(i).SecName, 3) = "UPX" Then
            frmMain.txtCode.Text = frmMain.txtCode.Text & "This application is protected by the UPX Packer" & vbCrLf
            Exit Sub
        End If
    Next
    Exit Sub
End If

  
    Dim g As Integer
    Dim strBuffer As String
    frmMain.txtCode.Text = vbNullString
    strBuffer = vbNullString
    
    If gDllProject = False Then
        strBuffer = strBuffer & "Type=Exe" & vbCrLf
    Else
        If LCase$(Right$(SFilePath, 3)) = "dll" Then
            strBuffer = strBuffer & "Type=OleDll" & vbCrLf
        End If
        If LCase$(Right$(SFilePath, 3)) = "ocx" Then
            strBuffer = strBuffer & "Type=Control" & vbCrLf
        End If
    End If

       If gVBHeader.ThreadFlags And 1 Then
        'ApartmentModel
       End If
       If gVBHeader.ThreadFlags And 2 Then
        'RequireLicense
        strBuffer = strBuffer & "RequireLicenseKey=1" & vbCrLf
       End If
       If gVBHeader.ThreadFlags And 4 Then
        'Unattended
        strBuffer = strBuffer & "Unattended=-1" & vbCrLf
       End If
       If gDllProject = True And gVBHeader.ThreadFlags And 8 Then
        'SingleThreaded
        strBuffer = strBuffer & "ThreadingModel=0" & vbCrLf
       End If
       If gVBHeader.ThreadFlags And &H10 And gVB6App = True Then
        'Retained
            strBuffer = strBuffer & "Retained=1" & vbCrLf
       End If

        'Form Loop
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 1 Then
                     strBuffer = strBuffer & "Form=" & gObjectNameArray(i) & ".frm" & vbCrLf
                    Exit For
                End If
            Next g
        Next i
        'Module Loop
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 2 Then
                     strBuffer = strBuffer & "Module=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".bas" & vbCrLf
                    Exit For
                End If
            Next g
        Next i
        'Class Loop
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 3 Then
                     strBuffer = strBuffer & "Class=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".cls" & vbCrLf
                    Exit For
                End If
            Next g
        Next i
        'UserControl Loop
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 4 Then
                    strBuffer = strBuffer & "UserControl=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".ctl" & vbCrLf
                    Exit For
                End If
            Next g
        Next i
        'Property Page Loop
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 5 Then
                    strBuffer = strBuffer & "PropertyPage=" & gObjectNameArray(i) & ".pag" & vbCrLf
                    Exit For
                End If
            Next g
        Next i
        'User Document Loop
        For i = 0 To UBound(gObject)
            For g = 0 To UBound(gObjectTypeList)
                If gObject(i).ObjectType = gObjectTypeList(g).value And gObjectTypeList(g).strType = 6 Then
                    strBuffer = strBuffer & "UserDocument=" & gObjectNameArray(i) & ".dob" & vbCrLf
                    Exit For
                End If
            Next g
        Next i
        
        'External Components
        Dim strEList As String
        If gVBHeader.ExternalComponentCount > 0 Then
            For i = 0 To UBound(gOcxList) - 1
                If gOcxList(i).strGuid <> vbNullString Then
                    Dim strGuid As String
                    Dim strVersion As String
                    
                    'TypeLib
                    strGuid = modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & gOcxList(i).strGuid & "\TypeLib", "")
                    'Debug.Print "SOFTWARE\Classes\CLSID\" & gOcxList(i).strGuid & "\TypeLib"
                    'Version
                    strVersion = modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & gOcxList(i).strGuid & "\Version", "")
                    '
                    'InProcServer32
                    'MsgBox modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & gOcxList(i).strGuid & "\InprocServer32", "")
                    
                   ' strBuffer = strBuffer & "Object=" & strGuid & "#" & strVersion & "#0; " & gOcxList(i).strocxName & vbCrLf
                  If InStr(1, strBuffer, "Object=" & strGuid & "#" & strVersion & "#0; " & gOcxList(i).strocxName) = False Then
  
                  strBuffer = strBuffer & "Object=" & strGuid & "#" & strVersion & "#0; " & gOcxList(i).strocxName & vbCrLf
                 
                 End If
                End If
            Next i
        End If
        
        
    If gVBHeader.aSubMain = 0 Then
        strBuffer = strBuffer & "Startup=" & cQuote & AppData.StartUpName & cQuote & vbCrLf
    Else
        strBuffer = strBuffer & "Startup=" & cQuote & "Sub Main" & cQuote & vbCrLf
    
    End If
    'Set Compile type
    If gProjectInfo.aNativeCode = 0 Then
        strBuffer = strBuffer & "CompilationType=-1" & vbCrLf
    Else
        strBuffer = strBuffer & "CompilationType=0" & vbCrLf
    End If
    
    strBuffer = strBuffer & "Description=" & cQuote & ProjectDescription & cQuote & vbCrLf
    strBuffer = strBuffer & "HelpFile=" & cQuote & HelpFile & cQuote & vbCrLf
    strBuffer = strBuffer & "Name=" & cQuote & ProjectName & cQuote & vbCrLf
    strBuffer = strBuffer & "Title=" & cQuote & ProjectTitle & cQuote & vbCrLf
    strBuffer = strBuffer & "ExeName32=" & cQuote & ProjectExename & cQuote & vbCrLf
    
    'Print version information
    If Len(gFileInfo.CompanyName) <> 0 Then
        strBuffer = strBuffer & "VersionCompanyName=" & cQuote & gFileInfo.CompanyName & cQuote & vbCrLf
    End If
    If Len(gFileInfo.FileDescription) <> 0 Then
        strBuffer = strBuffer & "VersionFileDescription=" & cQuote & gFileInfo.FileDescription & cQuote & vbCrLf
    End If
    If Len(gFileInfo.LegalCopyright) <> 0 Then
        strBuffer = strBuffer & "VersionLegalCopyright=" & cQuote & gFileInfo.LegalCopyright & cQuote & vbCrLf
    End If
    If Len(gFileInfo.Comments) <> 0 Then
        strBuffer = strBuffer & "VersionComments=" & cQuote & gFileInfo.Comments & cQuote & vbCrLf
    End If
    If Len(gFileInfo.LegalTradeMark) <> 0 Then
        strBuffer = strBuffer & "VersionLegalTrademarks=" & cQuote & gFileInfo.LegalTradeMark & cQuote & vbCrLf
    End If
    If Len(gFileInfo.ProductName) <> 0 Then
        strBuffer = strBuffer & "VersionProductName=" & cQuote & gFileInfo.ProductName & cQuote & vbCrLf
    End If
    'Show it now
    
    frmMain.txtCode.Text = strBuffer
End Sub
Sub WriteForms(FilePath As String, FormName As String, index As Integer)
'*****************************
'Purpose: To export the forms to a .frm file
'*****************************
Dim i As Integer
Dim nApi As Integer
On Error GoTo nofile:
        
       Dim F As Long
       F = FreeFile
        Open FilePath For Output As #F
            If gVB4App = False Then
                Print #F, "VERSION 5.00"
            Else
                Print #F, "VERSION 4.00"
            End If
            'Begin Object References
            If VBVersion <> 4 Then
                Dim iGuid As Integer
                For i = 0 To UBound(gExternalObjectHolder)
                    If gExternalObjectHolder(i).strFormName = FormName Then
                        For iGuid = 0 To UBound(gOcxList)
                            If gOcxList(iGuid).strLibname = gExternalObjectHolder(i).strLibname Then
                                Dim strGuid As String
                                Dim strVersion As String
                                
                                'TypeLib
                                strGuid = modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & gOcxList(iGuid).strGuid & "\TypeLib", "")
                                'Version
                                strVersion = modRegistry.RegQueryStringValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\" & gOcxList(iGuid).strGuid & "\Version", "")
                                Print #F, "Object = " & cQuote & "" & strGuid & "#" & strVersion & "#0" & cQuote & "; " & cQuote & gOcxList(iGuid).strocxName & cQuote
                                Exit For
                            End If
                        Next
                      
                    End If
                Next
                
                'Begin Form
                For i = 0 To frmMain.txtFinal.UBound
                    If frmMain.txtFinal(i).tag = modGlobals.gObjectNameArray(index) Then
                        Print #F, Trim(frmMain.txtFinal(i).Text)
                        Exit For
                    End If
                Next
            Else
            'VB4 Forms
                Print #F, frmMain.txtFinal(index).Text
            
            End If
            
            Print #F, "Attribute VB_Name = " & cQuote & frmMain.txtFinal(i).tag & cQuote
            Print #F, "Attribute VB_GlobalNameSpace = False"
            Print #F, "Attribute VB_Creatable = False"
            Print #F, "Attribute VB_PredeclaredId = True"
            Print #F, "Attribute VB_Exposed = False"
            
            'Print the procedures
            Print #F, "Option Explicit"
            Print #F, "'Generated by Semi VB Decompiler - VisualBasicZone.com"
            If VBVersion <> 4 Then
                If gProjectInfo.aNativeCode = 0 And modGlobals.gPcodeDecompile = True Then
                    Dim k As Integer
                    Dim strData As String
                    frmMain.FrameStatus.Visible = True
                    frmMain.txtStatus.Text = ""
                    For k = 0 To UBound(EventProcList) - 1
                        
                        If EventProcList(k) <> 0 Then
                           frmMain.txtStatus.Text = frmMain.txtStatus.Text & "Generating Form Event: " & EventProcList(k) & vbCrLf
                           frmMain.cmdSkipProcedure.Visible = True
                           strData = modPCode.DecompileProcToVB(EventProcList(k), True)
                        End If
                        Dim sTemp
                        sTemp = Split(strData, ".")
                        'MsgBox sTemp(0)
                        If sTemp(0) = FormName Then
                            Print #F, Replace(Replace(modPCode.DecompileProcToVB(EventProcList(k), False), ".", "_", 1, 1), ")()", ")", 1, 1)
                        End If
                    Next
                    'Generate Procedures
                    Dim ProcAddr() As Long
                    Dim g As Integer
                    Dim FileNum As Integer
                    FileNum = FreeFile
                    Open SFilePath For Binary Access Read As #FileNum
                    For k = 0 To UBound(gObjectInfoHolder)
                        If gObjectInfoHolder(k).NumberOfProcs > 0 Then
                        ReDim ProcAddr(gObjectInfoHolder(k).NumberOfProcs - 1)
                        Seek #FileNum, gObjectInfoHolder(k).aProcTable + 1 - OptHeader.ImageBase
                        Get #FileNum, , ProcAddr
                        For g = 0 To UBound(ProcAddr)
                            If ProcAddr(g) <> 0 And ProcAddr(g) <> -1 Then
                                If ProcAddr(g) < UBound(SubName) And ProcAddr(g) > LBound(SubName) Then
                                    SubName(ProcAddr(g)) = gObjectNameArray(k) & ".Proc" & ProcAddr(g)
                                    
                                    If FormName = gObjectNameArray(k) Then
                                    'MsgBox gObjectNameArray(k)
                                    'Generate procedure
                                        frmMain.txtStatus.Text = frmMain.txtStatus.Text & "Generating Form Procedure: " & ProcAddr(g) & vbCrLf
                                        Print #F, Replace(modPCode.DecompileProcToVB(ProcAddr(g)), ".", "", 1, 1)
                                    End If
                                End If
                            End If
                        Next
                        End If
                    Next
                    Close #FileNum
                    
                    frmMain.txtStatus.Text = ""
                    frmMain.FrameStatus.Visible = False
                
                End If
              
                For nApi = 0 To UBound(gProcedureList)
                    If UCase$(frmMain.txtFinal(i).tag) = UCase$(gProcedureList(nApi).strParent) Then
                        If gProcedureList(nApi).strProcedureName <> "" Then
                            If Right$(gProcedureList(nApi).strProcedureName, 1) = ")" Then
                                Print #F, "Private Sub " & gProcedureList(nApi).strProcedureName
                            Else
                                Print #F, "Private Sub " & gProcedureList(nApi).strProcedureName & "()"
                            End If
                            Print #F, "End Sub"
                        End If
                    End If
                Next
            
            End If

        Close #F
Exit Sub
nofile:
    MsgBox "Error_modOutput_WriteForms: " & err.Description
End Sub
Sub WriteFormFrx(FilePath As String, FormName As String, index As Integer)
'*****************************
'Purpose: Write the forms graphic files (.frx)
'*****************************
    Dim pFrxHeader As FRXITEMHDR
    Dim i As Integer
    Dim fFile As Long
    Dim PicFile As Long
    Dim bDeleteFrx As Boolean
    bDeleteFrx = False
    On Error Resume Next
    Kill FilePath & "\" & FormName & ".frx"

    
    fFile = FreeFile
    Open FilePath & "\" & FormName & ".frx" For Binary Access Write Lock Write As fFile
    For i = 0 To UBound(FrxPreview)
        If FrxPreview(i).ParentForm = FormName Then
            pFrxHeader.dwSizeImage = FrxPreview(i).length
            pFrxHeader.dwSizeImageEx = FrxPreview(i).length + 8
            pFrxHeader.dwKey = &H746C
            Put fFile, , pFrxHeader
            PicFile = FreeFile
            Dim Buffer() As Byte
            'Dim bEndByte As Integer
            ReDim Buffer(pFrxHeader.dwSizeImage)
            'bEndByte = 2573
            Open App.Path & "\dump\" & SFile & "\" & FrxPreview(i).strPath For Binary Access Read Lock Read As PicFile
                Get PicFile, , Buffer
                Put fFile, , Buffer
                Seek fFile, Loc(fFile)
                'Put fFile, , bEndByte
            Close PicFile
            
        End If
    Next
        If LOF(fFile) = 0 Then bDeleteFrx = True
    Close fFile
    If bDeleteFrx = True Then
        Kill FilePath & "\" & FormName & ".frx"
    End If
    
End Sub
Sub WriteModules(Filename As String, ObjectName As String, index As Integer)
'*****************************
'Purpose: To export the modules to a .bas file
'*****************************
    Dim F As Long
    F = FreeFile
    Open Filename For Output As #F
        Print #F, "Attribute VB_Name = " & cQuote & ObjectName & cQuote
        Print #F, "Option Explicit"
         Print #F, "'Generated by Semi VB Decompiler - VisualBasicZone.com"
                    'Generate Procedures
        If gProjectInfo.aNativeCode = 0 And modGlobals.gPcodeDecompile = True Then
                    Dim ProcAddr() As Long
                    Dim g As Integer
                    Dim k As Integer
                    Dim FileNum As Integer
                    FileNum = FreeFile
                    frmMain.FrameStatus.Visible = True
                    frmMain.txtStatus.Text = ""
                    frmMain.txtStatus.Visible = True
                    frmMain.cmdSkipProcedure.Visible = True
                    Open SFilePath For Binary Access Read As #FileNum
                    For k = 0 To UBound(gObjectInfoHolder)
                        If gObjectInfoHolder(k).NumberOfProcs > 0 Then
                        ReDim ProcAddr(gObjectInfoHolder(k).NumberOfProcs - 1)
                        Seek #FileNum, gObjectInfoHolder(k).aProcTable + 1 - OptHeader.ImageBase
                        Get #FileNum, , ProcAddr
                        For g = 0 To UBound(ProcAddr)
                            If ProcAddr(g) <> 0 And ProcAddr(g) <> -1 Then
                                If ProcAddr(g) < UBound(SubName) And ProcAddr(g) > LBound(SubName) Then
                                    SubName(ProcAddr(g)) = gObjectNameArray(k) & ".Proc" & ProcAddr(g)
                                    
                                    If ObjectName = gObjectNameArray(k) Then
                    
                                    'Generate procedure
                                        frmMain.txtStatus.Text = frmMain.txtStatus.Text & "Generating " & ObjectName & "  Procedure: " & ProcAddr(g) & vbCrLf
                                        Print #F, Replace(modPCode.DecompileProcToVB(ProcAddr(g)), ".", "", 1, 1)
                                    End If
                                End If
                            End If
                        Next
                        End If
                    Next
                    Close #FileNum
                    
                    frmMain.txtStatus.Text = ""
                    frmMain.FrameStatus.Visible = False
            End If
    
    Close #F
End Sub
Sub WriteClasses(Filename As String, ObjectName As String, index As Integer)
'*****************************
'Purpose: To export the classes to a .cls file
'*****************************
    Dim nApi As Integer
    Dim F As Long
    F = FreeFile
    Open Filename For Output As #F
        Print #F, "VERSION 1.0 CLASS"
        Print #F, "Begin"
        If gVB6App = True Then
            Print #F, "  MultiUse = -1  'True"
            Print #F, "  Persistable = 0  'NotPersistable"
            Print #F, "  DataBindingBehavior = 0  'vbNone"
            Print #F, "  DataSourceBehavior = 0   'vbNone"
            Print #F, "  MTSTransactionMode = 0   'NotAnMTSObject"
        End If
        Print #F, "End"
        Print #F, "Attribute VB_Name = " & cQuote & ObjectName & cQuote
        Print #F, "Attribute VB_GlobalNameSpace = False"
        Print #F, "Attribute VB_Creatable = True"
        Print #F, "Attribute VB_PredeclaredId = False"
        Print #F, "Attribute VB_Exposed = False"
        Print #F, "Attribute VB_Ext_KEY = " & cQuote & "SavedWithClassBuilder6" & cQuote & "," & cQuote & "Yes" & cQuote
        Print #F, "Attribute VB_Ext_KEY = " & cQuote & "Top_Level" & cQuote & " ," & cQuote & "No" & cQuote
        Print #F, "Option Explicit"
        Print #F, "'Generated by Semi VB Decompiler - VisualBasicZone.com"
        For nApi = 0 To UBound(gProcedureList)
            If UCase$(ObjectName) = UCase$(gProcedureList(nApi).strParent) Then
                If gProcedureList(nApi).strProcedureName <> "" Then
                    Print #F, "Sub " & gProcedureList(nApi).strProcedureName & "()"
                    Print #F, "End Sub"
                End If
            End If
        Next
        
    
    Close #F
End Sub
Sub WriteUserControls(Filename As String, ObjectName As String, index As Integer)
'*****************************
'Purpose: To export the controls to a .ctl file
'*****************************
    Dim F As Long
    Dim i  As Integer
    F = FreeFile
    Open Filename For Output As #F
        Print #F, "VERSION 5.00"
        For i = 0 To frmMain.txtFinal.UBound
            If frmMain.txtFinal(i).tag = modGlobals.gObjectNameArray(index) Then
                Print #F, frmMain.txtFinal(i).Text
                Exit For
            End If
        Next i
    Close #F
End Sub
Sub WritePropertyPage(Filename As String, ObjectName As String, index As Integer)
'*****************************
'Purpose: To export the controls to a .pag file
'*****************************
    Dim F As Long
    Dim i As Integer
    F = FreeFile
    Open Filename For Output As #F
        Print #F, "VERSION 5.00"
        For i = 0 To frmMain.txtFinal.UBound
            If frmMain.txtFinal(i).tag = modGlobals.gObjectNameArray(index) Then
                Print #F, frmMain.txtFinal(i).Text
                Exit For
            End If
        Next i
    Close #F
End Sub
Sub WriteDesigner(Filename As String, ObjectName As String, index As Integer)
'*****************************
'Purpose: To export the Designers to a .dsr file
'*****************************
    Dim F As Long
    F = FreeFile
    Open Filename For Output As #F
        Print #F, "VERSION 5.00"
        Print #F, "Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} " & ObjectName
        Print #F, " ClientHeight = 8775"
        Print #F, " ClientLeft = 1740"
        Print #F, " ClientTop = 1545"
        Print #F, " ClientWidth = 6585"
        Print #F, "End"
    Close #F
End Sub
Sub WriteUserDocument(ByVal Filename As String, ByVal ObjectName As String, ByVal index As Integer)
    Dim F As Long
    F = FreeFile
    Open Filename For Output As #F
        Print #F, "VERSION 5.00"
        Print #F, frmMain.txtFinal(index).Text
        Print #F, "Attribute VB_Name = " & cQuote & frmMain.txtFinal(index).tag & cQuote
        Print #F, "Attribute VB_GlobalNameSpace = False"
        Print #F, "Attribute VB_Creatable = True"
        Print #F, "Attribute VB_PredeclaredId = False"
        Print #F, "Attribute VB_Exposed = True"
            
    Close #F
    
End Sub
Public Sub AddExternalObject(strFormName As String, strLibname As String)
    Dim i As Integer
    For i = 0 To UBound(gExternalObjectHolder)
        If gExternalObjectHolder(i).strFormName = strFormName And gExternalObjectHolder(i).strLibname = strLibname Then
            Exit Sub
        End If
    Next i
    gExternalObjectHolder(UBound(gExternalObjectHolder)).strFormName = strFormName
    gExternalObjectHolder(UBound(gExternalObjectHolder)).strLibname = strLibname
    
    ReDim Preserve gExternalObjectHolder(UBound(gExternalObjectHolder) + 1)
    
End Sub


