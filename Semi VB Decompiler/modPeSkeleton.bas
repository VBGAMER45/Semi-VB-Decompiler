Attribute VB_Name = "modPeSkeleton"
'*********************************************
'modPESkeleton
'Orignally written by Sarge
'Modifed by vbamer45
'*********************************************
Option Explicit

'Common dialog constants
Const cdlOFNHideReadOnly = &H4
Const cdlOFNPathMustExist = &H800
Const cdlOFNFileMustExist = &H1000

'File constants
Public Const MAXNUMBERIMAGEENTRIES = 16  'Current PE file limit for number of DataDirectory directories
Public Const DOS_SIGNATURE = 23117    '"MZ" = 0x4D5A
Public Const PE_SIGNATURE = 17744    '"PE" + 0x00 = 0x50450000
Public Const NE_SIGNATURE = 17742
Public Const LENNAME = 8              'Length of Section Header names
Public Const MAXSECTIONS = 16        'Current PE file limit for number of SectionHeader sections
Public Const VBVERTEXT = "VB5!"      'Current compiler/linker output version
Public Const VBINTRPTRTEXT = "MSVBVM60.DLL"  'Current interpreter filename

'Application-wide variables
Global SFile As String         'Source file name
Global SFilePath As String     'Source file path
Global ErrorFlag As Boolean    'Generic error flag
Global InFileNumber As Integer
Global DecLoadOffset As Double 'App load address

'--------------------------Begin Object structures-----------------
Public Type VB_Signature                    'VB info
    VBVer As String * 4                     'BYTE * 4 = compiler/linker version
    VBIntrptr As String * 12                'BYTE * 12 = interpreter filename
End Type

'General VB6 application data
Public Type App_Data                    'Application specific data for RACE
    DosHeaderOffset As Integer       'DWORD = location of DOS header (offset)in app
    PeHeaderOffset As Double         'DWORD = location of PE header (offset)in app
    OptHeaderOffset As Double        'DWORD = location of OPT header (offset)in app
    SecHeaderOffset As Double        'DWORD = location of SEC header (offset)in app
    VBStartOffset As Double          'DWORD = location of VB application (offset)in app
    VBVerOffsetRaw As Double         'DWORD = location of VB signature (address) in app
    VBVerOffsetMasked As Double      'DWORD = location of VB signature (offset)in app
    VBIntrptrOffset As Double        'DWORD = location of VB interpreter (offset)in app
    ProjDataAppReference As Double   'DWORD = location of RACE reference (offset) in app
    AppModuleCount As Byte           'BYTE = number of VB modules in app
    CompileType As String            'PCODE or NCODE
    StartUpName As String            'Name of first referenced object (if any)
    StartUpType As Double            'Signature of startup object
    StartUpOffset As Double          'Address of object block start
End Type

'General VB6 application data
Public Type VBStart_Header
    PushStartOpcode As Integer      'BYTE = opcode 68 = push
    PushStartAddress As Double      'DWORD = address of VB signature
    CallStartOpcode As Integer      'BYTE = opcode E8 = jmp
    CallStartAddress As Double      'DWORD = address of interpreter entry
End Type

'Generic DOS file data
Public Type Dos_Header              'Standard DOS header
    Magic As Double                 'WORD
    NumBytesLastPage As Double      'WORD
    NumPages As Double              'WORD
    NumRelocates As Double          'WORD
    NumHeaderBlks As Double         'WORD
    NumMinBlks As Double            'WORD
    NumMaxBlks As Double            'WORD
    SSPointer As Double             'WORD
    SPPointer As Double             'WORD
    Checksum As Double              'WORD
    IPPointer As Double             'WORD
    CurrentSeg As Double            'WORD
    RelocTablePointer As Double     'WORD
    Overlay As Double               'WORD
    ReservedW1 As Double            'WORD
    ReservedW2 As Double            'WORD
    ReservedW3 As Double            'WORD
    ReservedW4 As Double            'WORD
    OEMType As Double               'WORD
    OEMData As Double               'WORD
    ReservedW5 As Double            'WORD
    ReservedW6 As Double            'WORD
    ReservedW7 As Double            'WORD
    ReservedW8 As Double            'WORD
    ReservedW9 As Double            'WORD
    ReservedW10 As Double           'WORD
    ReservedW11 As Double           'WORD
    ReservedW12 As Double           'WORD
    ReservedW13 As Double           'WORD
    ReservedW14 As Double           'WORD
    ExeHeaderPointer As Double      'DWORD
End Type

Public Type NE_Header
    signature As Double 'WORD
    VersionLinker As Double 'Byte
    RevisionLinker As Double 'Byte
    EntryTableOffset As Double 'WORD
    SizeOfEntryTable As Double 'WORD
    CRC As Double 'DWORD
    flags As Double 'WORD
    SegmentNumberAutomaticDataSegment As Double 'WORD
    InitialSizeHeap As Double 'WORD
    InitialSizeStack As Double 'WORD
    SegmentNumberOffsetCS As Double 'DWORD
    SegmentNumberOffsetSS As Double 'DWORD
    NumberEntriesSegmentTable As Double 'WORD
    NumberEntriesModuleReferenceTable As Double 'WORD
    SizeOfNonResidentNameTable As Double 'WORD
    SegmentTableOffset As Double 'WORD
    ResourceTableFileOffset As Double 'WORD
    ResidentNameTableOffset As Double 'WORD
    ModuleReferenceTableOffset As Double 'WORD
    ImportedNamesTableOffset As Double 'WORD
    NonResidentNameTableOffset As Double  'Dword
    NumberMovableEntriesInEntryTable As Double 'WORD
    LogicalSectorAlignmentShiftCount As Double 'WORD
    NumberResourceEntries As Double 'WORD
    ExecutableType As Double 'Byte
    Reserved1 As Double 'DWORD
    Reserved2 As Double 'DWORD
End Type

'PE file data
Public Type PE_Header               'Standard PE header
    Magic As Double                 'DWORD
    Machine As Double               'WORD
    NumSections As Double           'WORD
    TimeDate As Double              'DWORD
    SymbolTablePointer As Double    'DWORD
    NumSymbols As Double            'DWORD
    OptionalHdrSize As Double       'WORD
    Properties As Double            'WORD
End Type

'PE file data
Public Type Data_Dir                'Standard Data Directory
    name As String                  'Variable
    Address As Double               'DWORD
    Size As Double                  'DWORD
End Type

'PE file data
Public Type Opt_Header              'Standard Option Header
    Magic As Double                 'WORD
    MajLinkerVer As Integer         'BYTE
    MinLinkerVer As Integer         'BYTE
    CodeSize As Double              'DWORD
    InitDataSize As Double          'DWORD
    UninitDataSize As Double        'DWORD
    EntryPoint As Double            'DWORD
    CodeBase As Double              'DWORD
    DataBase As Double              'DWORD
    ImageBase As Double             'DWORD
    SectionAlignment As Double      'DWORD
    FileAlignment As Double         'DWORD
    MajOSVer As Double              'WORD
    MinOSVer As Double              'WORD
    MajImageVer As Double           'WORD
    MinImageVer As Double           'WORD
    MajSSysVer As Double            'WORD
    MinSSysVer As Double            'WORD
    Win32Ver As Double              'DWORD
    SizeImage As Double             'DWORD
    SizeHeader As Double            'DWORD
    Checksum As Double              'DWORD
    SSystem As Double               'WORD
    DLLProperties As Double         'WORD
    SSizeRes As Double              'DWORD
    SSizeCom As Double              'DWORD
    HSizeRes As Double              'DWORD
    HSizeCom As Double              'DWORD
    LFlags As Double                'DWORD
    NumRVA_Sizes As Double          'DWORD
    DataDirectory(MAXNUMBERIMAGEENTRIES) As Data_Dir
End Type

'PE file data
Public Type Sec_Header              'Standard Section Header
    SecName As String * LENNAME     'BYTE[LENNAME]
    Misc As Double                  'DWORD
    Address As Double               'DWORD
    SizeRawData As Double           'DWORD
    RawDataPointer As Double        'DWORD
    RelocationPointer As Double     'DWORD
    LineNumPointer As Double        'DWORD
    NumRelocations As Double        'WORD
    NumLineNumbers As Double        'WORD
    Properties As Double            'DWORD
End Type


Private Type IMAGE_IMPORT_DESCRIPTOR
    lpImportByName As Long ''\\ 0 for terminating null import descriptor
    TimeDateStamp As Long  ''\\ 0 if not bound,
                           ''\\ -1 if bound, and real date\time stamp
                           ''\\ in IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT (new BIND)
                           ''\\ O.W. date/time stamp of DLL bound to (Old BIND)
    ForwarderChain As Long ''\\ -1 if no forwarders
    lpName As Long
    lpFirstThunk As Long ''\\ RVA to IAT (if bound this IAT has actual addresses)
End Type

Private Type IMAGE_EXPORT_DIRECTORY
    Characteristics As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    lpName As Long
    base As Long
    NumberOfFunctions As Long
    NumberOfNames As Long
    lpAddressOfFunctions As Long    '\\ Three parrallel arrays...(LONG)
    lpAddressOfNames As Long        '\\ (LONG)
    lpAddressOfNameOrdinals As Long '\\ (INTEGER)
End Type

Public Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

Public Type UPX_STRUCT
   upxMagic As String * 4     ' ??????? "UPX!"
   upxVersion As Byte         ' ?????? UPX'? (????????: 0C ??????
   upxFormat As Byte          ' ?????????? ?????? ????? (PE, ELF, DOS
   upxMethod As Byte          ' ????? ?????? (???? NRV ??? UCL, ?? 02)
   upxLevel As Byte           ' ??????? ?????? (?? 0 ?? 10)
   upxU_adler As Long         ' CRC ????? ????????? ? ????????????? ????
   upxC_adler As Long         ' CRC ????? ????????? ? ???????????? ????
   upxU_len As Long           ' ?????? ????? ????????? ? ????????????? ????
   upxC_len As Long           ' ?????? ????? ????????? ? ???????????? ????
   upxU_file_size As Long     ' ?????? ?????????????? ?????????.
   upxFilter As Integer       ' ????? ??????????
   upxCRC As Byte             ' CRC ?????????
End Type
Global UPXStruct As UPX_STRUCT

Public AppData As App_Data
Public VBStartHeader As VBStart_Header
Public VBSignature As VB_Signature
Public PEImport() As IMAGE_IMPORT_DESCRIPTOR
Public PeExport As IMAGE_EXPORT_DIRECTORY
Public DosHeader As Dos_Header
Public PEHeader As PE_Header
Public NEHeader As NE_Header
Public OptHeader As Opt_Header
Public SecHeader(MAXSECTIONS) As Sec_Header
Public bNEFormat As Boolean
Private Type IMPORT_API_LOOKUP
    ApiName As String
    Address As Long
    InternalAddr As Long
End Type
Public exeIMPORT_APINAME() As IMPORT_API_LOOKUP
Public VBVersion As Integer
'Used to get correct entry point for VB4/VB5
Global mImageBaseAlign As Double

Public Function CheckHeader() As Boolean

    'Assume a good file
    CheckHeader = True
    gDllProject = False
    VBVersion = 0 'Set version to zero
    
    '************************
    'All files must start with the DOS signature
    '************************
       
    'Save the DOSHeader offset (always 0x0000!)
    AppData.DosHeaderOffset = 0
    
    'Get the DOS signature, check for error
    Call GetDOSSignature
    
    If ErrorFlag = True Then
        CheckHeader = False
        Exit Function
    End If
        
    '*******************************
    'The DOS header follows the DOS signature
    '*******************************
     
    'Get the Dos header, check for error
    Call GetDOSHeader
    
    If ErrorFlag = True Then
        CheckHeader = False
        Exit Function
    End If
       
    '*******************************
    'The DOS header holds the PE file signature
    '*******************************
         
    'Move to the location where the PE signature should be
    Seek #InFileNumber, DosHeader.ExeHeaderPointer + 1
    
    'Get the PE signature, check error
    Call GetPESignature
                      
    If ErrorFlag = True Then
        CheckHeader = False
        Exit Function
    End If
    'If app is NE Format
    If bNEFormat = True Then
        Exit Function
    End If
    '************************************
    'The PEfile header data exists just after the PE signature
    '************************************
   
    'Save the PEHeader offset
    AppData.PeHeaderOffset = Seek(InFileNumber) - 1
    
    'Get the PEHeader
    Call GetPEHeader

    '************************************
    'The OPTion header exists just after the PE header
    '***********************************
      
    'Save the OPTHeader offset
    AppData.OptHeaderOffset = Seek(InFileNumber) - 1
        
    'Get the OPTHeader
    Call GetPEOptionHeader
      
    '***********************************
    'The SECtion headers exist just after the option header
    '***********************************
       
    'Save the SECtionHeader offset
    AppData.SecHeaderOffset = Seek(InFileNumber) - 1
     
    'Get the SecHeader
    Call GetPESecHeader

    bISVBNET = False
    If OptHeader.DataDirectory(14).Address <> 0 Then

        bISVBNET = True
        Call ProccessVBNETFile(GetPtrFromRVA(OptHeader.DataDirectory(14).Address) + 1, InFileNumber)
        MsgBox "This is a VB.NET compiled file and is not supported right now... View the .Net Console under the Tools menu.", vbExclamation
        VBVersion = 7
    End If
    
    'These sections are not included; they're
    'not needed for VB6 analysis, but could be
    'added if more PE file analysis is desired:
        'DebugDirectory
        'ResourceSection
        'ImportsSection
        'Needed for Pcode

        Seek #InFileNumber, GetPtrFromRVA2(OptHeader.DataDirectory(1).Address) + 1

        Dim ImportHolder As IMAGE_IMPORT_DESCRIPTOR
        ReDim PEImport(0)
        ReDim ImportList(0)
       ' Do

            Get #InFileNumber, , ImportHolder

           '' If ImportHolder.lpName = 0 Then Exit Do
           'Save it in the import table
        
            ReDim PEImport(UBound(PEImport) + 1)
            PEImport(UBound(PEImport)).ForwarderChain = ImportHolder.ForwarderChain
            PEImport(UBound(PEImport)).lpFirstThunk = ImportHolder.lpFirstThunk
            PEImport(UBound(PEImport)).lpImportByName = ImportHolder.lpImportByName
            PEImport(UBound(PEImport)).lpName = ImportHolder.lpName


            PEImport(UBound(PEImport)).TimeDateStamp = ImportHolder.TimeDateStamp
            
            
            Seek InFileNumber, GetPtrFromRVA2(PEImport(UBound(PEImport)).lpName) + 1
     
            
            ImportList(0).strName = GetUntilNull(InFileNumber)
            
            'Check the version of the vb runtime
            'MsgBox ImportList(0).strName
            If ImportList(0).strName = "VB40032.DLL" Then
                'MsgBox "This is a VB4 file and is not supported", vbExclamation
                Call modNative.VBFunction_Description_Init(App.Path & "\data\vb4api.txt")
                gVB4App = True
                VBVersion = 4
                mImageBaseAlign = ((OptHeader.ImageBase + OptHeader.EntryPoint) - GetPtrFromRVA(OptHeader.EntryPoint))
                
                VBStartHeader.PushStartAddress = GetPtrFromRVA2(OptHeader.EntryPoint)
                Exit Function
            End If
            If ImportList(0).strName = "MSVBVM50.DLL" Then
                'Load the VB5 imports
                gVB5App = True
                Call modNative.VBFunction_Description_Init(App.Path & "\data\vb5api.txt")
                 VBVersion = 5
            End If

            If ImportList(0).strName = "MSVBVM60.DLL" Then
                'Load the VB6 imports
                Call modNative.VBFunction_Description_Init(App.Path & "\data\vb6api.txt")
                gVB6App = True
                 VBVersion = 6
            End If
            ReDim exeIMPORT_APINAME(1 To 1)
            Call ScanTable(InFileNumber, GetPtrFromRVA2(PEImport(UBound(PEImport)).lpFirstThunk) + 1, GetPtrFromRVA2(PEImport(UBound(PEImport)).lpImportByName) + 1, exeIMPORT_APINAME())
     
       ' Loop
        'ExportsSection
        'Needed for dll's and ocx's
        'Used for dll projects

        If OptHeader.DataDirectory(0).Address <> 0 Then
     'On Error GoTo badDll
            Dim ExportPointer As Long
            'Get Dll Header
            Seek #InFileNumber, OptHeader.DataDirectory(0).Address + 29
            Get #InFileNumber, , ExportPointer
            Seek #InFileNumber, ExportPointer + 1
            Get #InFileNumber, , ExportPointer
            Seek #InFileNumber, ExportPointer + 3
            Get #InFileNumber, , ExportPointer
            
            VBStartHeader.PushStartAddress = ExportPointer
            gDllProject = True
  
            Seek #InFileNumber, GetPtrFromRVA2(OptHeader.DataDirectory(0).Address) + 1
          
            Get #InFileNumber, , PeExport
            Dim ExportName() As Long
            Dim ExportOrdinal() As Long
            Dim ExportProcedure() As Integer
            ReDim ExportName(PeExport.NumberOfNames - 1)
            ReDim ExportOrdinal(PeExport.NumberOfFunctions - 1)
            ReDim ExportProcedure(PeExport.NumberOfFunctions - 1)
            'Get Name array
            Seek #InFileNumber, GetPtrFromRVA2(PeExport.lpAddressOfNames) + 1
            Get #InFileNumber, , ExportName
            Dim strHolder As String
            Dim i As Integer
            For i = 0 To UBound(ExportName)
                Seek #InFileNumber, GetPtrFromRVA2(ExportName(i)) + 1
                strHolder = GetUntilNull(InFileNumber)
    
                If strHolder = "DllCanUnloadNow" Then
                    Seek #InFileNumber, GetPtrFromRVA2(ExportProcedure(i)) + 1
                    gVB5App = True
                    gVB6App = True
                End If
            Next
            'Get Ordinal Array
            Seek #InFileNumber, GetPtrFromRVA2(PeExport.lpAddressOfNameOrdinals) + 1
            Get #InFileNumber, , ExportOrdinal

            'Get Procedure Array
            Dim bLong As Long
            Seek #InFileNumber, GetPtrFromRVA2(PeExport.lpAddressOfFunctions) + 1
            ''Get #InFileNumber, , ExportProcedure
            'For i = 0 To UBound(ExportProcedure)
            bLong = GetDWord
'            ExportProcedure(0) = bLong
            'Next

            If bLong > 0 And gVB6App = True Then
                'Dim oData As Integer
                Dim oData As Long
                
                Seek #InFileNumber, GetPtrFromRVA2(bLong) + 1
                'Get #InFileNumber, GetPtrFromRVA2(bLong) + 3, oData
                Seek #InFileNumber, GetPtrFromRVA2(bLong) + 3
                oData = GetWord
            
                mImageBaseAlign = ((OptHeader.ImageBase + OptHeader.EntryPoint) - GetPtrFromRVA2(OptHeader.EntryPoint))
                AppData.VBStartOffset = GetPtrFromRVA2(oData)
                VBStartHeader.PushStartAddress = GetPtrFromRVA2(oData) + mImageBaseAlign
                gDllProject = True
            End If
            If VBStartHeader.PushStartAddress < 0 Then
                CheckHeader = False
                Exit Function
            End If
            Seek #InFileNumber, AppData.VBStartOffset + 1
            Call GetVBVer
            If ErrorFlag = True Then
                CheckHeader = False
            End If
           Exit Function
        End If
    '****************************
    'Start the VB app analysis
    '****************************
        
    'Calculate the load offset mask
    DecLoadOffset# = OptHeader.ImageBase
    Dim bVBVersion As Byte
On Error GoTo NormalVB

    mImageBaseAlign = ((OptHeader.ImageBase + OptHeader.EntryPoint) - GetPtrFromRVA(OptHeader.EntryPoint))
    Seek #InFileNumber, GetPtrFromRVA(OptHeader.EntryPoint + 1)
    bVBVersion = GetByte
    

    'VB5
    If bVBVersion = 104 Then
        Seek #InFileNumber, GetPtrFromRVA(OptHeader.EntryPoint + 1) + 1
        Dim vbStart As Long

        Get #InFileNumber, , vbStart
        VBStartHeader.PushStartAddress = vbStart

        Exit Function
    End If
    
    'Ocx'

    '**************************************
    'The VB Startheader holds the jump vector
    '**************************************
NormalVB:
    'Get the APP data VB app start location = OPTHeader.EntryPoint
    AppData.VBStartOffset = OptHeader.EntryPoint
    mImageBaseAlign = OptHeader.ImageBase
    'Point file at the VB code start position
    Seek #InFileNumber, AppData.VBStartOffset + 1
    
    'Get the VBStartHeader, check error
    Call GetVBStartHeader

    If ErrorFlag = True Then
        CheckHeader = False
        Exit Function
    End If
          
    '**************************************
    'The VB start vector holds the compiler signature
    '**************************************
    
    'Get the APP data VB signature offset
    AppData.VBVerOffsetRaw = VBStartHeader.PushStartAddress
    
    'Calculate the APP offset
    AppData.VBVerOffsetMasked = AppData.VBVerOffsetRaw - DecLoadOffset#

    'Point file at the VB signature position
    On Error Resume Next
    Seek #InFileNumber, AppData.VBVerOffsetMasked + 1
    
    'Check for VB version (compiler) of this file, check error
    Call GetVBVer
  
    If ErrorFlag = True Then
        CheckHeader = False
        Exit Function
    End If
    
    'Assign this location to our reference
    AppData.ProjDataAppReference = AppData.VBVerOffsetMasked
           
    '*****************************
    'Check if the interpreter name exists
    '*****************************
    
    'Point file at the Data Directory #1 position
    Seek #InFileNumber, OptHeader.DataDirectory(1).Address + 1
    
    'Move ahead 12 bytes
    Seek #InFileNumber, Seek(InFileNumber) + 12
    
    'Get the APP data interpreter address offset
    AppData.VBIntrptrOffset = GetDWord()
    
    'Move to the interpreter signature
    Seek #InFileNumber, AppData.VBIntrptrOffset + 1
    
    'Get the interpreter
    Call GetVBIntrptr
    VBVersion = 6
    gVB6App = True
    If ErrorFlag = True Then
        'CheckHeader = False
        Exit Function
    End If

    
    'If we got here, this is definitely a valid VB6 app
Exit Function
badDll:
        CheckHeader = False
        Exit Function
End Function

Public Sub GetDOSSignature()

    'Get the first two characters
    DosHeader.Magic = GetWord()
    
    'Check for error
    If DosHeader.Magic <> DOS_SIGNATURE Then
        ErrorFlag = True
    End If
    
End Sub

Public Sub GetDOSHeader()

    'Get DOS header data
        DosHeader.NumBytesLastPage = GetWord()
        DosHeader.NumPages = GetWord()
        DosHeader.NumRelocates = GetWord()
        DosHeader.NumHeaderBlks = GetWord()
        DosHeader.NumMinBlks = GetWord()
        DosHeader.NumMaxBlks = GetWord()
        DosHeader.SSPointer = GetWord()
        DosHeader.SPPointer = GetWord()
        DosHeader.Checksum = GetWord()
        DosHeader.IPPointer = GetWord()
        DosHeader.CurrentSeg = GetWord()
        DosHeader.RelocTablePointer = GetWord()
        DosHeader.Overlay = GetWord()
        DosHeader.ReservedW1 = GetWord()
        DosHeader.ReservedW2 = GetWord()
        DosHeader.ReservedW3 = GetWord()
        DosHeader.ReservedW4 = GetWord()
        DosHeader.OEMType = GetWord()
        DosHeader.OEMData = GetWord()
        DosHeader.ReservedW5 = GetWord()
        DosHeader.ReservedW6 = GetWord()
        DosHeader.ReservedW7 = GetWord()
        DosHeader.ReservedW8 = GetWord()
        DosHeader.ReservedW9 = GetWord()
        DosHeader.ReservedW10 = GetWord()
        DosHeader.ReservedW11 = GetWord()
        DosHeader.ReservedW12 = GetWord()
        DosHeader.ReservedW13 = GetWord()
        DosHeader.ReservedW14 = GetWord()
        DosHeader.ExeHeaderPointer = GetDWord()
        
        'Make sure the potential PE signature location seems reasonable
        If ((DosHeader.ExeHeaderPointer > 4096) Or (DosHeader.ExeHeaderPointer < 64)) Then
            ErrorFlag = True
        End If
        
End Sub

Public Sub GetPESignature()

    'Get the first two characters
    PEHeader.Magic = GetWord()
    bNEFormat = False
    If PEHeader.Magic = PE_SIGNATURE Then
        bNEFormat = False
        GetWord
    End If
   
    If PEHeader.Magic = NE_SIGNATURE Then
        bNEFormat = True
        Call modNeFormat.GetNEHeader
        Call modNeFormat.ProcessNeFile
        'Get the VB Version
        VBVersion = ReturnVBVersionNe(InFileNumber)
    End If
    'Check for error
    If PEHeader.Magic <> PE_SIGNATURE Then
        ErrorFlag = True
    End If
    
End Sub

Public Sub GetPEOptionHeader()

    'Now get the "optional" header data
    OptHeader.Magic = GetWord()
    OptHeader.MajLinkerVer = GetByte()
    OptHeader.MinLinkerVer = GetByte()
    OptHeader.CodeSize = GetDWord()
    OptHeader.InitDataSize = GetDWord()
    OptHeader.UninitDataSize = GetDWord()
    
    OptHeader.EntryPoint = GetDWord()
    OptHeader.CodeBase = GetDWord()
    OptHeader.DataBase = GetDWord()
    OptHeader.ImageBase = GetDWord()
    
    OptHeader.SectionAlignment = GetDWord()

    OptHeader.FileAlignment = GetDWord()
   
    OptHeader.MajOSVer = GetWord()
    OptHeader.MinOSVer = GetWord()
    OptHeader.MajImageVer = GetWord()
    OptHeader.MinImageVer = GetWord()
    OptHeader.MajSSysVer = GetWord()
    OptHeader.MinSSysVer = GetWord()
    OptHeader.Win32Ver = GetDWord()
    OptHeader.SizeImage = GetDWord()
    OptHeader.SizeHeader = GetDWord()
    OptHeader.Checksum = GetDWord()
    OptHeader.SSystem = GetWord()
    OptHeader.DLLProperties = GetWord()
    OptHeader.SSizeRes = GetDWord()
    OptHeader.SSizeCom = GetDWord()
    OptHeader.HSizeRes = GetDWord()
    OptHeader.HSizeCom = GetDWord()
    OptHeader.LFlags = GetDWord()
    OptHeader.NumRVA_Sizes = GetDWord()
    OptHeader.DataDirectory(0).name = "EXPORT"
    OptHeader.DataDirectory(0).Address = GetDWord()
    OptHeader.DataDirectory(0).Size = GetDWord()
    OptHeader.DataDirectory(1).name = "IMPORT"
    'MsgBox Loc(InFileNumber)
    OptHeader.DataDirectory(1).Address = GetDWord()
    'MsgBox Loc(InFileNumber)
    'MsgBox OptHeader.DataDirectory(1).Address
    OptHeader.DataDirectory(1).Size = GetDWord()
    OptHeader.DataDirectory(2).name = "RESOURCE"
    OptHeader.DataDirectory(2).Address = GetDWord()
    OptHeader.DataDirectory(2).Size = GetDWord()
    OptHeader.DataDirectory(3).name = "EXCEPTION"
    OptHeader.DataDirectory(3).Address = GetDWord()
    OptHeader.DataDirectory(3).Size = GetDWord()
    OptHeader.DataDirectory(4).name = "SECURITY"
    OptHeader.DataDirectory(4).Address = GetDWord()
    OptHeader.DataDirectory(4).Size = GetDWord()
    OptHeader.DataDirectory(5).name = "BASERELOC"
    OptHeader.DataDirectory(5).Address = GetDWord()
    OptHeader.DataDirectory(5).Size = GetDWord()
    OptHeader.DataDirectory(6).name = "DEBUG"
    OptHeader.DataDirectory(6).Address = GetDWord()
    OptHeader.DataDirectory(6).Size = GetDWord()
    OptHeader.DataDirectory(7).name = "COPYRIGHT"
    OptHeader.DataDirectory(7).Address = GetDWord()
    OptHeader.DataDirectory(7).Size = GetDWord()
    OptHeader.DataDirectory(8).name = "GLOBALPTR"
    OptHeader.DataDirectory(8).Address = GetDWord()
    OptHeader.DataDirectory(8).Size = GetDWord()
    OptHeader.DataDirectory(9).name = "TLS"
    OptHeader.DataDirectory(9).Address = GetDWord()
    OptHeader.DataDirectory(9).Size = GetDWord()
    OptHeader.DataDirectory(10).name = "LOAD_CONFIG"
    OptHeader.DataDirectory(10).Address = GetDWord()
    OptHeader.DataDirectory(10).Size = GetDWord()
    OptHeader.DataDirectory(11).name = "BOUND_IMPORT"
    OptHeader.DataDirectory(11).Address = GetDWord()
    OptHeader.DataDirectory(11).Size = GetDWord()
    OptHeader.DataDirectory(12).name = "IAT"
    OptHeader.DataDirectory(12).Address = GetDWord()
    OptHeader.DataDirectory(12).Size = GetDWord()
    OptHeader.DataDirectory(13).name = "DELAY_IMPORT"
    OptHeader.DataDirectory(13).Address = GetDWord()
    OptHeader.DataDirectory(13).Size = GetDWord()
    OptHeader.DataDirectory(14).name = "COM_DESCRIPTOR"
    OptHeader.DataDirectory(14).Address = GetDWord()
    OptHeader.DataDirectory(14).Size = GetDWord()
    OptHeader.DataDirectory(15).name = "unused"
    OptHeader.DataDirectory(15).Address = GetDWord()
    OptHeader.DataDirectory(15).Size = GetDWord()
    
End Sub

Public Sub GetPEHeader()
   
    'Fill up the PE header structure
    PEHeader.Machine = GetWord()
    PEHeader.NumSections = GetWord()
    PEHeader.TimeDate = GetDWord()
    PEHeader.SymbolTablePointer = GetDWord()
    PEHeader.NumSymbols = GetDWord()
    PEHeader.OptionalHdrSize = GetWord()
    PEHeader.Properties = GetWord()

End Sub

Public Sub GetPESecHeader()

    Dim Counter As Integer
    Dim Counter2 As Integer
        
    'Name = SecName
    'Address = RVA
    'RawDataPointer = Offset
    'Size = SizeRawData
    'Flags = Properties
    
    'All names are max 8 chars, starting with "."; get all 8 chars, remove the
    'trailing blanks, loop back until all sections done
    For Counter = 1 To PEHeader.NumSections
        SecHeader(Counter).SecName = Input(LENNAME, #InFileNumber%)
        For Counter2 = 1 To LENNAME
            'Remove the trailing blanks
            If Asc(Mid$(SecHeader(Counter).SecName, Counter2, 1)) = 0 Then
                Mid$(SecHeader(Counter).SecName, Counter2, 1) = vbNullString
            End If
        Next
        
        'Fill in the rest of the structure
        SecHeader(Counter).Misc = GetDWord()
        SecHeader(Counter).Address = GetDWord()
        SecHeader(Counter).SizeRawData = GetDWord()
        SecHeader(Counter).RawDataPointer = GetDWord()
        SecHeader(Counter).RelocationPointer = GetDWord()
        SecHeader(Counter).LineNumPointer = GetDWord()
        SecHeader(Counter).NumRelocations = GetWord()
        SecHeader(Counter).NumLineNumbers = GetWord()
        SecHeader(Counter).Properties = GetDWord()
    Next
    
End Sub

Public Sub GetVBStartHeader()

    'All VB start headers are a "Push" (5 bytes), followed by
    'a "Call" (5 bytes)

    VBStartHeader.PushStartOpcode = GetByte()
    VBStartHeader.PushStartAddress = GetDWord()
    VBStartHeader.CallStartOpcode = GetByte()
    VBStartHeader.CallStartAddress = GetDWord()

    'This should be the hex code "68 xx xx xx xx"
    If VBStartHeader.PushStartOpcode <> ConvertHex("68") And VBStartHeader.PushStartOpcode <> ConvertHex("5A") Then
        ErrorFlag = True
        Exit Sub
    End If
   
    'This should be the hex code "E8 xx xx xx xx"
    If VBStartHeader.CallStartOpcode <> ConvertHex("E8") And VBStartHeader.CallStartOpcode <> ConvertHex("11") Then
        ErrorFlag = True
    End If

End Sub

Public Sub GetVBVer()

    'Fill in the VBSignature structure
    Mid$(VBSignature.VBVer$, 1) = Chr$(GetByte())
    Mid$(VBSignature.VBVer$, 2) = Chr$(GetByte())
    Mid$(VBSignature.VBVer$, 3) = Chr$(GetByte())
    Mid$(VBSignature.VBVer$, 4) = Chr$(GetByte())
    
    'The version should be "VB5!"
    If VBSignature.VBVer$ <> VBVERTEXT Then
        ErrorFlag = True
    End If
    
End Sub

Public Sub GetVBIntrptr()

    'Get VB interpreter name
    VBSignature.VBIntrptr$ = GetDosString()

    'The interpreter should be "MSVBVM60.DLL" or "MSVBVM50.DLL"
    If VBSignature.VBIntrptr$ <> VBINTRPTRTEXT Then
        ErrorFlag = True
    End If
    
End Sub

Public Function GetDWord() As Double
    
   GetDWord# = GetWord()
   GetDWord# = GetDWord# + 65536# * GetWord()
      
End Function

Public Function GetWord() As Double

    GetWord# = GetByte()
    GetWord# = GetWord# + 256# * GetByte()
       
End Function

Public Function GetByte() As Byte

    Dim DataByte As Byte
    
    'Read the data
    Get #InFileNumber, , DataByte
    
    'Return it
    GetByte = DataByte
      
End Function


Public Function ConvertHex(HexData As String) As Double

    Dim count As Integer
    
    'Get HexData as a 8 byte leading zero string
    HexData$ = "0000000" & HexData$
    HexData$ = Right$(HexData$, 8)
    
    'Sum up powers of 16 for 8 byte values
    ConvertHex = 0#
    For count = 1 To 8
        ConvertHex = ConvertHex + 16 ^ (count - 1) * (HexToDec(Mid$(HexData$, 9 - count, 1)))
    Next
    
End Function


Public Function GetDosString() As String

    'Clear string
    GetDosString$ = ""
     
    'Get a Dos string char by char
    Do
        'Get a char
        GetDosString$ = GetDosString$ & Chr$(GetByte())
        
    'Continue until we get a "0"
    Loop Until (Asc(Right$(GetDosString$, 1)) = 0)
        
    'Remove trailing zero
    GetDosString$ = Left$(GetDosString$, Len(GetDosString$) - 1)
  
End Function

Public Function HexToDec(HexDigit As String) As Double
    'Simple brute force hex-to-decimal conversion
    If HexDigit = "F" Then
        HexToDec# = 15#
    ElseIf HexDigit = "E" Then
        HexToDec# = 14#
    ElseIf HexDigit = "D" Then
        HexToDec# = 13#
    ElseIf HexDigit = "C" Then
        HexToDec# = 12#
    ElseIf HexDigit = "B" Then
        HexToDec# = 11#
    ElseIf HexDigit = "A" Then
        HexToDec# = 10#
    Else
        HexToDec# = Val(HexDigit$)
    End If
    
End Function

Private Sub ScanTable(fp As Integer, ByVal OffsetADR As Long, ByVal OffsetSTR As Long, ByRef outADRarray() As IMPORT_API_LOOKUP)
'*****************************
'Purpose: Used for processing the import table
'*****************************
Dim l As Long, i As Long, s As Long

    
    i = UBound(outADRarray()) - 1

    Get #fp, OffsetADR, l
    Do
       
        i = i + 1
        ReDim Preserve outADRarray(1 To i)
        outADRarray(i).Address = l
        Get #fp, OffsetSTR, s
        If (s And &H80000000) = 0 Then
            'Import by Name
            outADRarray(i).ApiName = ScanString(fp, GetPtrFromRVA2(s) + 3)
            outADRarray(i).InternalAddr = OffsetADR
        Else
            'Import by ordinal
            
            outADRarray(i).ApiName = "!ordinal : " & (s And &H7FFFFFFF)
           
        End If
     ' MsgBox OffsetADR & " " & OffsetSTR & " " & outADRarray(i).ApiName

        OffsetSTR = OffsetSTR + 4
        OffsetADR = OffsetADR + 4
        Get #fp, OffsetADR, l
    Loop Until l = 0

End Sub
Private Function ScanString(fp As Integer, ByVal offset As Long) As String
'*****************************
'Purpose: Used for processing the import table
'*****************************
    Dim b As Byte
    Get #fp, offset, b
    Do
        ScanString = ScanString & Chr$(b)
        offset = offset + 1
        Get #fp, offset, b
    Loop Until b = 0

End Function

Public Function GetPtrFromRVA(ByVal iRVA As Integer) As Long
'*****************************
'Purpose: To get the real entrypoint used for VB5
'*****************************
      Dim num2 As Integer
      Dim num3 As Integer
      num3 = PEHeader.NumSections - 1
      num2 = 0
      Do While (num2 <= num3)
            If ((iRVA >= SecHeader(num2).Address) And (iRVA < (SecHeader(num2).Address + SecHeader(num2).SizeRawData))) Then
                  GetPtrFromRVA = (iRVA - (SecHeader(num2).Address - SecHeader(num2).RawDataPointer))
                  Exit Function
            End If
            num2 = num2 + 1
      Loop
      GetPtrFromRVA = iRVA
End Function
Public Function GetPtrFromRVA2(ByVal iRVA As Long) As Long
'*****************************
'Purpose: To get the real entrypoint used for VB5
'*****************************
      Dim num2 As Integer
      Dim num3 As Integer
      num3 = PEHeader.NumSections - 1
      num2 = 0
      Do While (num2 <= num3)
            If ((iRVA >= SecHeader(num2).Address) And (iRVA < (SecHeader(num2).Address + SecHeader(num2).SizeRawData))) Then
                  GetPtrFromRVA2 = (iRVA - (SecHeader(num2).Address - SecHeader(num2).RawDataPointer))
                  Exit Function
            End If
            num2 = num2 + 1
      Loop
      GetPtrFromRVA2 = iRVA
End Function
Function ReturnVBVersionNe(iFileNumber As Integer) As Integer
    Dim lNum1 As Long
    Dim lNum2 As Long
    Dim iNum1 As Integer
    Dim i As Integer
    Dim strNameArray() As String
    Dim bSize As Byte
  lNum2 = DosHeader.ExeHeaderPointer + NEHeader.ModuleReferenceTableOffset + 1
  lNum1 = DosHeader.ExeHeaderPointer + NEHeader.ImportedNamesTableOffset + 1
  ReDim strNameArray(NEHeader.NumberEntriesModuleReferenceTable)
  For i = 1 To NEHeader.NumberEntriesModuleReferenceTable
    Get iFileNumber, lNum2, iNum1: lNum2 = lNum2 + 2
    Get iFileNumber, lNum1 + iNum1, bSize
    strNameArray(i) = Space$(Asc(bSize))
    Get iFileNumber, , strNameArray(i)
    If Left$(strNameArray(i), 5) = "VBRUN" Then
      ReturnVBVersionNe = Val(Mid$(strNameArray(i), 6, 1)) ' VB Version
    ElseIf Left$(strNameArray(i), 5) = "VB400" Then
      ReturnVBVersionNe = 4
    End If
  Next i
End Function

