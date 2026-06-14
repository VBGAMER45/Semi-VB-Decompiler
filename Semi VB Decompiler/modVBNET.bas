Attribute VB_Name = "modVBNET"
'*********************************************
'modVBNet
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************
'Credits
'Parts Translated to VB6 code from CorHdr.h from Microsoft
Option Explicit
'.NET version 2.0 Common Runtime
Private Type VBNETCORE2
    CB As Long 'DWORD
    MajorRuntimeVersion As Integer 'WORD
    MinorRuntimeVersion As Integer 'WORD
    MetaData As IMAGE_DATA_DIRECTORY
    flags As Long 'DWORD
    EntryPointToken As Long 'DWORD
    Resources As IMAGE_DATA_DIRECTORY
    StrongNameSignature As IMAGE_DATA_DIRECTORY
    CodeManagerTable As IMAGE_DATA_DIRECTORY
    VTableFixups As IMAGE_DATA_DIRECTORY
    ExportAddressTableJumps As IMAGE_DATA_DIRECTORY
    ManagedNativeHeader As IMAGE_DATA_DIRECTORY
End Type

Private Type MetadataStorageSignatureType
    lSignature As Double 'DWORD  �Magic� signature for physical metadata, currently 0x424A5342   - BSJB
    iMajorVersion As Double 'WORD Major version (1 for the first release of the common language runtime)
    iMinorVersion As Double 'WORD Minor version (1 for the first release of the common language runtime)
    iExtraData As Double 'DWORD Reserved; set to 0
    iLength As Double 'DWORD  Length of the version string
    iVersionString As String 'BYTE[?] Version string

End Type
Private Type MetaDataStorageHeaderType
    fFlags As Double 'BYTE
    Padding As Double 'BYTE
    iStreams As Double 'WORD
End Type
'Array of stream headers
Private Type StreamHeaderType
    iOffset As Double 'DWORD Offset in the file for this stream
    iSize As Double 'DWORD Size of the stream in bytes
    rcName As String 'NTZ Name of the stream; a zero-terminated ANSI string no longer than seven characters
End Type
Global gVBNetStreamHeaders() As StreamHeaderType
Private Type MetaDataStreamsHeaderType
    Reserved1 As Double 'BYTE
    Reserved2 As Double 'BYTE
    Reserved3 As Double 'BYTE
    Reserved4 As Double 'BYTE
    major As Double 'BYTE
    minor As Double 'BYTE
    Heaps As Double 'BYTE
    Rid As Double 'BYTE
    MaskValid(7) As Byte
    Sorted(7) As Byte
End Type
Dim MetaDataStreamsHeader As MetaDataStreamsHeaderType

'CLR Header entry point flags.
Private Const COMIMAGE_FLAGS_ILONLY = &H1
Private Const COMIMAGE_FLAGS_32BITREQUIRED = &H2
Private Const COMIMAGE_FLAGS_IL_LIBRARY = &H4
Private Const COMIMAGE_FLAGS_TRACKDEBUGDATA = &H10000
'Version flags for image
Private Const COR_VERSION_MAJOR = 2
Private Const COR_VERSION_MINOR = 0
Private Const COR_DELETED_NAME_LENGTH = 8
Private Const COR_VTABLEGAP_NAME_LENGTH = 8
'Maximum size of a NativeType descriptor.
Private Const NATIVE_TYPE_MAX_CB = 8
Private Const COR_ILMETHOD_SECT_SMALL_MAX_DATASIZE = &HFF
'MIH FLAGS
Private Const IMAGE_COR_MIH_METHODRVA = &H1
Private Const IMAGE_COR_MIH_EHRVA = &H2
Private Const IMAGE_COR_MIH_BASICBLOCK = &H8
' V-table constants
Private Const COR_VTABLE_32BIT = &H1  'V-table slots are 32-bits in size
Private Const COR_VTABLE_64BIT = &H2  'V-table slots are 64-bits in size.
Private Const COR_VTABLE_FROM_UNMANAGED = &H4 'If set, transition from unmanaged
Private Const COR_VTABLE_FROM_UNMANAGED_RETAIN_APPDOMAIN = &H8  ' If set, transition from unmanaged with keeping the current appdomain.
Private Const COR_VTABLE_CALL_MOST_DERIVED = &H10      'Call most derived method described by
'EATJ constants
Private Const IMAGE_COR_EATJ_THUNK_SIZE = 32         ' Size of a jump thunk reserved range.
'Max name lengths
Private Const MAX_CLASS_NAME = 1024
Private Const MAX_PACKAGE_NAME = 1024

'TypeDef/ExportedType attr bits, used by DefineTypeDef.
Private Enum CorTypeAttr
   'Use this mask to retrieve the type visibility information.
    tdVisibilityMask = &H7
    tdNotPublic = &H0                     'Class is not public scope.
    tdPublic = &H1                        'Class is public scope.
    tdNestedPublic = &H2                  'Class is nested with public visibility.
    tdNestedPrivate = &H3                 'Class is nested with private visibility.
    tdNestedFamily = &H4                  'Class is nested with family visibility.
    tdNestedAssembly = &H5                'Class is nested with assembly visibility.
    tdNestedFamANDAssem = &H6             'Class is nested with family and assembly visibility.
    tdNestedFamORAssem = &H7              'Class is nested with family or assembly visibility.
End Enum

Dim strTableNames(44) As String '"Module" , "TypeRef" , "TypeDef" ,"FieldPtr","Field", "MethodPtr","Method","ParamPtr" , "Param", "InterfaceImpl", "MemberRef", "Constant", "CustomAttribute", "FieldMarshal", "DeclSecurity", "ClassLayout", "FieldLayout", "StandAloneSig" , "EventMap","EventPtr", "Event", "PropertyMap", "PropertyPtr", "Properties","MethodSemantics","MethodImpl","ModuleRef","TypeSpec","ImplMap","FieldRVA","ENCLog","ENCMap","Assembly","AssemblyProcessor","AssemblyOS","AssemblyRef","AssemblyRefProcessor","AssemblyRefOS","File","ExportedType","ManifestResource","NestedClass","TypeTyPar","MethodTyPar"

Global gVBNETHeader As VBNETCORE2
Global gVBNETMetaData As MetadataStorageSignatureType
Dim gVBNETMetaDataHeader As MetaDataStorageHeaderType

Dim Console As clsConsole

Global bISVBNET As Boolean

'Hold the stream information
Dim MetaDataByteArray() As Byte
Dim StringsByteArray() As Byte
Dim USByteArray() As Byte
Dim GuidByteArray() As Byte
Dim BlobByteArray() As Byte

Dim OffsetString As Long
Dim OffsetBlob As Long
Dim OffsetGuid As Long
Dim Valid As Currency
Dim bDebug As Boolean
Dim iRows() As Long
Dim TableOffset As Long
Dim Sizes() As Long
'Begin Program6.csc Structs
Private Type FieldPtrTable
    index As Long
End Type
Private Type MethodPtrTable
    index As Long
End Type
Private Type ExportedTypeTable
    flags As Long
    typedefindex  As Long
    name As Long
    nspace As Long
    coded As Long
End Type
Private Type NestedClassTable
    nestedclass As Long
    enclosingclass As Long
End Type
Private Type MethodImpTable
    classindex As Long
    codedbody As Long
    codeddef As Long
End Type
Private Type ClassLayoutTable
     packingsize As Integer
     classsize As Long
     parent As Long
End Type
Private Type ManifestResourceTable
    offset As Long
    flags As Long
    name As Long
    coded As Long
End Type
Private Type ModuleRefTable
    name As Long
End Type
Private Type FileTable
    flags As Long
    name As Long
    index As Long
End Type
Private Type EventTable
    attr As Integer
    name As Long
    coded As Long
End Type
Private Type EventMapTable
    index As Long
    eindex As Long
End Type
Private Type MethodSemanticsTable
    methodsemanticsattributes As Integer
    methodindex As Long
    association As Long
End Type
Private Type PropertyMapTable
    parent As Long
    propertylist As Long
End Type
Private Type PropertyTable
    flags As Long
    name As Long
    type As Long
End Type
Private Type ConstantsTable
    dtype As Integer
    parent As Long
    value As Long
End Type
Private Type FieldLayoutTable
    offset As Long
    fieldindex As Long
End Type
Private Type FieldRVATable
    rva  As Long
    fieldi As Long
End Type
Private Type FieldMarshalTable
     coded As Long
     index As Long
End Type
Private Type FieldTable
    flags As Long
    name As Long
    sig As Long
End Type
Private Type ParamTable
    pAttr As Integer
    sequence As Long
    name As Long
End Type
Private Type TypeSpecTable
     signature As Long
End Type
Private Type MemberRefTable
    clas As Long
    name As Long
    sig As Long
End Type
Private Type StandAloneSigTable
    index As Long
End Type
Private Type InterfaceImplTable
    classindex As Long
    interfaceindex As Long
End Type
Private Type TypeDefTable
    flags As Long
    name As Long
    nspace As Long
    cindex As Long
    findex As Long
    mindex As Long
End Type
Private Type CustomAttributeTable
    parent As Long
    type As Long
    value As Long
End Type
Private Type AssemblyRefTable
    major As Integer
    minor As Integer
    build As Integer
    revision As Integer
    flags As Long
    publickey As Long
    name As Long
    culture As Long
    hashvalue As Long
End Type
Private Type AssemblyTable
    HashAlgId As Long
    major As Long
    minor As Long
    build As Long
    revision As Long
    flags As Long
    publickey As Long
    name As Long
    culture As Long
End Type
Private Type ModuleTable
    Generation As Long
    name As Long
    Mvid As Long
    EncId As Long
    EncBaseId As Long
End Type
Private Type TypeRefTable
    resolutionscope As Long
    name As Long
    nspace As Long
End Type
Private Type MethodTable
    rva As Long
    impflags As Long
    flags As Long
    name As Long
    signature As Long
    param As Long
End Type
Private Type DeclSecurityTable
    action As Long
    coded As Long
    bindex As Long
End Type
Private Type ImplMapTable
    attr As Integer
    cindex As Long
    name As Long
    scope As Long
End Type
Dim AssemblyStruct() As AssemblyTable
Dim AssemblyRefStruct() As AssemblyRefTable
Dim CustomAttributeStruct() As CustomAttributeTable
Dim ModuleStruct() As ModuleTable
Dim TypeDefStruct() As TypeDefTable
Dim TypeRefStruct() As TypeRefTable
Dim InterfaceImplStruct() As InterfaceImplTable
Dim FieldPtrStruct() As FieldPtrTable
Dim MethodPtrStruct() As MethodPtrTable
Dim MethodStruct() As MethodTable
Dim StandAloneSigStruct() As StandAloneSigTable
Dim MemberRefStruct() As MemberRefTable
Dim TypeSpecStruct() As TypeSpecTable
Dim ParamStruct() As ParamTable
Dim FieldStruct() As FieldTable
Dim FieldMarshalStruct() As FieldMarshalTable
Dim FieldRVAStruct() As FieldRVATable
Dim FieldLayoutStruct() As FieldLayoutTable
Dim ConstantsStruct() As ConstantsTable
Dim PropertyMapStruct() As PropertyMapTable
Dim PropertyStruct() As PropertyTable
Dim MethodSemanticsStruct() As MethodSemanticsTable
Dim EventStruct() As EventTable
Dim EventMapStruct() As EventMapTable
Dim FileStruct() As FileTable
Dim ModuleRefStruct() As ModuleRefTable
Dim ManifestResourceStruct() As ManifestResourceTable
Dim ClassLayoutStruct() As ClassLayoutTable
Dim MethodImpStruct() As MethodImpTable
Dim NestedClassStruct() As NestedClassTable
Dim ExportedTypeStruct() As ExportedTypeTable
Dim DeclSecurityStruct() As DeclSecurityTable
Dim ImplMapStruct() As ImplMapTable
'end program6.csc

'Program10.csc
Dim spacesforrest As Long
Dim spacesfornested As Long
Dim spacefornamespace  As Long

'
Dim tinyformat As Boolean

Dim First12 As Long

Dim Spacesfortry As Long

Dim Placedend  As Boolean

Dim Notprototype As Boolean
Dim Writenamespace As Boolean
Dim lasttypedisplayed As Long

'*** modVBNET decompiler (IL + C#/VB.NET reconstruction) state ***
Private Const LANG_IL As Integer = 0
Private Const LANG_CS As Integer = 1
Private Const LANG_VB As Integer = 2
'Operand kinds for the IL opcode table
Private Const OK_NONE As Integer = 0
Private Const OK_I1 As Integer = 1
Private Const OK_U1 As Integer = 2
Private Const OK_VAR As Integer = 3
Private Const OK_I4 As Integer = 4
Private Const OK_I8 As Integer = 5
Private Const OK_R4 As Integer = 6
Private Const OK_R8 As Integer = 7
Private Const OK_BR1 As Integer = 8
Private Const OK_BR4 As Integer = 9
Private Const OK_TOK As Integer = 10
Private Const OK_STR As Integer = 11
Private Const OK_SWITCH As Integer = 12
'Decoded instruction stream for the current method body
Private gInsCount As Long
Private gInsPos() As Long
Private gInsName() As String
Private gInsText() As String
Private gInsVal() As Long
'Per-type captured output, used by the project tree and the solution builder
Private gNetTypeCount As Long
Private gNetTypeName() As String
Private gNetTypeCS() As String
Private gNetTypeVB() As String
Private gNetTypeIL() As String
Private gNetTypeMethods() As String
'Per-instruction switch target lists (csv of absolute IL offsets)
Private gSwitchTargets() As String
'Exception-handling clauses for the current method body
Private gEHCount As Long
Private gEHFlags() As Long
Private gEHTryOff() As Long
Private gEHTryLen() As Long
Private gEHHandOff() As Long
Private gEHHandLen() As Long
Private gEHToken() As Long
'Structured statement list for control-flow reconstruction
Private Const LT_STMT As Integer = 0
Private Const LT_CBR As Integer = 1
Private Const LT_BR As Integer = 2
Private Const LT_RET As Integer = 3
Private Const LT_THROW As Integer = 4
Private Const LT_SWITCH As Integer = 5
Private Const LT_LABEL As Integer = 6
Private Const LT_RAW As Integer = 7
Private Const LT_OPEN As Integer = 8
Private Const LT_MID As Integer = 9
Private Const LT_CLOSE As Integer = 10
Private gLCount As Long
Private gLType() As Integer
Private gLText() As String
Private gLAlt() As String
Private gLTarget() As Long
Private gLSwitch() As String
Private gLOffset() As Long
Private gLDead() As Boolean
'Render context for the current language
Private gLang As Integer
Private gTerm As String, gThisKw As String, gNullKw As String
Private gNewKw As String, gRetKw As String, gThrowKw As String, gEqOp As String
Private gCurOff As Long

'SemiVBDecompilerHelper.dll (.NET) is no longer required - its BitConverter /
'bit-shift helpers are implemented in pure VB6 further down this module.
'Used to check if .Net is installed
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Sub InitDotNet()
'"Module" , "TypeRef" , "TypeDef" ,"FieldPtr","Field", "MethodPtr","Method","ParamPtr" , "Param", "InterfaceImpl", "MemberRef", "Constant", "CustomAttribute", "FieldMarshal", "DeclSecurity", "ClassLayout", "FieldLayout", "StandAloneSig" , "EventMap","EventPtr", "Event", "PropertyMap", "PropertyPtr", "Properties","MethodSemantics","MethodImpl","ModuleRef","TypeSpec","ImplMap","FieldRVA","ENCLog","ENCMap","Assembly","AssemblyProcessor","AssemblyOS","AssemblyRef","AssemblyRefProcessor","AssemblyRefOS","File","ExportedType","ManifestResource","NestedClass","TypeTyPar","MethodTyPar"
    strTableNames(0) = "Module"
    strTableNames(1) = "TypeRef"
    strTableNames(2) = "TypeDef"
    strTableNames(3) = "FieldPtr"
    strTableNames(4) = "Field"
    strTableNames(5) = "MethodPtr"
    strTableNames(6) = "Method"
    strTableNames(7) = "ParamPtr"
    strTableNames(8) = "Param"
    strTableNames(9) = "InterfaceImpl"
    strTableNames(10) = "MemberRef"
    strTableNames(11) = "Constant"
    strTableNames(12) = "CustomAttribute"
    strTableNames(13) = "FieldMarshal"
    strTableNames(14) = "DeclSecurity"
    strTableNames(15) = "ClassLayout"
    strTableNames(16) = "FieldLayout"
    strTableNames(17) = "StandAloneSig"
    strTableNames(18) = "EventMap"
    strTableNames(19) = "EventPtr"
    strTableNames(20) = "Event"
    strTableNames(21) = "PropertyMap"
    strTableNames(22) = "PropertyPtr"
    strTableNames(23) = "Properties"
    strTableNames(24) = "MethodSemantics"
    strTableNames(25) = "MethodImpl"
    strTableNames(26) = "ModuleRef"
    strTableNames(27) = "TypeSpec"
    strTableNames(28) = "ImplMap"
    strTableNames(29) = "FieldRVA"
    strTableNames(30) = "ENCLog"
    strTableNames(31) = "ENCMap"
    strTableNames(32) = "Assembly"
    strTableNames(33) = "AssemblyProcessor"
    strTableNames(34) = "AssemblyOS"
    strTableNames(35) = "AssemblyRef"
    strTableNames(36) = "AssemblyRefProcessor"
    strTableNames(37) = "AssemblyRefOS"
    strTableNames(38) = "File"
    strTableNames(39) = "ExportedType"
    strTableNames(40) = "ManifestResource"
    strTableNames(41) = "NestedClass"
    strTableNames(42) = "TypeTyPar"
    strTableNames(43) = "MethodTyPar"

    
    'Setup Offest strings
    OffsetString = 2
    OffsetBlob = 2
    OffsetGuid = 2
    
    bDebug = True
    'program 10
    spacesforrest = 2
    '
    Placedend = False
    Notprototype = False
End Sub
Public Function GetDWordByFile(FileNum As Integer) As Double
    
   GetDWordByFile# = GetWordByFile(FileNum)
   GetDWordByFile# = GetDWordByFile# + 65536# * GetWordByFile(FileNum)
      
End Function

Public Function GetWordByFile(FileNum As Integer) As Double

    GetWordByFile# = GetByteByFile(FileNum)
    GetWordByFile# = GetWordByFile# + 256# * GetByteByFile(FileNum)
       
End Function

Public Function GetByteByFile(FileNum As Integer) As Byte
    Dim DataByte As Byte
    
    'Read the data
    Get #FileNum, , DataByte
    
    'Return it
    GetByteByFile = DataByte
      
End Function
Sub ProccessVBNETFile(lOffsetVBNETHEADER As Long, FileNum As Integer)
    Dim F As Integer
    F = FileNum
    
    If IsDotNetInstalled = True Then
        'MsgBox ".Net Runtime is installed"
    Else
        If gQuietMode = False Then MsgBox ".Net Runtime is not installed on this machine. Please download and install it.  To Process .Net files", vbCritical
        Exit Sub
    End If
    
    'tablenames() ="Module" , "TypeRef" , "TypeDef" ,"FieldPtr","Field", "MethodPtr","Method","ParamPtr" , "Param", "InterfaceImpl", "MemberRef", "Constant", "CustomAttribute", "FieldMarshal", "DeclSecurity", "ClassLayout", "FieldLayout", "StandAloneSig" , "EventMap","EventPtr", "Event", "PropertyMap", "PropertyPtr", "Properties","MethodSemantics","MethodImpl","ModuleRef","TypeSpec","ImplMap","FieldRVA","ENCLog","ENCMap","Assembly","AssemblyProcessor","AssemblyOS","AssemblyRef","AssemblyRefProcessor","AssemblyRefOS","File","ExportedType","ManifestResource","NestedClass","TypeTyPar","MethodTyPar"
    Set Console = New clsConsole
    Console.Clear
    Console.WriteLine ("***************************************")
    Console.WriteLine ("Semi VB Decompiler - IL Disassembler")
    Console.WriteLine ("http://www.visualbasiczone.com")
    Console.WriteLine ("***************************************")
    Console.WriteLine ("")
    'Load .Net Data
    Call InitDotNet
    
    'Get CLR Header
    Get F, lOffsetVBNETHEADER, gVBNETHeader
    
    'Get MetaData Infomation
    Seek F, GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) + 1
    
    Call GetMetaData(F)
    Call FillTableSizes
    Call ReadTablesIntoStructures
    Call DisplayTablesForDebugging
    'Call ReadandDisplayVTableFixup
    Call ReadandDisplayExportAddressTableJumps
    Call DisplayModuleRefs
    'Call DisplayAssembleyRefs
    'Call DisplayAssembley
    Call DisplayFileTable
    'Call DisplayClassExtern
    Call DisplayResources
    Call DisplayModuleAndMore
    Call DisplayTypeDefs
    Call DisplayTypeDefsAndMethods

    'Full decompilation pass: IL disassembly + best-effort C#/VB.NET source.
    Call DecompileDotNet

    Console.SaveConsoleToFile App.Path & "\dump\" & SFile & "\netconsole.txt"
    'Make the console report menu visible
    frmMain.mnuToolsNetConsole.Visible = True
    'Enable the .NET solution export now that classes have been reconstructed
    frmMain.mnuFileBuildSolution.Enabled = True
    'Add to recent files
    Call frmMain.AddToRecentList(SFilePath, SFile)
End Sub
Private Sub GetMetaData(FileNum As Integer)
    Dim lMetaDataStartPos As Long
    lMetaDataStartPos = Loc(FileNum)
     
    gVBNETMetaData.lSignature = GetDWordByFile(FileNum)
    gVBNETMetaData.iMajorVersion = GetWordByFile(FileNum)
    gVBNETMetaData.iMinorVersion = GetWordByFile(FileNum)
    gVBNETMetaData.iExtraData = GetDWordByFile(FileNum)
    gVBNETMetaData.iLength = GetDWordByFile(FileNum)
    Dim bArray() As Byte
    ReDim bArray(gVBNETMetaData.iLength - 1)
    Get FileNum, , bArray
    Dim strBuffer As String
    Dim i As Integer
    For i = 0 To UBound(bArray)
        strBuffer = strBuffer & Chr$(bArray(i))
    Next
    gVBNETMetaData.iVersionString = strBuffer
    'Begin Getting MetaData Storage Header
    gVBNETMetaDataHeader.fFlags = GetByteByFile(FileNum)
    gVBNETMetaDataHeader.Padding = GetByteByFile(FileNum)
    gVBNETMetaDataHeader.iStreams = GetWordByFile(FileNum)
    Dim g As Integer
    Dim bByte As Byte
    ReDim gVBNetStreamHeaders(gVBNETMetaDataHeader.iStreams - 1)
    For i = 0 To gVBNETMetaDataHeader.iStreams - 1
        gVBNetStreamHeaders(i).iOffset = GetDWordByFile(FileNum)
        gVBNetStreamHeaders(i).iSize = GetDWordByFile(FileNum)
        gVBNetStreamHeaders(i).rcName = GetUntilNull(FileNum)

        Do While True
            If (Loc(FileNum) Mod 4) = 0 Then Exit Do
            Get FileNum, , bByte
        Loop
        
    Next
   

            
    'Get Metadata Table Streams Header
    For i = 0 To UBound(gVBNetStreamHeaders)
        'The metadata streams #~ and #- begin with the following header:
        If gVBNetStreamHeaders(i).rcName = "#~" Or gVBNetStreamHeaders(i).rcName = "#-" Then
            Seek #FileNum, lMetaDataStartPos + gVBNetStreamHeaders(i).iOffset + 1
            
            MetaDataStreamsHeader.Reserved1 = GetByteByFile(FileNum)
            MetaDataStreamsHeader.Reserved2 = GetByteByFile(FileNum)
            MetaDataStreamsHeader.Reserved4 = GetByteByFile(FileNum)
            MetaDataStreamsHeader.Reserved4 = GetByteByFile(FileNum)
            MetaDataStreamsHeader.major = GetByteByFile(FileNum)
            MetaDataStreamsHeader.minor = GetByteByFile(FileNum)
            MetaDataStreamsHeader.Heaps = GetByteByFile(FileNum)
            'Save MetaData to an array
            Seek #FileNum, lMetaDataStartPos + gVBNetStreamHeaders(i).iOffset + 1
            ReDim MetaDataByteArray(gVBNetStreamHeaders(i).iSize)
            Get #FileNum, Loc(FileNum) + 1, MetaDataByteArray
            
            'MsgBox MetaDataStreamsHeader.Major
            'MsgBox Loc(FileNum)
        End If
        '#Strings
        If gVBNetStreamHeaders(i).rcName = "#Strings" Then
            Seek #FileNum, lMetaDataStartPos + gVBNetStreamHeaders(i).iOffset + 1
            ReDim StringsByteArray(gVBNetStreamHeaders(i).iSize)
            Get #FileNum, , StringsByteArray
          
            
        End If
        '#US
        If gVBNetStreamHeaders(i).rcName = "#US" Then
            Seek #FileNum, lMetaDataStartPos + gVBNetStreamHeaders(i).iOffset + 1
            ReDim USByteArray(gVBNetStreamHeaders(i).iSize)
            Get #FileNum, , USByteArray
        End If
        '#GUID
        If gVBNetStreamHeaders(i).rcName = "#GUID" Then
            Seek #FileNum, lMetaDataStartPos + gVBNetStreamHeaders(i).iOffset + 1
            ReDim GuidByteArray(gVBNetStreamHeaders(i).iSize)
            Get #FileNum, , GuidByteArray
        End If
        '#BLOB
        If gVBNetStreamHeaders(i).rcName = "#Blob" Then
            Seek #FileNum, lMetaDataStartPos + gVBNetStreamHeaders(i).iOffset + 1
            ReDim BlobByteArray(gVBNetStreamHeaders(i).iSize)
            Get #FileNum, , BlobByteArray
        End If
        
    Next
    'Do Offset functions for heaps
    If (MetaDataStreamsHeader.Heaps And &H1) = &H1 Then
        OffsetString = 4
    End If
    If (MetaDataStreamsHeader.Heaps And &H2) = &H2 Then
        OffsetGuid = 4
    End If
    If (MetaDataStreamsHeader.Heaps And &H4) = &H4 Then
        OffsetBlob = 4
    End If
    
    TableOffset = 24
    'valid = BitConverter.ToInt64 (metadata, 8);
    '#TODO
    'Valid = modVBNET.BitConverterToInt64(MetaDataByteArray, 8)
    Valid = BitConverterToInt64(MetaDataByteArray, 8)
    'Debug.Print "VAILD: " & Valid
    ReDim iRows(64)
    
    Dim k As Long
    For k = 0 To 63
        Dim lTablepresent As Long
        '#DONE
        'int tablepresent = (int)(valid >> k ) & 1;
        ''lTablepresent = HiLo.DWordShiftR(Valid, k) And 1
        lTablepresent = isTablePresent(MetaDataByteArray, k)
        'lTablepresent = HiLo.INT64ShiftR(Valid, k) And 1
        If lTablepresent = 1 Then
            '
            '#DONE
            'rows[k] = BitConverter.ToInt32(metadata , tableoffset);
            'iRows(k) = modVBNET.BitConverterToInt32(MetaDataByteArray, TableOffset)
            iRows(k) = BitConverterToInt32(MetaDataByteArray, TableOffset)
            'MsgBox strTableNames(k) & " R: " & iRows(k)
            Console.WriteLine (strTableNames(k) & " " & iRows(k))
            'MsgBox iRows(k)
            TableOffset = TableOffset + 4
        End If
    Next
    Console.WriteLine ("")

End Sub
Public Sub ShowStrings(fxgEXEInfo As MSFlexGrid)
    Dim F As Integer
    Dim i As Integer
    Dim Counter As Long
    Dim lOldLoc As Long
    Counter = 1
    F = FreeFile
    Dim strBuffer As String
    Open SFilePath For Binary Access Read As #F
    For i = 0 To UBound(gVBNetStreamHeaders)
        If gVBNetStreamHeaders(i).rcName = "#Strings" Then
            Seek F, GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) + gVBNetStreamHeaders(i).iOffset + 2
            Do
                lOldLoc = Loc(F)
                strBuffer = GetUntilNull(F)
                'MsgBox strBuffer & " Loc: " & Loc(F)
                'If strBuffer <> "" Then
                   If Counter = 1 Then
                        fxgEXEInfo.TextArray(2) = lOldLoc - GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) - gVBNetStreamHeaders(i).iOffset
                        fxgEXEInfo.TextArray(3) = strBuffer
                        Counter = Counter + 4
                   
                   Else
                    fxgEXEInfo.AddItem lOldLoc - GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) - gVBNetStreamHeaders(i).iOffset
                    fxgEXEInfo.TextArray(Counter) = strBuffer
                      Counter = Counter + 2
                   End If
                'End If
            Loop Until strBuffer = ""
        End If
    Next
    Close #F
End Sub
Public Sub ShowUserStrings(fxgEXEInfo As MSFlexGrid)
    Dim F As Integer
    Dim i As Integer
    Dim Counter As Long
    Dim lOldLoc As Long
    Dim length As Byte
    Counter = 1
    F = FreeFile
    Dim strBuffer As String
    Dim lFinalSize As Long
    Open SFilePath For Binary Access Read As #F
  
    For i = 0 To UBound(gVBNetStreamHeaders)
        If gVBNetStreamHeaders(i).rcName = "#US" Then
            lFinalSize = GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) + gVBNetStreamHeaders(i).iOffset + gVBNetStreamHeaders(i).iSize + 2
            Seek F, GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) + gVBNetStreamHeaders(i).iOffset + 2
            Do
                lOldLoc = Loc(F)
                'strBuffer = GetUntilNull(F)
                length = GetByte2(F)
               ' MsgBox "Length: " & length
                Seek F, Loc(F)
                length = length - 1
                length = length / 2
                strBuffer = GetUnicodeString(F, CInt(length))
                'MsgBox strBuffer & " Loc: " & Loc(F)
                   If Counter = 1 Then
                        fxgEXEInfo.TextArray(2) = lOldLoc - GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) - gVBNetStreamHeaders(i).iOffset
                        fxgEXEInfo.TextArray(3) = strBuffer
                        Counter = Counter + 4
                   Else
                    fxgEXEInfo.AddItem lOldLoc - GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) - gVBNetStreamHeaders(i).iOffset
                    fxgEXEInfo.TextArray(Counter) = strBuffer
                      Counter = Counter + 2
                   End If
                length = GetByte2(F)
                
            'Loop Until strBuffer = ""
            
            Loop Until Loc(F) >= lFinalSize
        End If
    Next
    Close #F
End Sub
Public Sub ShowGUIDHeap(fxgEXEInfo As MSFlexGrid)
    Dim F As Integer
    Dim i As Integer
    Dim Counter As Long
    Dim lOldLoc As Long
    Counter = 1
    F = FreeFile
    Dim strBuffer As String
    Dim lFinalSize As Long
    Open SFilePath For Binary Access Read As #F
  
    For i = 0 To UBound(gVBNetStreamHeaders)
        If gVBNetStreamHeaders(i).rcName = "#GUID" Then
            lFinalSize = GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) + gVBNetStreamHeaders(i).iOffset + gVBNetStreamHeaders(i).iSize '+ 1
            Seek F, GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) + gVBNetStreamHeaders(i).iOffset + 1
            Do
                lOldLoc = Loc(F)
                strBuffer = ReturnGuid(F)
      
                   If Counter = 1 Then
                        fxgEXEInfo.TextArray(2) = lOldLoc - GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) - gVBNetStreamHeaders(i).iOffset
                        fxgEXEInfo.TextArray(3) = strBuffer
                        Counter = Counter + 4
                   Else
                    fxgEXEInfo.AddItem lOldLoc - GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) - gVBNetStreamHeaders(i).iOffset
                    fxgEXEInfo.TextArray(Counter) = strBuffer
                      Counter = Counter + 2
                   End If
            
            Loop Until Loc(F) >= lFinalSize
        End If
    Next
    Close #F
End Sub
Public Sub ShowBlobHeap(fxgEXEInfo As MSFlexGrid)
    Dim F As Integer
    Dim i As Integer
    Dim Counter As Long
    Dim lOldLoc As Long
    Dim bArray() As Byte
    Dim length As Byte
    Counter = 1
    F = FreeFile
    Dim strBuffer As String
    Dim lFinalSize As Long
    Open SFilePath For Binary Access Read As #F
  
    For i = 0 To UBound(gVBNetStreamHeaders)
        If gVBNetStreamHeaders(i).rcName = "#Blob" Then
            lFinalSize = GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) + gVBNetStreamHeaders(i).iOffset + gVBNetStreamHeaders(i).iSize '+ 1
            Seek F, GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) + gVBNetStreamHeaders(i).iOffset + 2
            Do
                lOldLoc = Loc(F)
                Get #F, , length
                ReDim bArray(length - 1)
                Get #F, , bArray
                   If Counter = 1 Then
                        fxgEXEInfo.TextArray(2) = lOldLoc - GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) - gVBNetStreamHeaders(i).iOffset
                        fxgEXEInfo.TextArray(3) = length
                        Counter = Counter + 4
                   Else
                      fxgEXEInfo.AddItem lOldLoc - GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress) - gVBNetStreamHeaders(i).iOffset
                      fxgEXEInfo.TextArray(Counter) = length
                      Counter = Counter + 2
                   End If
            
            Loop Until Loc(F) >= lFinalSize
        End If
    Next
    Close #F
End Sub

'*****************************
'Pure VB6 replacements for the old SemiVBDecompilerHelper.dll (.NET) wrappers.
'These are just little-endian byte -> integer conversions, exactly what
'System.BitConverter did, so the .NET helper and its COM reference are no
'longer needed.
'*****************************

'Build a signed 32-bit Long from 4 little-endian bytes without overflowing
'(the high byte's top bit becomes the Long's sign bit).  This matches both
'.NET BitConverter.ToInt32 and ToUInt32 marshaled into a VB6 Long (same bits).
Private Function BytesToLong32(bArray() As Byte, ByVal offset As Long) As Long
    Dim v As Long
    v = bArray(offset) + bArray(offset + 1) * 256& + bArray(offset + 2) * 65536
    If bArray(offset + 3) < 128 Then
        v = v + bArray(offset + 3) * 16777216
    Else
        v = v + (CLng(bArray(offset + 3)) - 256) * 16777216
    End If
    BytesToLong32 = v
End Function

Public Function BitConverterToInt16(bArray() As Byte, offset As Long) As Long
    Dim v As Long
    v = bArray(offset) + bArray(offset + 1) * 256&
    If v >= 32768 Then v = v - 65536        'sign-extend
    BitConverterToInt16 = v
End Function
Public Function BitConverterToUInt16(bArray() As Byte, offset As Long) As Long
    BitConverterToUInt16 = bArray(offset) + bArray(offset + 1) * 256&
End Function
Public Function BitConverterToUInt32(bArray() As Byte, offset As Long) As Long
    BitConverterToUInt32 = BytesToLong32(bArray, offset)
End Function
Public Function BitConverterToInt32(bArray() As Byte, offset As Long) As Long
    BitConverterToInt32 = BytesToLong32(bArray, offset)
End Function
Public Function BitConverterToInt64(bArray() As Byte, offset As Long) As Currency
    'Combine two 32-bit halves into a Currency.  Currency comfortably holds the
    'values this decompiler reads (metadata offsets / the table-valid mask);
    'genuinely huge 64-bit values would overflow, so guard and return 0.
    On Error GoTo overflow
    Dim lo As Long, hi As Long
    Dim loC As Currency, hiC As Currency
    lo = BytesToLong32(bArray, offset)
    hi = BytesToLong32(bArray, offset + 4)
    loC = lo: If lo < 0 Then loC = loC + 4294967296@
    hiC = hi: If hi < 0 Then hiC = hiC + 4294967296@
    BitConverterToInt64 = loC + hiC * 4294967296@
    Exit Function
overflow:
    BitConverterToInt64 = 0
End Function

'Bit k of the 64-bit "Valid" table mask stored at metadata offset 8 - i.e.
'(BitConverterToInt64(metadata, 8) >> k) And 1.  Tested directly on the bytes
'to avoid any 64-bit arithmetic.
Public Function isTablePresent(metadata() As Byte, ByVal k As Long) As Long
    Dim byteIndex As Long, bitInByte As Long, b As Long
    byteIndex = 8 + (k \ 8)
    bitInByte = k Mod 8
    b = metadata(byteIndex)
    isTablePresent = (b \ (2 ^ bitInByte)) And 1
End Function

'Logical (unsigned) right shift, replacing the old .NET helper.
Public Function DoRightBitShift(ByVal value As Long, ByVal count As Long) As Long
    If count <= 0 Then DoRightBitShift = value: Exit Function
    If count >= 32 Then DoRightBitShift = 0: Exit Function
    If value >= 0 Then
        DoRightBitShift = value \ (2 ^ count)
    Else
        Dim c As Currency
        c = CCur(value) + 4294967296@           'treat as unsigned 32-bit
        DoRightBitShift = CLng(Int(c / (2 ^ count)))
    End If
End Function
Private Function GetCodedIndexSize(strTableName As String) As Long
    If strTableName = "Implementation" Then
        If iRows(&H26) > 16384 Or iRows(&H23) > 16384 Or iRows(&H27) > 16384 Then
            GetCodedIndexSize = 4
        Else
            GetCodedIndexSize = 2
        End If
    End If
    If strTableName = "MemberForwarded" Then
        If iRows(4) >= 32768 Or iRows(6) >= 32768 Then
            GetCodedIndexSize = 4
        Else
            GetCodedIndexSize = 2
        End If
    End If
    If strTableName = "MethodDefOrRef" Then
    'rows[0x06] >= 32768 || rows[0x0A] >= 32768)
        If iRows(6) >= 32768 Or iRows(&HA) >= 32768 Then
            GetCodedIndexSize = 4
        Else
            GetCodedIndexSize = 2
        End If
    End If
    If strTableName = "HasSemantics" Then
    'if ( rows[0x14] >= 32768 || rows[0x17] >= 32768)
        If iRows(&H14) >= 32768 Or iRows(&H17) >= 32768 Then
            GetCodedIndexSize = 4
        Else
            GetCodedIndexSize = 2
        End If
    End If
    If strTableName = "HasDeclSecurity" Then
        'if ( rows[0x02] >= 16384 || rows[0x06] >= 16384 || rows[0x20] >= 16384)
        If iRows(2) >= 16384 Or iRows(6) >= 16384 Or iRows(&H20) >= 16384 Then
            GetCodedIndexSize = 4
        Else
            GetCodedIndexSize = 2
        End If
    End If
    If strTableName = "HasFieldMarshal" Then
        'if ( rows[0x04] >= 32768|| rows[0x08] >= 32768)
        If iRows(4) >= 32768 Or iRows(8) >= 32768 Then
            GetCodedIndexSize = 4
        Else
            GetCodedIndexSize = 2
        End If
    End If
    If strTableName = "TypeDefOrRef" Then
        'if ( rows[0x02] >= 16384 || rows[0x01] >= 16384  || rows[0x1B] >= 16384   )
        If iRows(2) >= 16384 Or iRows(1) >= 16384 Or iRows(&H1B) >= 16384 Then
            GetCodedIndexSize = 4
        Else
            GetCodedIndexSize = 2
        End If
    End If
    If strTableName = "ResolutionScope" Then
    ' rows[0x00] >= 16384 || rows[0x1a] >= 16384  || rows[0x23] >= 16384  || rows[0x01] >= 16384 )
        If iRows(0) >= 16484 Or iRows(&H1A) >= 16384 Or iRows(&H23) >= 16384 Or iRows(1) >= 16384 Then
            GetCodedIndexSize = 4
        Else
            GetCodedIndexSize = 2
        End If
    End If
    If strTableName = "HasConst" Then
    '( rows[4] >= 16384 || rows[8] >= 16384 || rows[0x17] >= 16384 )
        If iRows(4) >= 16384 Or iRows(8) >= 16384 Or iRows(&H17) >= 16484 Then
            GetCodedIndexSize = 4
        Else
            GetCodedIndexSize = 2
        End If
    End If
    If strTableName = "MemberRefParent" Then
    ' rows[0x08] >= 8192 || rows[0x04] >= 8192 || rows[0x17] >= 8192  )
        If iRows(8) >= 8192 Or iRows(4) >= 8192 Or iRows(&H17) >= 8192 Then
            GetCodedIndexSize = 4
        Else
            GetCodedIndexSize = 2
        End If
    End If
    If strTableName = "HasCustomAttribute" Then
    'rows[0x06] >= 2048 || rows[0x04] >= 2048 || rows[0x01] >= 2048 || rows[0x02] >= 2048 || rows[0x08] >= 2048 || rows[0x09] >= 2048 || rows[0x0a] >= 2048 || rows[0x00] >= 2048 || rows[0x0e] >= 2048 || rows[0x17] >= 2048 || rows[0x14] >= 2048 || rows[0x11] >= 2048 || rows[0x1a] >= 2048 || rows[0x1b] >= 2048 || rows[0x20] >= 2048 || rows[0x23] >= 2048 || rows[0x26] >= 2048 || rows[0x27] >= 2048 || rows[0x28] >= 2048 )
        If iRows(6) >= 2048 Or iRows(4) >= 2048 Or iRows(1) >= 2048 Or iRows(2) >= 2048 Or iRows(8) >= 2048 Or iRows(9) >= 2048 Or iRows(&HA) >= 2048 Or iRows(0) >= 2048 Or iRows(&HE) >= 2048 Or iRows(&H17) >= 2048 Or iRows(&H14) >= 2048 Or iRows(&H11) >= 2048 Or iRows(&H1A) >= 2048 Or iRows(&H1B) >= 2048 Or iRows(&H20) >= 2048 Or iRows(&H23) >= 2048 Or iRows(&H26) >= 2048 Or iRows(&H27) >= 2048 Or iRows(&H28) >= 2048 Then
            GetCodedIndexSize = 4
        Else
            GetCodedIndexSize = 2
        End If
    End If
    If strTableName = "HasCustomAttributeType" Then
        'if ( rows[2] >= 8192 || rows[1] >= 8192 || rows[6] >= 8192 || rows[0x0a] >= 8192 )
        If iRows(2) >= 8192 Or iRows(1) >= 8192 Or iRows(6) >= 8192 Or iRows(&HA) >= 8192 Then
            GetCodedIndexSize = 4
        Else
            GetCodedIndexSize = 2
        End If
    End If
    
End Function
Private Sub FillTableSizes()
    Dim modulesize  As Long
    Dim typerefsize  As Long
    Dim typedefsize As Long
    Dim fieldsize As Long
    Dim methodsize As Long
    Dim paramsize As Long
    Dim interfaceimplsize As Long
    Dim memberrefsize As Long
    Dim constantsize  As Long
    Dim customattributesize As Long
    Dim fieldmarshallsize As Long
    Dim declsecuritysize As Long
    Dim classlayoutsize As Long
    Dim fieldlayoutsize As Long
    Dim stanalonssigsize As Long
    Dim eventmapsize As Long
    Dim eventsize As Long
    Dim propertymapsize As Long
    Dim propertysize  As Long
    Dim methodsemantics As Long
    Dim methodimplsize As Long
    Dim modulerefsize As Long
    Dim typespecsize As Long
    Dim implmapsize As Long
    Dim fieldrvasize As Long
    Dim assemblysize As Long
    Dim assemblyrefsize  As Long
    Dim filesize As Long
    Dim exportedtype As Long
    Dim manifestresourcesize  As Long
    Dim nestedclasssize As Long
    
    
    modulesize = 2 + OffsetString + OffsetGuid + OffsetGuid + OffsetGuid
    typerefsize = GetCodedIndexSize("ResolutionScope") + OffsetString + OffsetString
    typedefsize = 4 + OffsetString + OffsetString + GetCodedIndexSize("TypeDefOrRef") + GetTableSize("Method") + GetTableSize("Field")
    fieldsize = 2 + OffsetString + OffsetBlob
    methodsize = 4 + 2 + 2 + OffsetString + OffsetBlob + GetTableSize("Param")
    paramsize = 2 + 2 + OffsetString
    interfaceimplsize = GetTableSize("TypeDef") + GetCodedIndexSize("TypeDefOrRef")
    memberrefsize = GetCodedIndexSize("MemberRefParent") + OffsetString + OffsetBlob
    constantsize = 2 + GetCodedIndexSize("HasConst") + OffsetBlob
    customattributesize = GetCodedIndexSize("HasCustomAttribute") + GetCodedIndexSize("HasCustomAttributeType") + OffsetBlob
    fieldmarshallsize = GetCodedIndexSize("HasFieldMarshal") + OffsetBlob
    declsecuritysize = 2 + GetCodedIndexSize("HasDeclSecurity") + OffsetBlob
    classlayoutsize = 2 + 4 + GetTableSize("TypeDef")
    fieldlayoutsize = 4 + GetTableSize("Field")
    stanalonssigsize = OffsetBlob
    eventmapsize = GetTableSize("TypeDef") + GetTableSize("Event")
    eventsize = 2 + OffsetString + GetCodedIndexSize("TypeDefOrRef")
    propertymapsize = GetTableSize("Properties") + GetTableSize("TypeDef")
    propertysize = 2 + OffsetString + OffsetBlob
    methodsemantics = 2 + GetTableSize("Method") + GetCodedIndexSize("HasSemantics")
    methodimplsize = GetTableSize("TypeDef") + GetCodedIndexSize("MethodDefOrRef") + GetCodedIndexSize("MethodDefOrRef")
    modulerefsize = OffsetString
    typespecsize = OffsetBlob
    implmapsize = 2 + GetCodedIndexSize("MemberForwarded") + OffsetString + GetTableSize("ModuleRef")
    fieldrvasize = 4 + GetTableSize("Field")
    assemblysize = 4 + 2 + 2 + 2 + 2 + 4 + OffsetBlob + OffsetString + OffsetString
    assemblyrefsize = 2 + 2 + 2 + 2 + 4 + OffsetBlob + OffsetString + OffsetString + OffsetBlob
    filesize = 4 + OffsetString + OffsetBlob
    exportedtype = 4 + 4 + OffsetString + OffsetString + GetCodedIndexSize("Implementation")
    manifestresourcesize = 4 + 4 + OffsetString + GetCodedIndexSize("Implementation")
    nestedclasssize = GetTableSize("TypeDef") + GetTableSize("TypeDef")
    
     'sizes  ='( modulesize, typerefsize , typedefsize ,2, fieldsize ,2,methodsize ,2,paramsize ,interfaceimplsize,memberrefsize ,constantsize ,customattributesize ,fieldmarshallsize ,declsecuritysize ,classlayoutsize ,fieldlayoutsize,stanalonssigsize ,eventmapsize ,2,eventsize ,propertymapsize ,2,propertysize ,methodsemantics ,methodimplsize ,modulerefsize ,typespecsize ,implmapsize ,fieldrvasize ,2 , 2 , assemblysize ,4,12,assemblyrefsize ,6,14,filesize ,exportedtype ,manifestresourcesize ,nestedclasssize)
    ReDim Sizes(42)
    Sizes(0) = modulesize
    Sizes(1) = typerefsize
    Sizes(2) = typedefsize
    Sizes(3) = 2
    Sizes(4) = fieldsize
    Sizes(5) = 2
    Sizes(6) = methodsize
    Sizes(7) = 2
    Sizes(8) = paramsize
    Sizes(9) = interfaceimplsize
    Sizes(10) = memberrefsize
    Sizes(11) = constantsize
    Sizes(12) = customattributesize
    Sizes(13) = fieldmarshallsize
    Sizes(14) = declsecuritysize
    Sizes(15) = classlayoutsize
    Sizes(16) = fieldlayoutsize
    Sizes(17) = stanalonssigsize
    Sizes(18) = eventmapsize
    Sizes(19) = 2
    Sizes(20) = eventsize
    Sizes(21) = propertymapsize
    Sizes(22) = 2
    Sizes(23) = propertysize
    Sizes(24) = methodsemantics
    Sizes(25) = methodimplsize
    Sizes(26) = modulerefsize
    Sizes(27) = typespecsize
    Sizes(28) = implmapsize
    Sizes(29) = fieldrvasize
    Sizes(30) = 2
    Sizes(31) = 2
    Sizes(32) = assemblysize
    Sizes(33) = 4
    Sizes(34) = 12
    Sizes(35) = assemblyrefsize
    Sizes(36) = 6
    Sizes(37) = 14
    Sizes(38) = filesize
    Sizes(39) = exportedtype
    Sizes(40) = manifestresourcesize
    Sizes(41) = nestedclasssize
    
    'fieldmarshallsize ,declsecuritysize ,classlayoutsize ,fieldlayoutsize,stanalonssigsize ,eventmapsize
    ',2,eventsize ,propertymapsize ,2,propertysize ,methodsemantics ,methodimplsize ,modulerefsize
    ',typespecsize ,implmapsize ,fieldrvasize ,2 , 2 , assemblysize ,4,12,assemblyrefsize ,6,14,
    'filesize ,exportedtype ,manifestresourcesize ,nestedclasssize)
    
End Sub
Private Function GetTableSize(strTableName As String) As Long
    Dim i As Long
    For i = 0 To UBound(strTableNames)
        If strTableNames(i) = strTableName Then
            Exit For
        End If
        
    Next
    If iRows(i) > 65535 Then
        GetTableSize = 4
    Else
        GetTableSize = 2
    End If
    
End Function
Public Sub ReadTablesIntoStructures()
    Dim old As Long
    Dim tablehasrows As Boolean
    Dim offs As Long
    Dim k As Integer
    old = TableOffset
    tablehasrows = tablepresent(0)
    offs = TableOffset
    TableOffset = old
    'MsgBox "TABLEOffset: " & TableOffset & " OffsetString: " & OffsetString & " OffsetGuid:" & OffsetGuid & " OffsetBlob: " & OffsetBlob
    Console.WriteLine ("Module Table Offset " & offs & " Size " & Sizes(0))
    'Module
   ' MsgBox "Module Table Offset " & offs & " Size " & Sizes(0)
    
    If tablehasrows = True Then
        ReDim ModuleStruct(iRows(0) + 1)
        For k = 1 To iRows(0)
            ModuleStruct(k).Generation = BitConverterToUInt16(MetaDataByteArray, offs)
            offs = offs + 2
            ModuleStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            'MsgBox "Mod" & ModuleStruct(i).name
            offs = offs + OffsetString
            ModuleStruct(k).Mvid = ReadGuidIndex(MetaDataByteArray, offs)
            offs = offs + OffsetGuid
            ModuleStruct(k).EncId = ReadGuidIndex(MetaDataByteArray, offs)
            offs = offs + OffsetGuid
            ModuleStruct(k).EncBaseId = ReadGuidIndex(MetaDataByteArray, offs)
            offs = offs + OffsetGuid
        Next
    End If
    'TypeRef
    old = TableOffset
    tablehasrows = tablepresent(1)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim TypeRefStruct(iRows(1) + 1)
        For k = 1 To iRows(1)
            TypeRefStruct(k).resolutionscope = ReadCodedIndex(MetaDataByteArray, offs, "ResolutionScope")
            offs = offs + GetCodedIndexSize("ResolutionScope")
            TypeRefStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            TypeRefStruct(k).nspace = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
        Next
    End If
    'TypeDef
    old = TableOffset
    tablehasrows = tablepresent(2)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim TypeDefStruct(iRows(2) + 1)
        For k = 1 To iRows(2)
            TypeDefStruct(k).flags = BitConverterToInt32(MetaDataByteArray, offs)
            offs = offs + 4
            TypeDefStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            TypeDefStruct(k).nspace = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            TypeDefStruct(k).cindex = ReadCodedIndex(MetaDataByteArray, offs, "TypeDefOrRef")
            offs = offs + GetCodedIndexSize("TypeDefOrRef")
            TypeDefStruct(k).findex = ReadTableIndex(MetaDataByteArray, offs, "Field")
            offs = offs + GetTableSize("Field")
            TypeDefStruct(k).mindex = ReadTableIndex(MetaDataByteArray, offs, "Method")
            offs = offs + GetTableSize("Method")
        Next
    End If
    'FieldPtr
    old = TableOffset
    tablehasrows = tablepresent(3)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim FieldPtrStruct(iRows(3) + 1)
        For k = 1 To iRows(3)
            FieldPtrStruct(k).index = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 1
        Next
    End If
    'Field
    old = TableOffset
    tablehasrows = tablepresent(4)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim FieldStruct(iRows(4) + 1)
        For k = 1 To iRows(4)
            FieldStruct(k).flags = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            FieldStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            FieldStruct(k).sig = ReadBlobIndex(MetaDataByteArray, offs)
            offs = offs + OffsetBlob
        Next
    End If
    'MethodPtr
    old = TableOffset
    tablehasrows = tablepresent(5)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim MethodPtrStruct(iRows(5) + 1)
        For k = 1 To iRows(5)
            MethodPtrStruct(k).index = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
        Next
    End If
    'Method
    old = TableOffset
    tablehasrows = tablepresent(6)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim MethodStruct(iRows(6) + 1)
        For k = 1 To iRows(6)
            MethodStruct(k).rva = BitConverterToInt32(MetaDataByteArray, offs)
            offs = offs + 4
            MethodStruct(k).impflags = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            MethodStruct(k).flags = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            MethodStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            MethodStruct(k).signature = ReadBlobIndex(MetaDataByteArray, offs)
            offs = offs + OffsetBlob
            MethodStruct(k).param = ReadTableIndex(MetaDataByteArray, offs, "Param")
            offs = offs + GetTableSize("Param")
        Next
    End If
    'Param
    old = TableOffset
    tablehasrows = tablepresent(8)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim ParamStruct(iRows(8) + 1)
        For k = 1 To iRows(8)
            ParamStruct(k).pAttr = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            ParamStruct(k).sequence = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            ParamStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
        Next
    End If
    'InterfaceImpl
    old = TableOffset
    tablehasrows = tablepresent(9)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim InterfaceImplStruct(iRows(9) + 1)
        For k = 1 To iRows(9)
            InterfaceImplStruct(k).classindex = ReadCodedIndex(MetaDataByteArray, offs, "TypeDefOrRef")
            offs = offs + GetCodedIndexSize("TypeDefOrRef")
            InterfaceImplStruct(k).interfaceindex = ReadTableIndex(MetaDataByteArray, offs, "TypeDef")
            offs = offs + GetTableSize("TypeDef")
        Next
    End If
    'MemberRef
    old = TableOffset
    tablehasrows = tablepresent(10)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim MemberRefStruct(iRows(10) + 1)
        For k = 1 To iRows(10)
            MemberRefStruct(k).clas = ReadCodedIndex(MetaDataByteArray, offs, "MemberRefParent")
            offs = offs + GetCodedIndexSize("MemberRefParent")
            MemberRefStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            MemberRefStruct(k).sig = ReadBlobIndex(MetaDataByteArray, offs)
            offs = offs + OffsetBlob
        Next k
    End If
    'Constants
    old = TableOffset
    tablehasrows = tablepresent(11)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim ConstantsStruct(iRows(11) + 1)
        For k = 1 To iRows(11)
            ConstantsStruct(k).dtype = MetaDataByteArray(offs)
            offs = offs + 2
            ConstantsStruct(k).parent = ReadCodedIndex(MetaDataByteArray, offs, "HasConst")
            offs = offs + GetCodedIndexSize("HasConst")
            ConstantsStruct(k).value = ReadBlobIndex(MetaDataByteArray, offs)
            offs = offs + OffsetBlob
        Next
        
    End If
    'CustomAttribute
    old = TableOffset
    tablehasrows = tablepresent(12)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim CustomAttributeStruct(iRows(12) + 1)
        For k = 1 To iRows(12)
            CustomAttributeStruct(k).parent = ReadCodedIndex(MetaDataByteArray, offs, "HasCustomAttribute")
            offs = offs + GetCodedIndexSize("HasCustomAttribute")
            CustomAttributeStruct(k).type = ReadCodedIndex(MetaDataByteArray, offs, "HasCustomAttributeType")
            offs = offs + GetCodedIndexSize("HasCustomAttributeType")
            CustomAttributeStruct(k).value = ReadBlobIndex(MetaDataByteArray, offs)
            offs = offs + OffsetBlob
        Next
        
    End If
    'FieldMarshal
    old = TableOffset
    tablehasrows = tablepresent(13)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim FieldMarshalStruct(iRows(13) + 1)
        For k = 1 To iRows(13)
            FieldMarshalStruct(k).coded = ReadCodedIndex(MetaDataByteArray, offs, "HasFieldMarshal")
            offs = offs + GetCodedIndexSize("HasFieldMarshal")
            FieldMarshalStruct(k).index = ReadBlobIndex(MetaDataByteArray, offs)
            offs = offs + OffsetBlob
        Next
    End If
    'DeclSecurity
    old = TableOffset
    tablehasrows = tablepresent(14)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim DeclSecurityStruct(iRows(14) + 1)
        For k = 1 To iRows(14)
            DeclSecurityStruct(k).action = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            DeclSecurityStruct(k).coded = ReadCodedIndex(MetaDataByteArray, offs, "HasDeclSecurity")
            offs = offs + GetCodedIndexSize("HasDeclSecurity")
            DeclSecurityStruct(k).bindex = ReadBlobIndex(MetaDataByteArray, offs)
            offs = offs + OffsetBlob
        Next
    End If
    'ClassLayout
    old = TableOffset
    tablehasrows = tablepresent(15)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim ClassLayoutStruct(iRows(15) + 1)
        For k = 1 To iRows(15)
            ClassLayoutStruct(k).packingsize = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            ClassLayoutStruct(k).classsize = BitConverterToInt32(MetaDataByteArray, offs)
            offs = offs + 4
            ClassLayoutStruct(k).parent = ReadTableIndex(MetaDataByteArray, offs, "TypeDef")
            offs = offs + GetTableSize("TypeDef")
        Next
        
    End If
    'FieldLayout
    old = TableOffset
    tablehasrows = tablepresent(16)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim FieldLayoutStruct(iRows(16) + 1)
        For k = 1 To iRows(16)
            FieldLayoutStruct(k).offset = BitConverterToInt32(MetaDataByteArray, offs)
            offs = offs + 4
            FieldLayoutStruct(k).fieldindex = ReadTableIndex(MetaDataByteArray, offs, "Field")
            offs = offs + GetTableSize("Field")
        Next
    End If
    'StandAloneSig
    old = TableOffset
    tablehasrows = tablepresent(17)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim StandAloneSigStruct(iRows(17) + 1)
        For k = 1 To iRows(17)
            StandAloneSigStruct(k).index = ReadBlobIndex(MetaDataByteArray, offs)
            offs = offs + OffsetBlob
        Next
    End If
    'EventMap
    old = TableOffset
    tablehasrows = tablepresent(18)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim EventMapStruct(iRows(18) + 1)
        For k = 1 To iRows(18)
            EventMapStruct(k).index = ReadTableIndex(MetaDataByteArray, offs, "TypeDef")
            offs = offs + GetTableSize("TypeDef")
            EventMapStruct(k).eindex = ReadTableIndex(MetaDataByteArray, offs, "Event")
            offs = offs + GetTableSize("Event")
        Next
    End If
    'Event
    old = TableOffset
    tablehasrows = tablepresent(20)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim EventStruct(iRows(20) + 1)
        For k = 1 To iRows(20)
            EventStruct(k).attr = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            EventStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            EventStruct(k).coded = ReadCodedIndex(MetaDataByteArray, offs, "TypeDefOrRef")
            offs = offs + GetCodedIndexSize("TypeDefOrRef")
        Next
    End If
    'PropertyMap
    old = TableOffset
    tablehasrows = tablepresent(21)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim PropertyMapStruct(iRows(21) + 1)
        For k = 1 To iRows(21)
            PropertyMapStruct(k).parent = ReadTableIndex(MetaDataByteArray, offs, "TypeDef")
            offs = offs + GetTableSize("TypeDef")
            PropertyMapStruct(k).propertylist = ReadTableIndex(MetaDataByteArray, offs, "Properties")
            offs = offs + GetTableSize("Properties")
        Next
    End If
    'Property
    old = TableOffset
    tablehasrows = tablepresent(23)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim PropertyStruct(iRows(23) + 1)
        For k = 1 To iRows(23)
            PropertyStruct(k).flags = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            PropertyStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            PropertyStruct(k).type = ReadBlobIndex(MetaDataByteArray, offs)
            offs = offs + OffsetBlob
        Next
    End If
    'MethodSemantics
    old = TableOffset
    tablehasrows = tablepresent(24)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim MethodSemanticsStruct(iRows(24) + 1)
        For k = 1 To iRows(24)
            MethodSemanticsStruct(k).methodsemanticsattributes = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            MethodSemanticsStruct(k).methodindex = ReadTableIndex(MetaDataByteArray, offs, "Method")
            offs = offs + GetTableSize("Method")
            MethodSemanticsStruct(k).association = ReadCodedIndex(MetaDataByteArray, offs, "HasSemantics")
            offs = offs + GetCodedIndexSize("HasSemantics")
        Next
    End If
    'MethodImpl
    old = TableOffset
    tablehasrows = tablepresent(25)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim MethodImpStruct(iRows(25) + 1)
        For k = 1 To iRows(25)
            MethodImpStruct(k).classindex = ReadTableIndex(MetaDataByteArray, offs, "TypeDef")
            offs = offs + GetTableSize("TypeDef")
            MethodImpStruct(k).codedbody = ReadCodedIndex(MetaDataByteArray, offs, "MethodDefOrRef")
            offs = offs + GetCodedIndexSize("MethodDefOrRef")
            MethodImpStruct(k).codeddef = ReadCodedIndex(MetaDataByteArray, offs, "MethodDefOrRef")
            offs = offs + GetCodedIndexSize("MethodDefOrRef")
        Next
    End If
    'ModuleRef
    old = TableOffset
    tablehasrows = tablepresent(26)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim ModuleRefStruct(iRows(26) + 1)
        For k = 1 To iRows(26)
            ModuleRefStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
        Next
    End If
    'TypeSpec
    old = TableOffset
    tablehasrows = tablepresent(27)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim TypeSpecStruct(iRows(27) + 1)
        For k = 1 To iRows(27)
            TypeSpecStruct(k).signature = ReadBlobIndex(MetaDataByteArray, offs)
            offs = offs + OffsetBlob
        Next
    End If
    'ImplMap
    old = TableOffset
    tablehasrows = tablepresent(28)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim ImplMapStruct(iRows(28) + 1)
        For k = 1 To iRows(28)
            ImplMapStruct(k).attr = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            ImplMapStruct(k).cindex = ReadCodedIndex(MetaDataByteArray, offs, "MemberForwarded")
            offs = offs + GetCodedIndexSize("MemberForwarded")
            ImplMapStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            ImplMapStruct(k).scope = ReadTableIndex(MetaDataByteArray, offs, "ModuleRef")
            offs = offs + GetTableSize("ModuleRef")
        Next
    End If
    'FieldRVA
    old = TableOffset
    tablehasrows = tablepresent(29)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim FieldRVAStruct(iRows(29) + 1)
        For k = 1 To iRows(29)
            FieldRVAStruct(k).rva = BitConverterToInt32(MetaDataByteArray, offs)
            offs = offs + 4
            FieldRVAStruct(k).fieldi = ReadTableIndex(MetaDataByteArray, offs, "Field")
            offs = offs + GetTableSize("Field")
        Next
    End If
    'Assembley
    old = TableOffset
    tablehasrows = tablepresent(32)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim AssemblyStruct(iRows(32) + 1)
        For k = 1 To iRows(32)
            AssemblyStruct(k).HashAlgId = BitConverterToInt32(MetaDataByteArray, offs)
            offs = offs + 4
            AssemblyStruct(k).major = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            AssemblyStruct(k).minor = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            AssemblyStruct(k).build = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            AssemblyStruct(k).revision = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            AssemblyStruct(k).flags = BitConverterToInt32(MetaDataByteArray, offs)
            offs = offs + 4
            AssemblyStruct(k).publickey = ReadBlobIndex(MetaDataByteArray, offs)
            offs = offs + OffsetBlob
            AssemblyStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            AssemblyStruct(k).culture = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
        Next
    End If
    'AssemblyRef
    old = TableOffset
    tablehasrows = tablepresent(35)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim AssemblyRefStruct(iRows(35) + 1)
        For k = 1 To iRows(35)
            AssemblyRefStruct(k).major = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            AssemblyRefStruct(k).minor = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            AssemblyRefStruct(k).build = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            AssemblyRefStruct(k).revision = BitConverterToInt16(MetaDataByteArray, offs)
            offs = offs + 2
            AssemblyRefStruct(k).flags = BitConverterToInt32(MetaDataByteArray, offs)
            offs = offs + 4
            AssemblyRefStruct(k).publickey = ReadBlobIndex(MetaDataByteArray, offs)
            offs = offs + OffsetBlob
            AssemblyRefStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            AssemblyRefStruct(k).culture = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            AssemblyRefStruct(k).hashvalue = ReadBlobIndex(MetaDataByteArray, offs)
            offs = offs + OffsetBlob
        Next
    End If
    
    'File
    old = TableOffset
    tablehasrows = tablepresent(38)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim FileStruct(iRows(38) + 1)
        For k = 1 To iRows(38)
            FileStruct(k).flags = BitConverterToInt32(MetaDataByteArray, offs)
            offs = offs + 4
            FileStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            FileStruct(k).index = ReadBlobIndex(MetaDataByteArray, offs)
            offs = offs + OffsetBlob
        Next
    End If
    'ExportedType
    old = TableOffset
    tablehasrows = tablepresent(39)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim ExportedTypeStruct(iRows(39) + 1)
        For k = 1 To iRows(39)
            ExportedTypeStruct(k).flags = BitConverterToInt32(MetaDataByteArray, offs)
            offs = offs + 4
            ExportedTypeStruct(k).typedefindex = BitConverterToInt32(MetaDataByteArray, offs)
            offs = offs + 4
            ExportedTypeStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            ExportedTypeStruct(k).nspace = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            ExportedTypeStruct(k).coded = ReadCodedIndex(MetaDataByteArray, offs, "Implementation")
            offs = offs + GetCodedIndexSize("Implementation")
        Next
    End If
    'ManifestResource
    old = TableOffset
    tablehasrows = tablepresent(40)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim ManifestResourceStruct(iRows(40) + 1)
        For k = 1 To iRows(40)
            ManifestResourceStruct(k).offset = BitConverterToInt32(MetaDataByteArray, offs)
            offs = offs + 4
            ManifestResourceStruct(k).flags = BitConverterToInt32(MetaDataByteArray, offs)
            offs = offs + 4
            ManifestResourceStruct(k).name = ReadStringIndex(MetaDataByteArray, offs)
            offs = offs + OffsetString
            ManifestResourceStruct(k).coded = ReadCodedIndex(MetaDataByteArray, offs, "Implementation")
            offs = offs + GetCodedIndexSize("")
        Next
    End If
    'Nested Classes
    old = TableOffset
    tablehasrows = tablepresent(41)
    offs = TableOffset
    TableOffset = old
    If tablehasrows = True Then
        ReDim NestedClassStruct(iRows(41) + 1)
        For k = 1 To iRows(41)
            NestedClassStruct(k).nestedclass = ReadTableIndex(MetaDataByteArray, offs, "TypeDef")
            offs = offs + GetTableSize("TypeDef")
            NestedClassStruct(k).enclosingclass = ReadTableIndex(MetaDataByteArray, offs, "TypeDef")
            offs = offs + GetTableSize("TypeDef")
        Next
    End If
    
End Sub
Public Function tablepresent(tableindex As Byte) As Boolean
    Dim tablebit As Long
    'tablebit = HiLo.DWordShiftR(Valid, tableindex) And 1
    tablebit = isTablePresent(MetaDataByteArray, tableindex)
    Dim J As Integer
    'For j = 0 To tableindex
    Do While J < tableindex
        Dim o As Integer
        o = Sizes(J) * iRows(J)
        TableOffset = TableOffset + o
        J = J + 1
    Loop
    'Next
    If tablebit = 1 Then
        tablepresent = True
    Else
        tablepresent = False
    End If
    
End Function
Public Function ReadCodedIndex(metadataarray() As Byte, offset As Long, nameoftable As String) As Long
    Dim ReturnIndex  As Long
    Dim codedindexsize As Long
    codedindexsize = GetCodedIndexSize(nameoftable)
    If codedindexsize = 2 Then
        ReturnIndex = BitConverterToUInt16(metadataarray, offset)
    End If
    If codedindexsize = 4 Then
        ReturnIndex = BitConverterToUInt32(metadataarray, offset)
    End If
    ReadCodedIndex = ReturnIndex
End Function
Public Function ReadTableIndex(metadataarray() As Byte, arrayoffset As Long, tablename As String) As Long
    Dim ReturnIndex  As Long
    Dim tableSize As Long
    tableSize = GetTableSize(tablename)
    If tableSize = 2 Then
        ReturnIndex = BitConverterToUInt16(metadataarray, arrayoffset)
    End If
    
    If tableSize = 4 Then
        ReturnIndex = BitConverterToUInt32(metadataarray, arrayoffset)
    End If
    
    ReadTableIndex = ReturnIndex
End Function
Public Function ReadStringIndex(metadataarray() As Byte, arrayoffset As Long) As Long
    Dim ReturnIndex  As Long
    If OffsetString = 2 Then
        ReturnIndex = BitConverterToUInt16(metadataarray, arrayoffset)
    End If
    
    If OffsetString = 4 Then
        ReturnIndex = BitConverterToUInt32(metadataarray, arrayoffset)
    End If
    
    ReadStringIndex = ReturnIndex
End Function
Public Function ReadBlobIndex(metadataarray() As Byte, arrayoffset As Long) As Long
    Dim ReturnIndex  As Long
    If OffsetBlob = 2 Then
        ReturnIndex = BitConverterToUInt16(metadataarray, arrayoffset)
    End If
    If OffsetBlob = 4 Then
        ReturnIndex = BitConverterToUInt32(metadataarray, arrayoffset)
    End If
    ReadBlobIndex = ReturnIndex
End Function
Public Function ReadGuidIndex(metadataarray() As Byte, arrayoffset As Long)
    Dim ReturnIndex  As Long
    If OffsetGuid = 2 Then
        ReturnIndex = BitConverterToUInt16(metadataarray, arrayoffset)
    End If
    If OffsetGuid = 4 Then
        ReturnIndex = BitConverterToUInt32(metadataarray, arrayoffset)
    End If
    
    ReadGuidIndex = ReturnIndex
End Function
Public Function GetResolutionScopeTable(rvalue As Long) As String
    Dim tag As Long
    Dim returnstring As String
    tag = rvalue And &H3
    If tag = 0 Then
        returnstring = returnstring & "Module"
    End If
    If tag = 1 Then
        returnstring = returnstring & "ModuleRef"
    End If
    If tag = 2 Then
        returnstring = returnstring & "AssemblyRef"
    End If
    If tag = 3 Then
        returnstring = returnstring & "TypeRef"
    End If
    
    GetResolutionScopeTable = returnstring
    
    
End Function
Private Function GetString(starting As Long) As String
    Dim ending As Long
    ending = starting
    If starting < 0 Then
        GetString = ""
        Exit Function
    End If
    If starting >= UBound(StringsByteArray) Then
        GetString = ""
        Exit Function
    End If
    Dim strBuffer As String
    Do While (StringsByteArray(ending) <> 0)
        strBuffer = strBuffer & Chr$(StringsByteArray(ending))
        ending = ending + 1
    Loop
    GetString = strBuffer
End Function

Private Function GetResolutionScopeValue(rvalue As Long) As Long
    GetResolutionScopeValue = DoRightBitShift(rvalue, 2)
End Function
Private Function CreateSpaces(lNumber As Long) As String
    CreateSpaces = Space$(lNumber)
End Function
Private Function CorSigUncompressData(blobarray() As Byte, index As Integer, answer As Integer) As Long
    Dim howmanybytes As Long
    howmanybytes = 0
    answer = 0
    If (blobarray(index) And &H80) = 0 Then
        howmanybytes = 0
        answer = blobarray(index)
    End If

    If (blobarray(index) And &HC0) = &H80 Then
        howmanybytes = 2
        '#TODO answer = ((blobarray[index] & 0x3f) <<8 ) or  blobarray(index+1)
    End If
    If (blobarray(index) And &HE0) = &HC0 Then
        howmanybytes = 3
        '#TODO answer = ((blobarray[index] & 0x1f) <<24 ) |  (blobarray[index+1] << 16) |  (blobarray[index+2] << 8) | blobarray[index+3];
    End If
    CorSigUncompressData = howmanybytes
End Function
Private Function GetActionSecurity(actionbyte As Long) As String
    Select Case actionbyte
        Case 1
             GetActionSecurity = "request"
        Case 2
            GetActionSecurity = "demand"
        Case 3
            GetActionSecurity = "assert"
        Case 4
            GetActionSecurity = "deny"
        Case 5
            GetActionSecurity = "permitonly"
        Case 6
            GetActionSecurity = "linkcheck"
        Case 7
            GetActionSecurity = "inheritcheck"
        Case 8
            GetActionSecurity = "reqmin"
        Case 9
            GetActionSecurity = "reqopt"
        Case 10
            GetActionSecurity = "reqrefuse"
        Case 11
            GetActionSecurity = "prejitgrant"
        Case 12
            GetActionSecurity = "prejitdeny"
        Case 13
            GetActionSecurity = "noncasdemand"
        Case 14
            GetActionSecurity = "noncaslinkdemand"
        Case 15
            GetActionSecurity = "noncasinheritance"
    End Select
End Function
Private Sub ReadandDisplayExportAddressTableJumps()
    Console.WriteLine ("// Export Address Table Jumps:")
    If modVBNET.gVBNETHeader.ExportAddressTableJumps.VirtualAddress = 0 Then
        Console.WriteLine ("// No data.")
        Console.WriteLine ("")
    End If
End Sub
Private Sub DisplayModuleRefs()
    Dim i As Long
    Dim strData As String
    
    If tablepresent(26) = False Then Exit Sub
    
    For i = 0 To UBound(ModuleRefStruct)
        strData = GetString(ModuleRefStruct(i).name)
        Console.WriteLine (".module extern " & strData)
    Next
End Sub
Private Function GetTypeAttributeFlagsForClassExtern(typeattributeflags As Long) As String
    typeattributeflags = typeattributeflags And &H7
    If typeattributeflags = 1 Then
        GetTypeAttributeFlagsForClassExtern = "public"
    End If
    If typeattributeflags = 2 Then
        GetTypeAttributeFlagsForClassExtern = "nested public"
    End If
End Function
Private Function GetManifestResourceValue(manifiestvalue As Long) As Long
'#TODO return manifiestvalue>> 2
GetManifestResourceValue = DoRightBitShift(manifiestvalue, 2)
    
End Function
Private Function GetManifestResourceTable(manifiestvalue As Long) As String
    Dim tag As Integer
    Dim returnstring As String
    tag = manifiestvalue And &H3
    If tag = 0 Then
        returnstring = returnstring & "File"
    End If
    If tag = 1 Then
        returnstring = returnstring & "AssemblyRef"
    End If
    GetManifestResourceTable = returnstring
End Function
Private Function GetManifestResourceAttributes(manifiestvalue As Long) As String
    Dim returnstring As String
    If (manifiestvalue And &H1) = 1 Then
        returnstring = returnstring & "public "
    End If
    If (manifiestvalue And &H2) = 2 Then
        returnstring = returnstring & "private "
    End If
    GetManifestResourceAttributes = returnstring
End Function
Private Function IsTypeNested(typeindex As Long) As Boolean
    'If NestedClassStruct = Null Then
    '    IsTypeNested = False
    '    Exit Function
    'End If
    If tablepresent(41) = False Then
        IsTypeNested = False
        Exit Function
    End If
    
    Dim i As Integer
    For i = 0 To UBound(NestedClassStruct)
        If NestedClassStruct(i).nestedclass = typeindex Then
            IsTypeNested = True
            Exit Function
        End If
    Next
    IsTypeNested = False
End Function
Private Function GetTypeAttributeFlags(typeattributeflags As Long, typeindex As Long) As String
    Dim returnstring As String
    Dim Visibiltymask  As Long
    Dim visibiltymaskstring As String
    Visibiltymask = typeattributeflags And &H7
    Select Case Visibiltymask
        Case 0
            visibiltymaskstring = "private "
        Case 1
            visibiltymaskstring = "public "
        Case 2
            visibiltymaskstring = "nested public "
        Case 3
            visibiltymaskstring = "nested private "
        Case 4
            visibiltymaskstring = "nested family "
        Case 5
            visibiltymaskstring = "nested assembly "
        Case 6
            visibiltymaskstring = "nested famandassem "
        Case 7
            visibiltymaskstring = "nested famorassem "
            
    End Select
    Dim classlayoutmask As Long
    classlayoutmask = typeattributeflags And &H18
    Dim classlayoutstring  As String
    If classlayoutmask = 0 Then
        classlayoutstring = "auto "
    End If
    If classlayoutmask = 8 Then
        classlayoutstring = "sequential "
    End If
    If classlayoutmask = &H10 Then
        classlayoutstring = "explicit "
    End If
    Dim interfacestring As String
    
    If (typeattributeflags And &H20) = &H20 Then
        interfacestring = "interface "
    End If
    Dim abstractstring As String
    If (typeattributeflags And &H80) = &H80 Then
        abstractstring = "abstract "
    End If
    Dim sealedstring As String
    If (typeattributeflags And &H100) = &H100 Then
        sealedstring = "sealed "
    End If
    Dim specialnamestring As String
    If (typeattributeflags And &H400) = &H400 Then
        specialnamestring = "specialname "
    End If
    Dim importstring As String
    If (typeattributeflags And &H1000) = &H1000 Then
        importstring = "import "
    End If
    Dim serializablestring As String
    If (typeattributeflags And &H2000) = &H2000 Then
        serializablestring = "serializable "
    End If
    Dim stringformatmask As Long
    stringformatmask = typeattributeflags And &H30000
    Dim stringformastring As String
    If stringformatmask = 0 Then
        stringformastring = "ansi "
    End If
    If stringformatmask = &H10000 Then
        stringformastring = "unicode "
    End If
    If stringformatmask = &H20000 Then
        stringformastring = "autochar "
    End If
    Dim beforefieldinitstring As String
    If (typeattributeflags And &H100000) = &H100000 Then
        beforefieldinitstring = "beforefieldinit "
    End If
    If IsTypeNested(typeindex) Then
        returnstring = interfacestring & abstractstring & classlayoutstring & stringformastring & serializablestring & sealedstring & importstring & visibiltymaskstring & beforefieldinitstring
        GetTypeAttributeFlags = returnstring
    Else
        returnstring = interfacestring & visibiltymaskstring & abstractstring & classlayoutstring & stringformastring & importstring & serializablestring & sealedstring & specialnamestring & beforefieldinitstring
        GetTypeAttributeFlags = returnstring
    End If
End Function
Private Function GetTypeDefOrRefTable(codedvalue As Long) As String
    Dim returnstring As String
    Dim tag As Integer
    tag = codedvalue And &H3
    
    If tag = 0 Then
        returnstring = returnstring & "TypeDef"
    End If
    If tag = 1 Then
        returnstring = returnstring & "TypeRef"
    End If
    If tag = 2 Then
        returnstring = returnstring & "TypeSpec"
    End If
    GetTypeDefOrRefTable = returnstring
End Function
Private Function GetParentForNestedType(typeindex As Long) As Long
On Error GoTo errHandle
    'If NestedClassStruct = Null Then
    '    GetParentForNestedType = False
    '    Exit Function
    'End If
    Dim i As Integer
    For i = 0 To UBound(NestedClassStruct)
        If typeindex = NestedClassStruct(i).nestedclass Then
            Exit For
        End If
    Next
    GetParentForNestedType = NestedClassStruct(i).enclosingclass
Exit Function
errHandle:
Exit Function
End Function
Private Function GetPinvokeAttributes(attr As Long, returnattribute As String) As String
    Dim returnstring As String
    If (attr And &H1) = 1 Then
        returnattribute = " nomangle"
    End If
    If (attr And &H6) = 6 Then
        returnattribute = returnattribute & " autochar"
    ElseIf (attr And &H2) = 2 Then
        returnattribute = returnattribute & " ansi"
    ElseIf (attr And &H4) = 4 Then
        returnattribute = returnattribute & " unicode"
    End If
    If (attr And &H40) = &H40 Then
        returnattribute = returnattribute & " lasterr"
    End If
    
    If (attr And &H500) = &H500 Then
        returnstring = returnstring & "fastcall"
    End If
    If (attr And &H300) = &H300 Then
        returnstring = returnstring & "stdcall"
    End If
    If (attr And &H100) = &H100 Then
        returnstring = returnstring & "winapi"
    End If
    If (attr And &H200) = &H200 Then
        returnstring = returnstring & "cdecl"
    End If
    If (attr And &H400) = &H400 Then
        returnstring = returnstring & "thiscall"
    End If
    
    GetPinvokeAttributes = returnstring
End Function

Private Function GetSafeArrayType(bSafeArrayType As Byte) As String
    Select Case bSafeArrayType
        
        Case 0
            bSafeArrayType = ""
        Case 1
            bSafeArrayType = "null"
        Case 2
            bSafeArrayType = "int16"
        Case 3
            bSafeArrayType = "int32"
        Case 4
            bSafeArrayType = "float32"
        Case 5
            bSafeArrayType = "float34"
        Case 6
            bSafeArrayType = "currency"
        Case 7
            bSafeArrayType = "date"
        Case 8
            bSafeArrayType = "bstr"
        Case 9
            bSafeArrayType = "idispatch"
        Case &HA
            bSafeArrayType = "error"
        Case &HB
            bSafeArrayType = "bool"
        Case &HC
            bSafeArrayType = "variant"
        Case &HD
            bSafeArrayType = "iunknown"
        Case &HE
            bSafeArrayType = "decimal"
        Case &HF
            bSafeArrayType = "illegal"
        Case &H10
            bSafeArrayType = "int8"
        Case &H11
            bSafeArrayType = "unsigned int8"
        Case &H12
            bSafeArrayType = "unsigned int16"
        Case &H13
            bSafeArrayType = "unsigned int32"
        Case &H14
            bSafeArrayType = "int64"
        Case &H15
            bSafeArrayType = "unsigned int64"
        Case &H16
            bSafeArrayType = "int"
        Case &H17
            bSafeArrayType = "unsigned int"
        Case &H18
            bSafeArrayType = "void"
        Case &H19
            bSafeArrayType = "hresult"
        Case &H1A
            bSafeArrayType = "*"
        Case &H1B
            bSafeArrayType = "safearray"
        Case &H1C
            bSafeArrayType = "carray"
        Case &H1D
            bSafeArrayType = "userdefined"
        Case &H1E
            bSafeArrayType = "lpstr"
        Case &H1F
            bSafeArrayType = "lpwstr"
        Case &H20
            bSafeArrayType = "illegal"
        Case &H21
            bSafeArrayType = "illegal"
        Case &H22
            bSafeArrayType = "illegal"
        Case &H23
            bSafeArrayType = "illegal"
        Case &H24
            bSafeArrayType = "record"
        Case Else
            bSafeArrayType = "illegal"
            
    End Select
End Function
Private Function GetParamAttrforMethodCalling(methodindex As Long) As String

End Function
Public Sub CreateSignatures()

End Sub
Public Sub CreateSignatureForEachType(stype As Byte, index As Long, row As Long)

End Sub
Public Sub CreateMethodDefSignature(blobarray() As Byte, row As Integer)

End Sub
Private Function GetElementType(index As Long, blobarray() As Byte, ByRef howmanybytes As Long) As String
    howmanybytes = 0
    Dim returnstring As String
    Dim bType As Byte
    bType = blobarray(index)
    If bType >= &H1 And bType <= &HE Then
        returnstring = GetType(bType)
        howmanybytes = 1
    End If
    GetElementType = returnstring
End Function
Private Function GetFileAttributes(fileflags As Long) As String
    If fileflags = 1 Then
        GetFileAttributes = "nometadata "
    Else
        GetFileAttributes = ""
    
    End If
End Function
Public Sub DisplayFileTable()
    Dim i As Long
    On Error GoTo errHandle

    For i = 0 To UBound(FileStruct)
        Console.WriteLine (".file " & GetFileAttributes(FileStruct(i).flags) & NameReserved(GetString(FileStruct(i).name)))
    
        Dim table As Long
        'High byte of the metadata token is the table id (token = table<<24 | row);
        'the old code mistakenly called isTablePresent here.
        table = DoRightBitShift(gVBNETHeader.EntryPointToken, 24)
        If table = &H26 Then
            Dim row As Long
            row = gVBNETHeader.EntryPointToken And &HFFFFFF
            If row = i Then
                Console.WriteLine ("    .entrypoint")
            End If
            If FileStruct(i).index <> 0 Then
                Console.WriteLine ("    .hash = (")
                Dim index As Long
                index = FileStruct(i).index
                'DisplayFormattedColumns(index ,13 , false);
            End If
        End If
    Next
Exit Sub
errHandle:
Exit Sub
End Sub
Private Sub DisplayResources()

End Sub
Private Sub DisplayTypeDefs()
    If UBound(TypeDefStruct) <> 2 Then
        Console.WriteLine ("//")
        Console.WriteLine ("// ============== CLASS STRUCTURE DECLARATION ==================")
        Console.WriteLine ("//")
        Writenamespace = True
        Dim i  As Long
        For i = 2 To UBound(TypeDefStruct)
            If IsTypeNested(i) Then
                Call DisplayOneTypePrototype(i)
            End If
        Next
        
    End If

End Sub
Private Sub DisplayOneTypePrototype(typedefindex As Long)
    Call DisplayOneTypeDefStart(typedefindex)
    Call DisplayNestedTypesPrototypes(typedefindex)
    
    '###;todod
    Call DisplayOneTypeDefEnd(typedefindex)

End Sub
Private Sub DisplayOneTypeDefEnd(typeindex As Long)
   '###;todod
    Dim dummy   As String
    If IsTypeNested(typeindex) = True Then
        dummy = dummy & CreateSpaces(spacesfornested)
    End If
    dummy = dummy & CreateSpaces(spacefornamespace)
    dummy = dummy & "} // end of class "
    Dim classname As String
    classname = NameReserved(GetString(TypeDefStruct(typeindex).name))
    dummy = dummy & classname
    Console.WriteLine (dummy)
    Dim namespacename  As String
    namespacename = NameReserved(GetString(TypeDefStruct(typeindex).nspace))
    Console.WriteLine ("")
    If namespacename <> "" Then
        Dim nspace1 As String
        Dim i As Long
        nspace1 = NameReserved(GetString(TypeDefStruct(typeindex).nspace))
        For i = typeindex + 1 To UBound(TypeDefStruct)
            If IsTypeNested(i) Then
                Exit For
            End If
        Next
        Dim nspace2 As String
        If i <> UBound(TypeDefStruct) Then
            nspace2 = NameReserved(GetString(TypeDefStruct(i).nspace))
        End If
        If nspace1 <> nspace2 Then
            If lasttypedisplayed = typeindex And Notprototype = True Then
                Console.WriteLine ("")
                Console.WriteLine ("// =============================================================")
                Console.WriteLine ("")
                Placedend = True
                Call DisplayCustomAttribute("TypeRef", 0, 2)
            End If
            If lasttypedisplayed = typeindex And Notprototype = True Then
                DisplayFinalCustomAttributes
            End If
            Console.WriteA ("}")
            Console.WriteLine (" // end of namespace " & namespacename)
            spacefornamespace = 0
            spacesforrest = 2
            Writenamespace = True
            Console.WriteLine ("")
        Else
            Writenamespace = False

        End If
    End If
    
End Sub
Private Sub DisplayCustomAttribute(tname As String, tabindex As Long, noofspaces As Long)

End Sub
Private Sub DisplayFinalCustomAttributes()

End Sub
Private Function DisplayNestedTypesPrototypes(typedefindex As Long)
    If tablepresent(41) = False Then
        Exit Function
    End If
    Dim i As Long
    For i = 0 To UBound(NestedClassStruct)
        If NestedClassStruct(i).enclosingclass = typedefindex Then
            spacesfornested = spacesfornested + 2
            DisplayOneTypePrototype (NestedClassStruct(i).nestedclass)
            spacesfornested = spacesfornested - 2
        End If
    Next
        
End Function
Private Sub DisplayOneTypeDefStart(typerow As Long)
    Dim namespacename As String
    namespacename = NameReserved(GetString(TypeDefStruct(typerow).nspace))
    If namespacename <> "" Then
        If Writenamespace = True Then
            Console.WriteLine (".namespace " & namespacename)
            Console.WriteLine ("{")
            spacefornamespace = 2
            spacesforrest = 4
        End If
    End If
    Dim typestring  As String
    If IsTypeNested(typerow) Then
        typestring = typestring & CreateSpaces(spacesfornested)
    End If
    typestring = typestring & CreateSpaces(spacefornamespace)
    typestring = typestring & ".class /*02" & typerow & "*/ "
    Dim attributeflags  As String
    attributeflags = GetTypeAttributeFlags(TypeDefStruct(typerow).flags, typerow)
    Console.WriteLine (typestring & " " & attributeflags & " " & NameReserved(GetString(TypeDefStruct(typerow).name)))
    Dim tablename   As String
    tablename = GetTypeDefOrRefTable(TypeDefStruct(typerow).cindex)
    Dim index As Long
    index = GetTypeDefOrRefValue(TypeDefStruct(typerow).cindex)
    Dim typeextends As String
    If tablename = "TypeRef" Then
        typeextends = DisplayTypeRefExtends(index)
    End If
    If tablename = "TypeDef" Then
        typeextends = GetNestedTypeAsString(index) & DisplayTypeDefExtends(index)
    End If
    If Len(typeextends) <> 0 Then
        typestring = ""
        If IsTypeNested(typerow) Then
            typestring = typestring & CreateSpaces(spacesfornested)
            typestring = typestring & CreateSpaces(spacefornamespace)
            typestring = typestring & "       extends " & typeextends
            Console.WriteLine (typestring)
        End If
    End If
    Dim interfacestring  As String
    interfacestring = DisplayAllInterfaces(typerow)
    If Len(interfacestring) <> 0 Then
        typestring = ""
        If IsTypeNested(typerow) Then
            typestring = typestring & CreateSpaces(spacesfornested)
            typestring = typestring & CreateSpaces(spacefornamespace)
            typestring = typestring & "       implements " & interfacestring
            Console.WriteA (typestring)
        End If
    End If
    typestring = ""
    If IsTypeNested(typerow) Then
        typestring = typestring & CreateSpaces(spacesfornested)
        typestring = typestring & CreateSpaces(spacefornamespace)
        typestring = typestring & "{"
        Console.WriteLine (typestring)
    End If
End Sub
Private Function DisplayAllInterfaces(typeindex As Long) As String
    Dim returnstring As String
    If tablepresent(9) = False Then
        DisplayAllInterfaces = ""
        Exit Function
    End If
    Dim i As Long
    For i = 1 To UBound(InterfaceImplStruct)
        If typeindex = InterfaceImplStruct(i).classindex Then
            Dim codedtablename As String
            Dim interfaceindex As Long
            Dim interfacename  As String
            codedtablename = GetTypeDefOrRefTable(InterfaceImplStruct(i).interfaceindex)
            interfaceindex = GetTypeDefOrRefValue(InterfaceImplStruct(i).interfaceindex)
            If codedtablename = "TypeRef" Then
                interfacename = DisplayTypeRefExtends(interfaceindex)
                If codedtablename = "TypeDef" Then
                    interfacename = GetNestedTypeAsString(interfaceindex) & DisplayTypeDefExtends(interfaceindex)
                    returnstring = returnstring & interfacename
                End If
                Dim nextclassindex  As Boolean
                If i = UBound(InterfaceImplStruct) - 1 Then
                    nextclassindex = False
                ElseIf typeindex <> InterfaceImplStruct(i + 1).classindex Then
                    nextclassindex = False
                Else
                    nextclassindex = True
                End If
                If nextclassindex = True Then
                    returnstring = returnstring & "," & vbCrLf & "                 " & CreateSpaces(spacefornamespace + spacesfornested)
                Else
                    returnstring = returnstring & vbCrLf
                End If
                
            End If
            
        End If
    Next
    DisplayAllInterfaces = returnstring
End Function
Private Function DisplayTypeDefExtends(typedefindex As Long) As String
    If typedefindex = 0 Then
        DisplayTypeDefExtends = ""
        Exit Function
    End If
    Dim name As String
    Dim returnstring As String
    name = NameReserved(GetString(TypeDefStruct(typedefindex).name))
    returnstring = NameReserved(GetString(TypeDefStruct(typedefindex).nspace))
    If Len(returnstring) <> 0 Then
        returnstring = returnstring & "."
        returnstring = returnstring & name & "/* 02" & typedefindex & " */"
        DisplayTypeDefExtends = returnstring
    End If
End Function
Private Function GetNestedTypeAsString(rowindex As Long) As String
    Dim netsedtypestring As String
    Dim namespaceandnameparent2 As String
    Dim namespaceandnameparent3 As String
    If IsTypeNested(rowindex) = True Then
        Dim rowindexparent  As Long
        rowindexparent = GetParentForNestedType(rowindex)
        If IsTypeNested(rowindexparent) = True Then
            Dim rowindexparentparent As Long
            rowindexparentparent = GetParentForNestedType(rowindexparent)
            If IsTypeNested(rowindexparentparent) = True Then
                Dim rowindexp3  As Long
                rowindexp3 = GetParentForNestedType(rowindexparentparent)
                Dim nameparent3 As String
                nameparent3 = NameReserved(GetString(TypeDefStruct(rowindexp3).name))
                namespaceandnameparent3 = NameReserved(GetString(TypeDefStruct(rowindexp3).nspace))
                If Len(namespaceandnameparent3) <> 0 Then
                    namespaceandnameparent3 = namespaceandnameparent3 & "."
                    namespaceandnameparent3 = namespaceandnameparent3 & nameparent3 & "/* 02" & rowindexp3 & " *//"

                End If
                Dim nameparent2  As String
                nameparent2 = NameReserved(GetString(TypeDefStruct(rowindexparentparent).name))
                namespaceandnameparent2 = NameReserved(GetString(TypeDefStruct(rowindexparentparent).nspace))
                If Len(namespaceandnameparent2) <> 0 Then
                    namespaceandnameparent2 = namespaceandnameparent2 & "."
                    namespaceandnameparent2 = namespaceandnameparent3 & namespaceandnameparent2 & nameparent2 & "/* 02" & rowindexparentparent & " *//"
                End If
                Dim nameparent1  As String
                    nameparent1 = NameReserved(GetString(TypeDefStruct(rowindexparent).name))
                    netsedtypestring = NameReserved(GetString(TypeDefStruct(rowindexparent).nspace))
                    If Len(netsedtypestring) <> 0 Then
                        netsedtypestring = netsedtypestring & "."
                        netsedtypestring = namespaceandnameparent2 & netsedtypestring + nameparent1 & "/* 02" & rowindexparent & " *//"
                    End If
            End If
        End If
    End If
    GetNestedTypeAsString = netsedtypestring
End Function
Private Sub DisplayTypeDefsAndMethods()
    Notprototype = True
    If UBound(TypeDefStruct) <> 2 Then
        Console.WriteLine ("")
        Console.WriteLine ("// =============================================================")
        Console.WriteLine ("")
    End If
    Console.WriteLine ("")
    Console.WriteLine ("// =============== GLOBAL FIELDS AND METHODS ===================")
    Console.WriteLine ("")
    If UBound(TypeDefStruct) <> 2 Then
        Console.WriteLine ("")
        Console.WriteLine ("// =============================================================")
        Console.WriteLine ("")
        Console.WriteLine ("")
        Console.WriteLine ("// =============== CLASS MEMBERS DECLARATION ===================")
        Console.WriteLine ("//   note that class flags, 'extends' and 'implements' clauses")
        Console.WriteLine ("//          are provided here for information only")
        Console.WriteLine ("")
        Dim kk As Long
        kk = UBound(TypeDefStruct)
        Dim i As Long
        For i = 0 To kk
            If GetString(TypeDefStruct(i).name) = "_Deleted" And gVBNetStreamHeaders(0).rcName = "#-" Then
                If IsTypeNested(i) = False Then
                    DisplayOneType (i)
                End If
            End If
        Next
    End If
    Call DisplayEnd
End Sub
Private Sub DisplayOneType(typedefindex As Long)
    DisplayOneTypeDefStart (typedefindex)
  '  DisplayNestedTypes (typedefindex)
    DisplayOneTypeDefEnd (typedefindex)

End Sub
Private Sub DisplayEnd()
    Dim nspace As String
    nspace = NameReserved(GetString(TypeDefStruct(UBound(TypeDefStruct) - 1).nspace))
    If Placedend = False Then
        Console.WriteLine ("")
        Console.WriteLine ("// =============================================================")
        Console.WriteLine ("")
        Placedend = True
        Console.WriteLine ("//*********** DISASSEMBLY COMPLETE ***********************")
    End If
    
End Sub
Private Sub DisplayGlobalFields()
    Dim startofnext As Long
    Dim Start As Long
    startofnext = 0
    If tablepresent(4) = True Or tablepresent(2) = True Then
        Exit Sub
    End If
    Start = TypeDefStruct(1).findex
    If UBound(TypeDefStruct) = 2 Then
        startofnext = UBound(FieldStruct)
    Else
        startofnext = TypeDefStruct(2).findex
    End If
    If Start <> startofnext Then
        Console.WriteLine ("//Global fields")
        Console.WriteLine ("//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        DisplayAllFields (1)

    End If
End Sub
Private Sub DisplayAllFields(typeindex As Long)

End Sub
Private Sub DisplayAllMethods(typeindex As Long)
    If tablepresent(2) = False Or tablepresent(6) = False Then
        Exit Sub
    End If
    
End Sub
Private Sub DisplayGlobalMethods()
    If tablepresent(2) = False Or tablepresent(6) = False Then
        Exit Sub
    End If
    Dim Start As Long
    Dim startofnext As Long
    startofnext = 0
    Start = TypeDefStruct(1).mindex
    If UBound(TypeDefStruct) = 2 Then
        startofnext = UBound(MethodStruct)
    Else
        startofnext = TypeDefStruct(2).findex
    End If
    If Start <> startofnext Then
        Console.WriteLine ("//Global methods")
        Console.WriteLine ("//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        spacesforrest = 0
        DisplayAllMethods (1)
        spacesforrest = 2

    End If

End Sub
Private Function NameReserved(name As String) As String
    If Len(name) = 0 Then
        NameReserved = name
    ElseIf Asc(Mid$(name, 1, 1)) = 7 Then
        NameReserved = "'\\a'"
    ElseIf Asc(Mid$(name, 1, 1)) = 8 Then
        NameReserved = "'\\b'"
    ElseIf Asc(Mid$(name, 1, 1)) = 9 Then
        NameReserved = "'\\t'"
    ElseIf Asc(Mid$(name, 1, 1)) = 10 Then
        NameReserved = "'\\n'"
    ElseIf Asc(Mid$(name, 1, 1)) = 11 Then
        NameReserved = "'\\v'"
    ElseIf Asc(Mid$(name, 1, 1)) = 12 Then
        NameReserved = "'\\f'"
    ElseIf Asc(Mid$(name, 1, 1)) = 13 Then
        NameReserved = "'\\r'"
    ElseIf Asc(Mid$(name, 1, 1)) = 32 Then
        NameReserved = "' '"
    ElseIf Mid$(name, 1, 1) = "'" Then
        NameReserved = "'\\''"
    ElseIf Mid$(name, 1, 1) = "\" Then
        NameReserved = "'\\\"
    ElseIf Len(name) = 2 And Asc(Mid$(name, 2, 1)) = 7 Then
        NameReserved = "'" & Mid$(name, 1, 1) & "\\a'"
    ElseIf Len(name) = 2 And Asc(Mid$(name, 2, 1)) = 8 Then
        NameReserved = "'" & Mid$(name, 1, 1) & "\\b'"
    ElseIf Len(name) = 2 And Asc(Mid$(name, 2, 1)) = 9 Then
        NameReserved = "'" & Mid$(name, 1, 1) & "\\t'"
    ElseIf Len(name) = 2 And Asc(Mid$(name, 2, 1)) = 10 Then
        NameReserved = "'" & Mid$(name, 1, 1) & "\\n'"
    ElseIf Len(name) = 2 And Asc(Mid$(name, 2, 1)) = 11 Then
        NameReserved = "'" & Mid$(name, 1, 1) & "\\v'"
    ElseIf Len(name) = 2 And Asc(Mid$(name, 2, 1)) = 12 Then
        NameReserved = "'" & Mid$(name, 1, 1) & "\\f'"
    ElseIf Len(name) = 2 And Asc(Mid$(name, 2, 1)) = 13 Then
        NameReserved = "'" & Mid$(name, 1, 1) & "\\r'"
        
        NameReserved = name
    End If
End Function
Private Function DisplayTypeRefExtends(typerefindex As Long) As String
    Dim returnstring  As String
    Dim resolutionscope As Long
    Dim Resolutionscopetable As String
    Dim resolutionscopeindex As Long
    Dim dummy As String
    resolutionscope = TypeRefStruct(typerefindex).resolutionscope
    Resolutionscopetable = GetResolutionScopeTable(resolutionscope)
    resolutionscopeindex = GetResolutionScopeValue(resolutionscope)
    If Resolutionscopetable = "Module" Then
    
    End If
    If Resolutionscopetable = "AssemblyRef" Then
        returnstring = "[" & NameReserved(GetString(AssemblyRefStruct(resolutionscopeindex).name))
        returnstring = returnstring & "/* 23" & resolutionscopeindex & " */]"

    End If
    If Resolutionscopetable = "ModuleRef" Then
        returnstring = "[.module " & NameReserved(GetString(ModuleRefStruct(resolutionscopeindex).name))
        returnstring = returnstring & "/* 1A" & resolutionscopeindex & " */]"
    End If
    If Resolutionscopetable = "TypeRef" Then
        Dim resolutionscopeindex1  As Long
        Dim Resolutionscopetable1  As String
        resolutionscopeindex1 = GetResolutionScopeValue(TypeRefStruct(resolutionscopeindex).resolutionscope)
        Resolutionscopetable1 = GetResolutionScopeTable(TypeRefStruct(resolutionscopeindex).resolutionscope)
        If Resolutionscopetable1 = "AssemblyRef" Then
            dummy = "[" & NameReserved(GetString(AssemblyRefStruct(resolutionscopeindex1).name)) & "/* 23" & resolutionscopeindex1 & " */]"
            Dim nspace1 As String
            nspace1 = NameReserved(GetString(TypeRefStruct(resolutionscopeindex).nspace))
            If nspace1 <> "" Then
                nspace1 = nspace1 & "."
                dummy = dummy & nspace1 & NameReserved(GetString(TypeRefStruct(resolutionscopeindex).name)) & "/* 01" & resolutionscopeindex & " *//"

            End If
        
        End If
    
    End If
    Dim namespaceindex   As Long
    Dim nspace As String
    namespaceindex = TypeRefStruct(typerefindex).nspace
    nspace = NameReserved(GetString(namespaceindex))
    returnstring = returnstring & nspace
    If Len(nspace) <> 0 Then
        returnstring = returnstring & "."
        Dim nameindex As Long
        nameindex = TypeRefStruct(typerefindex).name
        returnstring = dummy & returnstring & NameReserved(GetString(nameindex)) & "/* 01" & typerefindex & " */"

    End If
    DisplayTypeRefExtends = returnstring
End Function
Private Function GetTypeDefOrRefValue(codedvalue As Long) As Long
    GetTypeDefOrRefValue = DoRightBitShift(codedvalue, 2)
End Function

Private Function DecodeFirstByteofMethodSignature(firstbyte As Long, methodrow As Long) As String
    Dim returnstring As String
    If (firstbyte And &H20) = &H20 Then
        returnstring = "instance "
    End If
    If (firstbyte And &H40) = &H40 Then
        returnstring = "explicit instance "
    End If
    Dim firstbits As Long
    firstbits = firstbyte And &HF
    If firstbits = &H2 Then
        returnstring = returnstring & "unmanaged stdcall "
    End If
    If firstbits = &H3 Then
        returnstring = returnstring & "unmanaged thiscall "
    End If
    If firstbits = &H5 Then
        returnstring = returnstring & "unmanaged vararg "
    End If
    If firstbits = &H1 Then
        returnstring = returnstring & "unmanaged cdecl "
    End If
    If firstbits = &H4 Then
        returnstring = returnstring & "unmanaged fastcall "
    End If
    DecodeFirstByteofMethodSignature = returnstring
End Function
Private Function GetType(typebyte As Byte) As String
    Select Case typebyte
    
        Case 1
            GetType = "void"
        Case 2
            GetType = "bool"
        Case 3
            GetType = "char"
        Case 4
            GetType = "int8"
        Case 5
            GetType = "unsigned int8"
        Case 6
            GetType = "int16"
        Case 7
            GetType = "unsigned int16"
        Case 8
            GetType = "int32"
        Case 9
            GetType = "unsigned int32"
        Case 10
            GetType = "int64"
        Case 11
            GetType = "unsigned int64"
        Case 12
            GetType = "float32"
        Case 13
            GetType = "float64"
        Case 14
            GetType = "string"
        Case Else
            GetType = "unknown"
        
    End Select
End Function
Private Function GetPointerToken(index As Long, blobarray() As Byte, howmanybytes As Long) As String
    Dim returnstring   As String
    Dim howmanybytes2 As Long
    returnstring = GetElementType(index + 1, blobarray, howmanybytes2) & "*"
    howmanybytes = howmanybytes2 + 1
    GetPointerToken = returnstring
End Function
Private Function DisplayTablesForDebugging()
    Console.WriteLine ("")
    Console.WriteLine ("Module Table: Records " & UBound(ModuleStruct))
    Console.WriteLine ("Name=" & GetString(ModuleStruct(1).name) & " " & ModuleStruct(1).name)
    Console.WriteLine ("Generation=" & ModuleStruct(1).Generation & " Mvid=" & ModuleStruct(1).Mvid & " EncId=" & ModuleStruct(1).EncId & " EncBaseId=" & ModuleStruct(1).EncBaseId)
    Console.WriteLine ("")
    Console.WriteLine ("TypeRef Table: Records " & UBound(TypeRefStruct))
    Dim i As Long
    For i = 1 To UBound(TypeRefStruct) - 1
        Console.WriteLine ("Type " & i)
        Console.WriteLine ("Resolution Scope=" & GetResolutionScopeTable(TypeRefStruct(i).resolutionscope) & " " & GetResolutionScopeValue(TypeRefStruct(i).resolutionscope))
        Console.WriteLine ("NameSpace=" & GetString(TypeRefStruct(i).nspace) & " " & TypeRefStruct(i).nspace)
        Console.WriteLine ("Name=" & GetString(TypeRefStruct(i).name) & " " & TypeRefStruct(i).name)
    Next i
    Console.WriteLine ("")
    Console.WriteLine ("TypeDef Table:  Records " & UBound(TypeDefStruct))
    For i = 1 To UBound(TypeDefStruct) - 1
        Console.WriteLine ("Type " & i)
        Console.WriteLine ("Name=" & GetString(TypeDefStruct(i).name) & " " & TypeDefStruct(i).name)
        Console.WriteLine ("NameSpace=" & GetString(TypeDefStruct(i).nspace) & " " & TypeDefStruct(i).nspace)
        Console.WriteLine ("Field[" & TypeDefStruct(i).findex & "]")
        Console.WriteLine ("Method[" & TypeDefStruct(i).mindex & "]")
        
    Next i
    
End Function
Public Sub DisplayModuleAndMore()
    Console.WriteLine (".module " & NameReserved(GetString(ModuleStruct(1).name)))
    Console.WriteA ("// MVID: ")
    DisplayGuid (ModuleStruct(1).Mvid)
    Console.WriteLine ("")
    'Console.WriteLine (".imagebase " & ImageBase)
    'Console.WriteLine(".subsystem 0x{0}" , subsystem.ToString("X8"))
    'Console.WriteLine(".file alignment {0}" , filea)
    'Console.WriteLine(".corflags 0x{0}" , corflags.ToString("x8"))
    Console.WriteLine ("// Image base: 0x03000000")
    
End Sub
Public Sub DisplayGuid(GuidIndex As Long)

End Sub
Public Sub ShowCLRHeader(fxgEXEInfo As MSFlexGrid)
    fxgEXEInfo.ColWidth(0) = 2000
    fxgEXEInfo.TextArray(2) = "cb"
    fxgEXEInfo.TextArray(3) = gVBNETHeader.CB
    fxgEXEInfo.AddItem "MajorRuntimeVersion"
    fxgEXEInfo.TextArray(5) = gVBNETHeader.MajorRuntimeVersion
    fxgEXEInfo.AddItem "MinorRuntimeVersion"
    fxgEXEInfo.TextArray(7) = gVBNETHeader.MinorRuntimeVersion
    fxgEXEInfo.AddItem "MetaData Size"
    fxgEXEInfo.TextArray(9) = gVBNETHeader.MetaData.Size
    fxgEXEInfo.AddItem "MetaData VirtualAddress"
    fxgEXEInfo.TextArray(11) = GetPtrFromRVA2(gVBNETHeader.MetaData.VirtualAddress)
    fxgEXEInfo.AddItem "Flags"
    fxgEXEInfo.TextArray(13) = gVBNETHeader.flags
    fxgEXEInfo.AddItem "EntryPointToken"
    fxgEXEInfo.TextArray(15) = gVBNETHeader.EntryPointToken
    fxgEXEInfo.AddItem "Resources Size"
    fxgEXEInfo.TextArray(17) = gVBNETHeader.Resources.Size
    fxgEXEInfo.AddItem "Resources VirtualAddress"
    fxgEXEInfo.TextArray(19) = GetPtrFromRVA2(gVBNETHeader.Resources.VirtualAddress)
    fxgEXEInfo.AddItem "StrongNameSignature Size "
    fxgEXEInfo.TextArray(21) = gVBNETHeader.StrongNameSignature.Size
    fxgEXEInfo.AddItem "StrongNameSignature VirtualAddress"
    fxgEXEInfo.TextArray(23) = gVBNETHeader.StrongNameSignature.VirtualAddress
    fxgEXEInfo.AddItem "CodeManagerTable Size"
    fxgEXEInfo.TextArray(25) = gVBNETHeader.CodeManagerTable.Size
    fxgEXEInfo.AddItem "CodeManagerTable VirtualAddress"
    fxgEXEInfo.TextArray(27) = gVBNETHeader.CodeManagerTable.VirtualAddress
    fxgEXEInfo.AddItem "VTableFixups Size"
    fxgEXEInfo.TextArray(29) = gVBNETHeader.VTableFixups.Size
    fxgEXEInfo.AddItem "VTableFixups VirtualAddress"
    fxgEXEInfo.TextArray(31) = gVBNETHeader.VTableFixups.VirtualAddress
    fxgEXEInfo.AddItem "ExportAddressTableJumps Size"
    fxgEXEInfo.TextArray(33) = gVBNETHeader.ExportAddressTableJumps.Size
    fxgEXEInfo.AddItem "ExportAddressTableJumps.VirtualAddress"
    fxgEXEInfo.TextArray(35) = gVBNETHeader.ExportAddressTableJumps.VirtualAddress
    fxgEXEInfo.AddItem "ManagedNativeHeader Size"
    fxgEXEInfo.TextArray(37) = gVBNETHeader.ManagedNativeHeader.Size
    fxgEXEInfo.AddItem "ManagedNativeHeader VirtualAddress"
    fxgEXEInfo.TextArray(39) = gVBNETHeader.ManagedNativeHeader.VirtualAddress

End Sub
Public Sub ShowMetaDataHeader(fxgEXEInfo As MSFlexGrid)
    fxgEXEInfo.ColWidth(0) = 2000
    fxgEXEInfo.TextArray(2) = "Signature"
    fxgEXEInfo.TextArray(3) = gVBNETMetaData.lSignature
    fxgEXEInfo.AddItem "MajorVersion"
    fxgEXEInfo.TextArray(5) = gVBNETMetaData.iMajorVersion
    fxgEXEInfo.AddItem "MinorVersion"
    fxgEXEInfo.TextArray(7) = gVBNETMetaData.iMinorVersion
    fxgEXEInfo.AddItem "ExtraData"
    fxgEXEInfo.TextArray(9) = gVBNETMetaData.iExtraData
    fxgEXEInfo.AddItem "Length"
    fxgEXEInfo.TextArray(11) = gVBNETMetaData.iLength
    fxgEXEInfo.AddItem "VersionString"
    fxgEXEInfo.TextArray(13) = gVBNETMetaData.iVersionString
    fxgEXEInfo.AddItem "Flags"
    fxgEXEInfo.TextArray(15) = gVBNETMetaDataHeader.fFlags
    fxgEXEInfo.AddItem "Padding "
    fxgEXEInfo.TextArray(17) = gVBNETMetaDataHeader.Padding
    fxgEXEInfo.AddItem "Number of Streams"
    fxgEXEInfo.TextArray(19) = gVBNETMetaDataHeader.iStreams
End Sub
Private Function IsDotNetInstalled() As Boolean
    Dim ret As Long
    ret = LoadLibrary("mscoree.dll")
    If ret = 0 Then
        IsDotNetInstalled = False
    Else
        IsDotNetInstalled = True
    End If
End Function

'==========================================================================
'  FULL .NET DECOMPILER
'  ------------------------------------------------------------------------
'  Builds three views of the assembly from the metadata already parsed by
'  ReadTablesIntoStructures:
'     decompiled.il  - IL disassembly (types, fields, method signatures and
'                      full IL opcode listings for every method body)
'     decompiled.cs  - best-effort C# reconstruction
'     decompiled.vb  - best-effort VB.NET reconstruction
'  The C#/VB views emit accurate type/field/method signatures and a simple
'  stack-machine reconstruction of straight-line method bodies.  Anything
'  with real control flow is left as a comment pointing back at the IL.
'==========================================================================
Public Sub DecompileDotNet()
On Error GoTo done
    Dim ilOut As clsConsole, csOut As clsConsole, vbOut As clsConsole
    Set ilOut = New clsConsole
    Set csOut = New clsConsole
    Set vbOut = New clsConsole

    Dim fdec As Integer
    fdec = FreeFile
    Open SFilePath For Binary Access Read As #fdec

    ilOut.WriteLine "// ======================================================"
    ilOut.WriteLine "// Semi VB Decompiler - .NET IL Disassembly"
    ilOut.WriteLine "// Assembly: " & SFile
    ilOut.WriteLine "// ======================================================"
    ilOut.WriteLine ""
    csOut.WriteLine "// Semi VB Decompiler - reconstructed C# (best effort)"
    csOut.WriteLine "// Assembly: " & SFile
    csOut.WriteLine ""
    vbOut.WriteLine "' Semi VB Decompiler - reconstructed VB.NET (best effort)"
    vbOut.WriteLine "' Assembly: " & SFile
    vbOut.WriteLine ""

    Dim lastType As Long
    lastType = iRows(2)
    'Reset the per-type store that backs the project tree and solution builder
    gNetTypeCount = 0
    ReDim gNetTypeName(lastType + 1)
    ReDim gNetTypeCS(lastType + 1)
    ReDim gNetTypeVB(lastType + 1)
    ReDim gNetTypeIL(lastType + 1)
    ReDim gNetTypeMethods(lastType + 1)
    Dim t As Long
    For t = 1 To lastType
        Call EmitTypeAll(t, fdec, ilOut, csOut, vbOut)
    Next

    Close #fdec

    Dim outDir As String
    outDir = App.Path & "\dump\" & SFile & "\"
    ilOut.SaveConsoleToFile outDir & "decompiled.il"
    csOut.SaveConsoleToFile outDir & "decompiled.cs"
    vbOut.SaveConsoleToFile outDir & "decompiled.vb"

    If Not Console Is Nothing Then
        Console.WriteLine ""
        Console.WriteLine "// Decompiled output written to:"
        Console.WriteLine "//   " & outDir & "decompiled.il"
        Console.WriteLine "//   " & outDir & "decompiled.cs"
        Console.WriteLine "//   " & outDir & "decompiled.vb"
    End If
    Exit Sub
done:
    On Error Resume Next
    Close #fdec
    If Not Console Is Nothing Then Console.WriteLine "// DecompileDotNet error: " & err.Description
End Sub

'--- Type emission -------------------------------------------------------
Private Sub EmitTypeAll(t As Long, f As Integer, ilOut As clsConsole, csOut As clsConsole, vbOut As clsConsole)
On Error GoTo skip
    Dim nm As String, ns As String
    nm = GetString(TypeDefStruct(t).name)
    ns = GetString(TypeDefStruct(t).nspace)
    Dim isModule As Boolean
    isModule = (t = 1) And (nm = "<Module>")

    Dim fFirst As Long, fLast As Long, mFirst As Long, mLast As Long
    Call GetFieldRange(t, fFirst, fLast)
    Call GetMethodRange(t, mFirst, mLast)

    'Emit into per-type buffers so each class can be stored individually;
    'they are then appended to the whole-file outputs below.
    Dim ilT As clsConsole, csT As clsConsole, vbT As clsConsole
    Set ilT = New clsConsole
    Set csT = New clsConsole
    Set vbT = New clsConsole

    'IL class header
    ilT.WriteLine ".class " & GetTypeAttributeFlags(TypeDefStruct(t).flags, t) & FullName(ns, nm)
    Dim baseNm As String
    baseNm = GetExtendsName(t, LANG_IL)
    If Len(baseNm) > 0 Then ilT.WriteLine "       extends " & baseNm
    ilT.WriteLine "{"

    'C#/VB class header
    If Not isModule Then
        csT.WriteLine CsTypeHeader(t, ns, nm)
        csT.WriteLine "{"
        vbT.WriteLine VbTypeHeader(t, ns, nm)
    End If

    Dim i As Long
    For i = fFirst To fLast
        If i >= 1 And i <= iRows(4) Then Call EmitField(i, t, isModule, ilT, csT, vbT)
    Next
    For i = mFirst To mLast
        If i >= 1 And i <= iRows(6) Then Call EmitMethod(i, t, isModule, f, ilT, csT, vbT)
    Next

    ilT.WriteLine "} // end of class " & FullName(ns, nm)
    ilT.WriteLine ""
    If Not isModule Then
        csT.WriteLine "}"
        csT.WriteLine ""
        vbT.WriteLine "End " & TypeKind(t, LANG_VB)
        vbT.WriteLine ""
    End If

    'Append this type to the whole-file outputs
    ilOut.WriteA ilT.Text
    csOut.WriteA csT.Text
    vbOut.WriteA vbT.Text

    'Store this type for the project tree and the solution builder
    gNetTypeName(gNetTypeCount) = FullName(ns, nm)
    gNetTypeIL(gNetTypeCount) = ilT.Text
    gNetTypeCS(gNetTypeCount) = csT.Text
    gNetTypeVB(gNetTypeCount) = vbT.Text
    gNetTypeMethods(gNetTypeCount) = BuildMethodSigList(t)
    gNetTypeCount = gNetTypeCount + 1
    Exit Sub
skip:
End Sub

Private Sub GetFieldRange(t As Long, ByRef first As Long, ByRef last As Long)
    first = TypeDefStruct(t).findex
    If t < iRows(2) Then
        last = TypeDefStruct(t + 1).findex - 1
    Else
        last = iRows(4)
    End If
End Sub

Private Sub GetMethodRange(t As Long, ByRef first As Long, ByRef last As Long)
    first = TypeDefStruct(t).mindex
    If t < iRows(2) Then
        last = TypeDefStruct(t + 1).mindex - 1
    Else
        last = iRows(6)
    End If
End Sub

Private Function FullName(ns As String, nm As String) As String
    If Len(ns) > 0 Then
        FullName = ns & "." & nm
    Else
        FullName = nm
    End If
End Function

Private Function GetExtendsName(t As Long, lang As Integer) As String
    Dim coded As Long
    coded = TypeDefStruct(t).cindex
    If coded = 0 Then Exit Function
    GetExtendsName = GetTypeDefOrRefName(coded, lang)
End Function

'Resolve a 2-bit TypeDefOrRef coded index (as stored in metadata tables)
Private Function GetTypeDefOrRefName(coded As Long, lang As Integer) As String
    Dim tag As Long, row As Long
    tag = coded And 3
    row = coded \ 4
    Select Case tag
        Case 0: GetTypeDefOrRefName = GetTypeDefFullName(row, lang)
        Case 1: GetTypeDefOrRefName = GetTypeRefFullName(row, lang)
        Case Else: GetTypeDefOrRefName = "TypeSpec[" & row & "]"
    End Select
End Function

Private Function GetTypeDefFullName(row As Long, lang As Integer) As String
    On Error Resume Next
    If row < 1 Or row > iRows(2) Then GetTypeDefFullName = "object": Exit Function
    GetTypeDefFullName = MapSystemName(FullName(GetString(TypeDefStruct(row).nspace), GetString(TypeDefStruct(row).name)), lang)
End Function

Private Function GetTypeRefFullName(row As Long, lang As Integer) As String
    On Error Resume Next
    If row < 1 Or row > iRows(1) Then GetTypeRefFullName = "object": Exit Function
    GetTypeRefFullName = MapSystemName(FullName(GetString(TypeRefStruct(row).nspace), GetString(TypeRefStruct(row).name)), lang)
End Function

'Map common System.* names to language keywords; leave others untouched.
Private Function MapSystemName(full As String, lang As Integer) As String
    If lang = LANG_IL Then MapSystemName = full: Exit Function
    Dim cs As Boolean
    cs = (lang = LANG_CS)
    Select Case full
        Case "System.Void": MapSystemName = "void"
        Case "System.Object": MapSystemName = IIf(cs, "object", "Object")
        Case "System.String": MapSystemName = IIf(cs, "string", "String")
        Case "System.Boolean": MapSystemName = IIf(cs, "bool", "Boolean")
        Case "System.Char": MapSystemName = IIf(cs, "char", "Char")
        Case "System.SByte": MapSystemName = IIf(cs, "sbyte", "SByte")
        Case "System.Byte": MapSystemName = IIf(cs, "byte", "Byte")
        Case "System.Int16": MapSystemName = IIf(cs, "short", "Short")
        Case "System.UInt16": MapSystemName = IIf(cs, "ushort", "UShort")
        Case "System.Int32": MapSystemName = IIf(cs, "int", "Integer")
        Case "System.UInt32": MapSystemName = IIf(cs, "uint", "UInteger")
        Case "System.Int64": MapSystemName = IIf(cs, "long", "Long")
        Case "System.UInt64": MapSystemName = IIf(cs, "ulong", "ULong")
        Case "System.Single": MapSystemName = IIf(cs, "float", "Single")
        Case "System.Double": MapSystemName = IIf(cs, "double", "Double")
        Case "System.Decimal": MapSystemName = IIf(cs, "decimal", "Decimal")
        Case Else: MapSystemName = full
    End Select
End Function

Private Function TypeKind(t As Long, lang As Integer) As String
    Dim fl As Long
    fl = TypeDefStruct(t).flags
    Dim baseNm As String
    baseNm = GetExtendsName(t, LANG_IL)
    If (fl And &H20) <> 0 Then
        TypeKind = "Interface"
    ElseIf baseNm = "System.ValueType" Then
        TypeKind = "Structure"
    ElseIf baseNm = "System.Enum" Then
        TypeKind = "Enum"
    Else
        TypeKind = "Class"
    End If
    If lang = LANG_CS Then TypeKind = LCase$(TypeKind)
    If lang = LANG_CS And TypeKind = "structure" Then TypeKind = "struct"
End Function

Private Function CsTypeHeader(t As Long, ns As String, nm As String) As String
    Dim fl As Long
    fl = TypeDefStruct(t).flags
    Dim s As String
    s = CsVisibility(fl)
    Dim kind As String
    kind = TypeKind(t, LANG_CS)
    If kind = "class" Then
        If (fl And &H80) <> 0 And (fl And &H100) <> 0 Then
            s = s & "static "
        ElseIf (fl And &H80) <> 0 Then
            s = s & "abstract "
        ElseIf (fl And &H100) <> 0 Then
            s = s & "sealed "
        End If
    End If
    s = s & kind & " " & nm
    Dim baseNm As String
    baseNm = GetExtendsName(t, LANG_CS)
    If kind = "class" And baseNm <> "" And baseNm <> "object" Then
        s = s & " : " & baseNm
    End If
    CsTypeHeader = s
End Function

Private Function VbTypeHeader(t As Long, ns As String, nm As String) As String
    Dim fl As Long
    fl = TypeDefStruct(t).flags
    Dim s As String
    s = VbVisibility(fl)
    Dim kind As String
    kind = TypeKind(t, LANG_VB)
    If kind = "Class" Then
        If (fl And &H80) <> 0 And (fl And &H100) <> 0 Then
            s = s & "NotInheritable "
        ElseIf (fl And &H80) <> 0 Then
            s = s & "MustInherit "
        ElseIf (fl And &H100) <> 0 Then
            s = s & "NotInheritable "
        End If
    End If
    s = s & kind & " " & nm
    Dim baseNm As String
    baseNm = GetExtendsName(t, LANG_VB)
    If kind = "Class" And baseNm <> "" And baseNm <> "Object" Then
        s = s & vbCrLf & "    Inherits " & baseNm
    End If
    VbTypeHeader = s
End Function

Private Function CsVisibility(fl As Long) As String
    Select Case (fl And 7)
        Case 1, 2: CsVisibility = "public "
        Case 3: CsVisibility = "private "
        Case 4: CsVisibility = "protected "
        Case 6, 7: CsVisibility = "protected internal "
        Case Else: CsVisibility = "internal "
    End Select
End Function

Private Function VbVisibility(fl As Long) As String
    Select Case (fl And 7)
        Case 1, 2: VbVisibility = "Public "
        Case 3: VbVisibility = "Private "
        Case 4: VbVisibility = "Protected "
        Case 6, 7: VbVisibility = "Protected Friend "
        Case Else: VbVisibility = "Friend "
    End Select
End Function

'--- Field emission ------------------------------------------------------
Private Sub EmitField(fi As Long, t As Long, isModule As Boolean, ilOut As clsConsole, csOut As clsConsole, vbOut As clsConsole)
On Error GoTo skip
    Dim nm As String
    nm = GetString(FieldStruct(fi).name)
    Dim fl As Long
    fl = FieldStruct(fi).flags
    ilOut.WriteLine "    .field " & FieldAttrIL(fl) & ParseFieldSigLang(FieldStruct(fi).sig, LANG_IL) & " " & nm
    csOut.WriteLine "    " & CsFieldModifiers(fl) & ParseFieldSigLang(FieldStruct(fi).sig, LANG_CS) & " " & nm & ";"
    vbOut.WriteLine "    " & VbFieldModifiers(fl) & nm & " As " & ParseFieldSigLang(FieldStruct(fi).sig, LANG_VB)
    Exit Sub
skip:
End Sub

Private Function FieldAttrIL(fl As Long) As String
    Dim s As String
    Select Case (fl And 7)
        Case 1: s = "private "
        Case 2: s = "famandassem "
        Case 3: s = "assembly "
        Case 4: s = "family "
        Case 5: s = "famorassem "
        Case 6: s = "public "
        Case Else: s = "privatescope "
    End Select
    If (fl And &H10) <> 0 Then s = s & "static "
    If (fl And &H20) <> 0 Then s = s & "initonly "
    If (fl And &H40) <> 0 Then s = s & "literal "
    FieldAttrIL = s
End Function

Private Function CsFieldModifiers(fl As Long) As String
    Dim s As String
    s = CsVisibility(fl)
    If (fl And &H40) <> 0 Then
        s = s & "const "
    Else
        If (fl And &H10) <> 0 Then s = s & "static "
        If (fl And &H20) <> 0 Then s = s & "readonly "
    End If
    CsFieldModifiers = s
End Function

Private Function VbFieldModifiers(fl As Long) As String
    Dim s As String
    s = VbVisibility(fl)
    If (fl And &H40) <> 0 Then
        s = s & "Const "
    Else
        If (fl And &H10) <> 0 Then s = s & "Shared "
        If (fl And &H20) <> 0 Then s = s & "ReadOnly "
    End If
    VbFieldModifiers = s
End Function

'--- Method emission -----------------------------------------------------
Private Sub EmitMethod(mi As Long, t As Long, isModule As Boolean, f As Integer, ilOut As clsConsole, csOut As clsConsole, vbOut As clsConsole)
On Error GoTo skip
    Dim nm As String
    nm = GetString(MethodStruct(mi).name)
    Dim fl As Long
    fl = MethodStruct(mi).flags
    Dim isCtor As Boolean
    isCtor = (nm = ".ctor" Or nm = ".cctor")
    Dim typeName As String
    typeName = GetString(TypeDefStruct(t).name)

    Dim retIL As String, retCS As String, retVB As String
    Dim pIL() As String, pCS() As String, pVB() As String
    Dim pc As Long, hasThis As Boolean
    Call ParseMethodSigFull(MethodStruct(mi).signature, retIL, pIL, pc, hasThis, LANG_IL)
    Call ParseMethodSigFull(MethodStruct(mi).signature, retCS, pCS, pc, hasThis, LANG_CS)
    Call ParseMethodSigFull(MethodStruct(mi).signature, retVB, pVB, pc, hasThis, LANG_VB)

    Dim pName() As String
    ReDim pName(pc + 1)
    Call GetParamNames(mi, pc, pName)

    'IL signature
    Dim ilSig As String
    ilSig = "    .method " & MethodAttrIL(fl) & retIL & " " & nm & "(" & JoinParamsIL(pIL, pc) & ") cil managed"
    ilOut.WriteLine ilSig
    ilOut.WriteLine "    {"

    'C# signature
    Dim csName As String
    If isCtor Then csName = typeName Else csName = nm
    Dim csSig As String
    csSig = "    " & CsMethodModifiers(fl)
    If Not isCtor Then csSig = csSig & retCS & " "
    csSig = csSig & csName & "(" & JoinParamsNamed(pCS, pName, pc, LANG_CS) & ")"

    'VB signature
    Dim vbIsSub As Boolean
    vbIsSub = (retIL = "void") Or isCtor
    Dim vbName As String
    If isCtor Then vbName = "New" Else vbName = nm
    Dim vbSig As String
    vbSig = "    " & VbMethodModifiers(fl) & IIf(vbIsSub, "Sub ", "Function ") & vbName & "(" & JoinParamsNamed(pVB, pName, pc, LANG_VB) & ")"
    If Not vbIsSub Then vbSig = vbSig & " As " & retVB

    Dim rva As Long
    rva = MethodStruct(mi).rva
    If rva = 0 Then
        ilOut.WriteLine "        // no body (abstract / extern / pinvoke)"
        ilOut.WriteLine "    } // end of method " & nm
        ilOut.WriteLine ""
        csOut.WriteLine csSig & ";"
        csOut.WriteLine ""
        vbOut.WriteLine vbSig
        vbOut.WriteLine "    End " & IIf(vbIsSub, "Sub", "Function")
        vbOut.WriteLine ""
        Exit Sub
    End If

    Dim ilBytes() As Byte, codeSize As Long, localTok As Long
    If ReadMethodBodyBytes(f, rva, ilBytes, codeSize, localTok) Then
        Call DecodeInstructions(ilBytes, codeSize)
        'IL body
        Dim j As Long
        For j = 0 To gInsCount - 1
            ilOut.WriteLine "        " & FormatILLine(j)
        Next
        'C# body
        csOut.WriteLine csSig
        csOut.WriteLine "    {"
        csOut.WriteA ReconstructBody(mi, hasThis, pc, pName, retIL, localTok, LANG_CS)
        csOut.WriteLine "    }"
        csOut.WriteLine ""
        'VB body
        vbOut.WriteLine vbSig
        vbOut.WriteA ReconstructBody(mi, hasThis, pc, pName, retIL, localTok, LANG_VB)
        vbOut.WriteLine "    End " & IIf(vbIsSub, "Sub", "Function")
        vbOut.WriteLine ""
    Else
        ilOut.WriteLine "        // unable to read method body"
        csOut.WriteLine csSig & " { }"
        csOut.WriteLine ""
        vbOut.WriteLine vbSig
        vbOut.WriteLine "    End " & IIf(vbIsSub, "Sub", "Function")
        vbOut.WriteLine ""
    End If
    ilOut.WriteLine "    } // end of method " & nm
    ilOut.WriteLine ""
    Exit Sub
skip:
End Sub

Private Function MethodAttrIL(fl As Long) As String
    Dim s As String
    Select Case (fl And 7)
        Case 1: s = "private "
        Case 2: s = "famandassem "
        Case 3: s = "assembly "
        Case 4: s = "family "
        Case 5: s = "famorassem "
        Case 6: s = "public "
        Case Else: s = "privatescope "
    End Select
    If (fl And &H10) <> 0 Then s = s & "static "
    If (fl And &H20) <> 0 Then s = s & "final "
    If (fl And &H40) <> 0 Then s = s & "virtual "
    If (fl And &H80) <> 0 Then s = s & "hidebysig "
    If (fl And &H400) <> 0 Then s = s & "abstract "
    If (fl And &H800) <> 0 Then s = s & "specialname "
    MethodAttrIL = s
End Function

Private Function CsMethodModifiers(fl As Long) As String
    Dim s As String
    s = CsVisibility(fl)
    If (fl And &H10) <> 0 Then s = s & "static "
    If (fl And &H400) <> 0 Then
        s = s & "abstract "
    ElseIf (fl And &H40) <> 0 Then
        s = s & "virtual "
    End If
    CsMethodModifiers = s
End Function

Private Function VbMethodModifiers(fl As Long) As String
    Dim s As String
    s = VbVisibility(fl)
    If (fl And &H10) <> 0 Then s = s & "Shared "
    If (fl And &H400) <> 0 Then
        s = s & "MustOverride "
    ElseIf (fl And &H40) <> 0 Then
        s = s & "Overridable "
    End If
    VbMethodModifiers = s
End Function

Private Function JoinParamsIL(p() As String, pc As Long) As String
    Dim s As String, k As Long
    For k = 1 To pc
        If k > 1 Then s = s & ", "
        s = s & p(k)
    Next
    JoinParamsIL = s
End Function

Private Function JoinParamsNamed(p() As String, pName() As String, pc As Long, lang As Integer) As String
    Dim s As String, k As Long
    For k = 1 To pc
        If k > 1 Then s = s & ", "
        If lang = LANG_VB Then
            s = s & "ByVal " & pName(k) & " As " & p(k)
        Else
            s = s & p(k) & " " & pName(k)
        End If
    Next
    JoinParamsNamed = s
End Function

Private Sub GetParamNames(mi As Long, pc As Long, ByRef pName() As String)
    Dim k As Long
    For k = 1 To pc
        pName(k) = "A" & k
    Next
    If iRows(8) = 0 Then Exit Sub
    Dim pFirst As Long, pLast As Long
    pFirst = MethodStruct(mi).param
    If mi < iRows(6) Then
        pLast = MethodStruct(mi + 1).param - 1
    Else
        pLast = iRows(8)
    End If
    For k = pFirst To pLast
        If k >= 1 And k <= iRows(8) Then
            Dim seq As Long
            seq = ParamStruct(k).sequence
            If seq >= 1 And seq <= pc Then
                Dim n As String
                n = GetString(ParamStruct(k).name)
                If Len(n) > 0 Then pName(seq) = n
            End If
        End If
    Next
End Sub

'--- Signature blob decoding (ECMA-335 II.23.2) --------------------------
Private Function ReadCompressedUInt(arr() As Byte, ByRef pos As Long) As Long
    Dim b As Long
    b = arr(pos)
    If (b And &H80) = 0 Then
        ReadCompressedUInt = b
        pos = pos + 1
    ElseIf (b And &HC0) = &H80 Then
        ReadCompressedUInt = ((b And &H3F) * 256&) + arr(pos + 1)
        pos = pos + 2
    Else
        ReadCompressedUInt = ((b And &H1F) * 16777216) + (CLng(arr(pos + 1)) * 65536) + (CLng(arr(pos + 2)) * 256) + arr(pos + 3)
        pos = pos + 4
    End If
End Function

'Position the cursor at the start of blob data, skipping the length prefix.
Private Function BlobDataStart(blobIndex As Long, ByRef pos As Long) As Boolean
    If blobIndex <= 0 Then BlobDataStart = False: Exit Function
    On Error GoTo bad
    pos = blobIndex
    Dim length As Long
    length = ReadCompressedUInt(BlobByteArray, pos)
    BlobDataStart = True
    Exit Function
bad:
    BlobDataStart = False
End Function

Private Function ParseFieldSigLang(blobIndex As Long, lang As Integer) As String
    On Error GoTo bad
    Dim pos As Long
    If Not BlobDataStart(blobIndex, pos) Then ParseFieldSigLang = IIf(lang = LANG_VB, "Object", "object"): Exit Function
    'first byte is the FIELD calling convention (0x06)
    pos = pos + 1
    ParseFieldSigLang = ParseTypeSig(BlobByteArray, pos, lang)
    Exit Function
bad:
    ParseFieldSigLang = IIf(lang = LANG_VB, "Object", "object")
End Function

Private Sub ParseMethodSigFull(blobIndex As Long, ByRef ret As String, ByRef params() As String, ByRef pc As Long, ByRef hasThis As Boolean, lang As Integer)
    On Error GoTo bad
    pc = 0
    ReDim params(0)
    hasThis = False
    ret = "void"
    Dim pos As Long
    If Not BlobDataStart(blobIndex, pos) Then Exit Sub
    Dim first As Long
    first = BlobByteArray(pos)
    pos = pos + 1
    hasThis = (first And &H20) <> 0
    If (first And &H10) <> 0 Then
        Dim gp As Long
        gp = ReadCompressedUInt(BlobByteArray, pos)   'generic param count
    End If
    pc = ReadCompressedUInt(BlobByteArray, pos)
    ret = ParseTypeSig(BlobByteArray, pos, lang)
    ReDim params(pc + 1)
    Dim k As Long
    For k = 1 To pc
        params(k) = ParseTypeSig(BlobByteArray, pos, lang)
    Next
    Exit Sub
bad:
End Sub

Private Sub GetSigCounts(blobIndex As Long, ByRef argc As Long, ByRef hasThis As Boolean, ByRef returnsVal As Boolean)
    Dim ret As String
    Dim p() As String
    Dim pc As Long
    Call ParseMethodSigFull(blobIndex, ret, p, pc, hasThis, LANG_IL)
    argc = pc
    returnsVal = (ret <> "void")
End Sub

'Recursive type-signature reader.  Returns a name in the requested language.
Private Function ParseTypeSig(arr() As Byte, ByRef pos As Long, lang As Integer) As String
    On Error GoTo bad
    Dim e As Long
    e = arr(pos)
    pos = pos + 1
    Select Case e
        Case &H1 To &HE, &H18, &H19, &H1C
            ParseTypeSig = PrimName(e, lang)
        Case &HF      'PTR
            ParseTypeSig = ParseTypeSig(arr, pos, lang) & "*"
        Case &H10     'BYREF
            If lang = LANG_CS Then
                ParseTypeSig = "ref " & ParseTypeSig(arr, pos, lang)
            ElseIf lang = LANG_VB Then
                ParseTypeSig = ParseTypeSig(arr, pos, lang)
            Else
                ParseTypeSig = ParseTypeSig(arr, pos, lang) & "&"
            End If
        Case &H11, &H12   'VALUETYPE / CLASS
            Dim tok As Long
            tok = ReadCompressedUInt(arr, pos)
            ParseTypeSig = GetTypeDefOrRefName(tok, lang)
        Case &H13     'VAR (generic type param)
            ParseTypeSig = "T" & ReadCompressedUInt(arr, pos)
        Case &H1E     'MVAR (generic method param)
            ParseTypeSig = "T" & ReadCompressedUInt(arr, pos)
        Case &H1D     'SZARRAY
            ParseTypeSig = ParseTypeSig(arr, pos, lang) & IIf(lang = LANG_VB, "()", "[]")
        Case &H14     'ARRAY
            Dim el As String
            el = ParseTypeSig(arr, pos, lang)
            Dim rank As Long, nSizes As Long, nLo As Long, z As Long
            rank = ReadCompressedUInt(arr, pos)
            nSizes = ReadCompressedUInt(arr, pos)
            For z = 1 To nSizes
                Call ReadCompressedUInt(arr, pos)
            Next
            nLo = ReadCompressedUInt(arr, pos)
            For z = 1 To nLo
                Call ReadCompressedUInt(arr, pos)
            Next
            ParseTypeSig = el & IIf(lang = LANG_VB, "()", "[]")
        Case &H15     'GENERICINST
            pos = pos + 1   'skip the CLASS/VALUETYPE indicator
            Dim gtok As Long
            gtok = ReadCompressedUInt(arr, pos)
            Dim baseNm As String
            baseNm = GetTypeDefOrRefName(gtok, lang)
            Dim ac As Long, ga As String, w As Long
            ac = ReadCompressedUInt(arr, pos)
            For w = 1 To ac
                If w > 1 Then ga = ga & ", "
                ga = ga & ParseTypeSig(arr, pos, lang)
            Next
            If lang = LANG_VB Then
                'strip the `n arity marker if present
                ParseTypeSig = StripArity(baseNm) & "(Of " & ga & ")"
            ElseIf lang = LANG_CS Then
                ParseTypeSig = StripArity(baseNm) & "<" & ga & ">"
            Else
                ParseTypeSig = baseNm & "<" & ga & ">"
            End If
        Case &H1F, &H20   'CMOD_REQD / CMOD_OPT
            Call ReadCompressedUInt(arr, pos)   'skip modifier token
            ParseTypeSig = ParseTypeSig(arr, pos, lang)
        Case &H45     'PINNED
            ParseTypeSig = ParseTypeSig(arr, pos, lang)
        Case &H16     'TYPEDBYREF
            ParseTypeSig = "TypedReference"
        Case Else
            ParseTypeSig = IIf(lang = LANG_VB, "Object", "object")
    End Select
    Exit Function
bad:
    ParseTypeSig = IIf(lang = LANG_VB, "Object", "object")
End Function

Private Function StripArity(nm As String) As String
    Dim p As Long
    p = InStr(nm, "`")
    If p > 0 Then
        StripArity = Left$(nm, p - 1)
    Else
        StripArity = nm
    End If
End Function

Private Function PrimName(e As Long, lang As Integer) As String
    Dim cs As Boolean
    cs = (lang = LANG_CS)
    Dim il As Boolean
    il = (lang = LANG_IL)
    Select Case e
        Case &H1: PrimName = "void"
        Case &H2: PrimName = IIf(il, "bool", IIf(cs, "bool", "Boolean"))
        Case &H3: PrimName = IIf(il, "char", IIf(cs, "char", "Char"))
        Case &H4: PrimName = IIf(il, "int8", IIf(cs, "sbyte", "SByte"))
        Case &H5: PrimName = IIf(il, "unsigned int8", IIf(cs, "byte", "Byte"))
        Case &H6: PrimName = IIf(il, "int16", IIf(cs, "short", "Short"))
        Case &H7: PrimName = IIf(il, "unsigned int16", IIf(cs, "ushort", "UShort"))
        Case &H8: PrimName = IIf(il, "int32", IIf(cs, "int", "Integer"))
        Case &H9: PrimName = IIf(il, "unsigned int32", IIf(cs, "uint", "UInteger"))
        Case &HA: PrimName = IIf(il, "int64", IIf(cs, "long", "Long"))
        Case &HB: PrimName = IIf(il, "unsigned int64", IIf(cs, "ulong", "ULong"))
        Case &HC: PrimName = IIf(il, "float32", IIf(cs, "float", "Single"))
        Case &HD: PrimName = IIf(il, "float64", IIf(cs, "double", "Double"))
        Case &HE: PrimName = IIf(il, "string", IIf(cs, "string", "String"))
        Case &H18: PrimName = IIf(il, "native int", "IntPtr")
        Case &H19: PrimName = IIf(il, "native unsigned int", "UIntPtr")
        Case &H1C: PrimName = IIf(il, "object", IIf(cs, "object", "Object"))
        Case Else: PrimName = IIf(cs, "object", IIf(il, "object", "Object"))
    End Select
End Function

'--- Method body reading (ECMA-335 II.25.4) ------------------------------
Private Function ReadMethodBodyBytes(f As Integer, rva As Long, ByRef ilBytes() As Byte, ByRef codeSize As Long, ByRef localSigTok As Long) As Boolean
    On Error GoTo bad
    localSigTok = 0
    codeSize = 0
    gEHCount = 0
    If rva = 0 Then Exit Function
    Dim foff As Long
    Dim moreSects As Boolean
    moreSects = False
    foff = GetPtrFromRVA2(rva)
    Seek #f, foff + 1
    Dim b0 As Byte
    Get #f, , b0
    If (b0 And 3) = 2 Then
        'Tiny format: code size is in the top 6 bits, IL follows immediately.
        codeSize = b0 \ 4
    ElseIf (b0 And 3) = 3 Then
        'Fat format: 12-byte header.  Bit 3 of the flags = more sections (EH).
        moreSects = (b0 And &H8) <> 0
        Dim b1 As Byte
        Get #f, , b1
        Dim maxStack As Integer
        Get #f, , maxStack
        Dim cs As Long
        Get #f, , cs
        Get #f, , localSigTok
        codeSize = cs
    Else
        Exit Function
    End If
    If codeSize <= 0 Then
        ReDim ilBytes(0)
        ReadMethodBodyBytes = True
        Exit Function
    End If
    ReDim ilBytes(codeSize - 1)
    Get #f, , ilBytes
    If moreSects Then Call ReadEHClauses(f, foff, codeSize)
    ReadMethodBodyBytes = True
    Exit Function
bad:
    ReadMethodBodyBytes = False
End Function

'--- IL disassembly ------------------------------------------------------
Private Sub DecodeInstructions(il() As Byte, codeSize As Long)
    gInsCount = 0
    ReDim gInsPos(255)
    ReDim gInsName(255)
    ReDim gInsText(255)
    ReDim gInsVal(255)
    ReDim gSwitchTargets(255)
    Dim p As Long
    p = 0
    Do While p < codeSize
        If gInsCount > UBound(gInsPos) Then
            ReDim Preserve gInsPos(gInsCount + 256)
            ReDim Preserve gInsName(gInsCount + 256)
            ReDim Preserve gInsText(gInsCount + 256)
            ReDim Preserve gInsVal(gInsCount + 256)
            ReDim Preserve gSwitchTargets(gInsCount + 256)
        End If
        Dim startP As Long
        Dim curSwitch As String
        curSwitch = ""
        startP = p
        Dim full As Integer
        Dim b As Integer
        b = il(p): p = p + 1
        If b = &HFE Then
            Dim b2 As Integer
            b2 = il(p): p = p + 1
            full = 256 + b2
        Else
            full = b
        End If
        Dim nm As String, kind As Integer
        Call LookupOpcode(full, nm, kind)
        Dim txt As String, val As Long
        txt = "": val = 0
        Select Case kind
            Case OK_NONE
            Case OK_I1
                val = SByteVal(il(p)): p = p + 1: txt = CStr(val)
            Case OK_U1
                val = il(p): p = p + 1: txt = CStr(val)
            Case OK_VAR
                val = il(p) + il(p + 1) * 256&: p = p + 2: txt = CStr(val)
            Case OK_I4
                val = BitConverterToInt32(il, p): p = p + 4: txt = CStr(val)
            Case OK_I8
                val = BitConverterToInt32(il, p): p = p + 8: txt = CStr(val)
            Case OK_R4
                txt = "0x" & RawHex(il, p, 4): p = p + 4
            Case OK_R8
                txt = "0x" & RawHex(il, p, 8): p = p + 8
            Case OK_BR1
                Dim d1 As Long
                d1 = SByteVal(il(p)): p = p + 1: val = p + d1: txt = "IL_" & Hex4(val)
            Case OK_BR4
                Dim d4 As Long
                d4 = BitConverterToInt32(il, p): p = p + 4: val = p + d4: txt = "IL_" & Hex4(val)
            Case OK_TOK
                val = BitConverterToInt32(il, p): p = p + 4: txt = GetTokenDisplay(val, LANG_IL)
            Case OK_STR
                val = BitConverterToInt32(il, p): p = p + 4: txt = GetUserStringByToken(val)
            Case OK_SWITCH
                Dim n As Long, s As Long, afterSw As Long, tgt As Long
                n = BitConverterToInt32(il, p): p = p + 4
                afterSw = p + n * 4
                For s = 0 To n - 1
                    tgt = afterSw + BitConverterToInt32(il, p)
                    p = p + 4
                    If s > 0 Then curSwitch = curSwitch & ","
                    curSwitch = curSwitch & tgt
                Next
                val = n
                txt = "(" & n & " targets)"
        End Select
        gInsPos(gInsCount) = startP
        gInsName(gInsCount) = nm
        gInsText(gInsCount) = txt
        gInsVal(gInsCount) = val
        gSwitchTargets(gInsCount) = curSwitch
        gInsCount = gInsCount + 1
    Loop
End Sub

Private Function FormatILLine(j As Long) As String
    Dim s As String
    s = "IL_" & Hex4(gInsPos(j)) & ":  " & gInsName(j)
    If Len(gInsText(j)) > 0 Then s = s & " " & gInsText(j)
    FormatILLine = s
End Function

Private Function SByteVal(b As Byte) As Long
    If b < 128 Then
        SByteVal = b
    Else
        SByteVal = CLng(b) - 256
    End If
End Function

Private Function Hex4(n As Long) As String
    Hex4 = Right$("0000" & Hex$(n And &HFFFF&), 4)
End Function

Private Function RawHex(arr() As Byte, ByVal pos As Long, ByVal n As Long) As String
    Dim s As String, i As Long
    For i = n - 1 To 0 Step -1
        s = s & Right$("0" & Hex$(arr(pos + i)), 2)
    Next
    RawHex = s
End Function

'Opcode table (ECMA-335 Partition III).  full < 256 is a one-byte opcode;
'full >= 256 is a two-byte 0xFE-prefixed opcode (full = 256 + secondByte).
Private Sub LookupOpcode(full As Integer, ByRef nm As String, ByRef kind As Integer)
    kind = OK_NONE
    Select Case full
        Case &H0: nm = "nop"
        Case &H1: nm = "break"
        Case &H2: nm = "ldarg.0"
        Case &H3: nm = "ldarg.1"
        Case &H4: nm = "ldarg.2"
        Case &H5: nm = "ldarg.3"
        Case &H6: nm = "ldloc.0"
        Case &H7: nm = "ldloc.1"
        Case &H8: nm = "ldloc.2"
        Case &H9: nm = "ldloc.3"
        Case &HA: nm = "stloc.0"
        Case &HB: nm = "stloc.1"
        Case &HC: nm = "stloc.2"
        Case &HD: nm = "stloc.3"
        Case &HE: nm = "ldarg.s": kind = OK_U1
        Case &HF: nm = "ldarga.s": kind = OK_U1
        Case &H10: nm = "starg.s": kind = OK_U1
        Case &H11: nm = "ldloc.s": kind = OK_U1
        Case &H12: nm = "ldloca.s": kind = OK_U1
        Case &H13: nm = "stloc.s": kind = OK_U1
        Case &H14: nm = "ldnull"
        Case &H15: nm = "ldc.i4.m1"
        Case &H16: nm = "ldc.i4.0"
        Case &H17: nm = "ldc.i4.1"
        Case &H18: nm = "ldc.i4.2"
        Case &H19: nm = "ldc.i4.3"
        Case &H1A: nm = "ldc.i4.4"
        Case &H1B: nm = "ldc.i4.5"
        Case &H1C: nm = "ldc.i4.6"
        Case &H1D: nm = "ldc.i4.7"
        Case &H1E: nm = "ldc.i4.8"
        Case &H1F: nm = "ldc.i4.s": kind = OK_I1
        Case &H20: nm = "ldc.i4": kind = OK_I4
        Case &H21: nm = "ldc.i8": kind = OK_I8
        Case &H22: nm = "ldc.r4": kind = OK_R4
        Case &H23: nm = "ldc.r8": kind = OK_R8
        Case &H25: nm = "dup"
        Case &H26: nm = "pop"
        Case &H27: nm = "jmp": kind = OK_TOK
        Case &H28: nm = "call": kind = OK_TOK
        Case &H29: nm = "calli": kind = OK_TOK
        Case &H2A: nm = "ret"
        Case &H2B: nm = "br.s": kind = OK_BR1
        Case &H2C: nm = "brfalse.s": kind = OK_BR1
        Case &H2D: nm = "brtrue.s": kind = OK_BR1
        Case &H2E: nm = "beq.s": kind = OK_BR1
        Case &H2F: nm = "bge.s": kind = OK_BR1
        Case &H30: nm = "bgt.s": kind = OK_BR1
        Case &H31: nm = "ble.s": kind = OK_BR1
        Case &H32: nm = "blt.s": kind = OK_BR1
        Case &H33: nm = "bne.un.s": kind = OK_BR1
        Case &H34: nm = "bge.un.s": kind = OK_BR1
        Case &H35: nm = "bgt.un.s": kind = OK_BR1
        Case &H36: nm = "ble.un.s": kind = OK_BR1
        Case &H37: nm = "blt.un.s": kind = OK_BR1
        Case &H38: nm = "br": kind = OK_BR4
        Case &H39: nm = "brfalse": kind = OK_BR4
        Case &H3A: nm = "brtrue": kind = OK_BR4
        Case &H3B: nm = "beq": kind = OK_BR4
        Case &H3C: nm = "bge": kind = OK_BR4
        Case &H3D: nm = "bgt": kind = OK_BR4
        Case &H3E: nm = "ble": kind = OK_BR4
        Case &H3F: nm = "blt": kind = OK_BR4
        Case &H40: nm = "bne.un": kind = OK_BR4
        Case &H41: nm = "bge.un": kind = OK_BR4
        Case &H42: nm = "bgt.un": kind = OK_BR4
        Case &H43: nm = "ble.un": kind = OK_BR4
        Case &H44: nm = "blt.un": kind = OK_BR4
        Case &H45: nm = "switch": kind = OK_SWITCH
        Case &H46: nm = "ldind.i1"
        Case &H47: nm = "ldind.u1"
        Case &H48: nm = "ldind.i2"
        Case &H49: nm = "ldind.u2"
        Case &H4A: nm = "ldind.i4"
        Case &H4B: nm = "ldind.u4"
        Case &H4C: nm = "ldind.i8"
        Case &H4D: nm = "ldind.i"
        Case &H4E: nm = "ldind.r4"
        Case &H4F: nm = "ldind.r8"
        Case &H50: nm = "ldind.ref"
        Case &H51: nm = "stind.ref"
        Case &H52: nm = "stind.i1"
        Case &H53: nm = "stind.i2"
        Case &H54: nm = "stind.i4"
        Case &H55: nm = "stind.i8"
        Case &H56: nm = "stind.r4"
        Case &H57: nm = "stind.r8"
        Case &H58: nm = "add"
        Case &H59: nm = "sub"
        Case &H5A: nm = "mul"
        Case &H5B: nm = "div"
        Case &H5C: nm = "div.un"
        Case &H5D: nm = "rem"
        Case &H5E: nm = "rem.un"
        Case &H5F: nm = "and"
        Case &H60: nm = "or"
        Case &H61: nm = "xor"
        Case &H62: nm = "shl"
        Case &H63: nm = "shr"
        Case &H64: nm = "shr.un"
        Case &H65: nm = "neg"
        Case &H66: nm = "not"
        Case &H67: nm = "conv.i1"
        Case &H68: nm = "conv.i2"
        Case &H69: nm = "conv.i4"
        Case &H6A: nm = "conv.i8"
        Case &H6B: nm = "conv.r4"
        Case &H6C: nm = "conv.r8"
        Case &H6D: nm = "conv.u4"
        Case &H6E: nm = "conv.u8"
        Case &H6F: nm = "callvirt": kind = OK_TOK
        Case &H70: nm = "cpobj": kind = OK_TOK
        Case &H71: nm = "ldobj": kind = OK_TOK
        Case &H72: nm = "ldstr": kind = OK_STR
        Case &H73: nm = "newobj": kind = OK_TOK
        Case &H74: nm = "castclass": kind = OK_TOK
        Case &H75: nm = "isinst": kind = OK_TOK
        Case &H76: nm = "conv.r.un"
        Case &H79: nm = "unbox": kind = OK_TOK
        Case &H7A: nm = "throw"
        Case &H7B: nm = "ldfld": kind = OK_TOK
        Case &H7C: nm = "ldflda": kind = OK_TOK
        Case &H7D: nm = "stfld": kind = OK_TOK
        Case &H7E: nm = "ldsfld": kind = OK_TOK
        Case &H7F: nm = "ldsflda": kind = OK_TOK
        Case &H80: nm = "stsfld": kind = OK_TOK
        Case &H81: nm = "stobj": kind = OK_TOK
        Case &H82: nm = "conv.ovf.i1.un"
        Case &H83: nm = "conv.ovf.i2.un"
        Case &H84: nm = "conv.ovf.i4.un"
        Case &H85: nm = "conv.ovf.i8.un"
        Case &H86: nm = "conv.ovf.u1.un"
        Case &H87: nm = "conv.ovf.u2.un"
        Case &H88: nm = "conv.ovf.u4.un"
        Case &H89: nm = "conv.ovf.u8.un"
        Case &H8A: nm = "conv.ovf.i.un"
        Case &H8B: nm = "conv.ovf.u.un"
        Case &H8C: nm = "box": kind = OK_TOK
        Case &H8D: nm = "newarr": kind = OK_TOK
        Case &H8E: nm = "ldlen"
        Case &H8F: nm = "ldelema": kind = OK_TOK
        Case &H90: nm = "ldelem.i1"
        Case &H91: nm = "ldelem.u1"
        Case &H92: nm = "ldelem.i2"
        Case &H93: nm = "ldelem.u2"
        Case &H94: nm = "ldelem.i4"
        Case &H95: nm = "ldelem.u4"
        Case &H96: nm = "ldelem.i8"
        Case &H97: nm = "ldelem.i"
        Case &H98: nm = "ldelem.r4"
        Case &H99: nm = "ldelem.r8"
        Case &H9A: nm = "ldelem.ref"
        Case &H9B: nm = "stelem.i"
        Case &H9C: nm = "stelem.i1"
        Case &H9D: nm = "stelem.i2"
        Case &H9E: nm = "stelem.i4"
        Case &H9F: nm = "stelem.i8"
        Case &HA0: nm = "stelem.r4"
        Case &HA1: nm = "stelem.r8"
        Case &HA2: nm = "stelem.ref"
        Case &HA3: nm = "ldelem": kind = OK_TOK
        Case &HA4: nm = "stelem": kind = OK_TOK
        Case &HA5: nm = "unbox.any": kind = OK_TOK
        Case &HB3: nm = "conv.ovf.i1"
        Case &HB4: nm = "conv.ovf.u1"
        Case &HB5: nm = "conv.ovf.i2"
        Case &HB6: nm = "conv.ovf.u2"
        Case &HB7: nm = "conv.ovf.i4"
        Case &HB8: nm = "conv.ovf.u4"
        Case &HB9: nm = "conv.ovf.i8"
        Case &HBA: nm = "conv.ovf.u8"
        Case &HC2: nm = "refanyval": kind = OK_TOK
        Case &HC3: nm = "ckfinite"
        Case &HC6: nm = "mkrefany": kind = OK_TOK
        Case &HD0: nm = "ldtoken": kind = OK_TOK
        Case &HD1: nm = "conv.u2"
        Case &HD2: nm = "conv.u1"
        Case &HD3: nm = "conv.i"
        Case &HD4: nm = "conv.ovf.i"
        Case &HD5: nm = "conv.ovf.u"
        Case &HD6: nm = "add.ovf"
        Case &HD7: nm = "add.ovf.un"
        Case &HD8: nm = "mul.ovf"
        Case &HD9: nm = "mul.ovf.un"
        Case &HDA: nm = "sub.ovf"
        Case &HDB: nm = "sub.ovf.un"
        Case &HDC: nm = "endfinally"
        Case &HDD: nm = "leave": kind = OK_BR4
        Case &HDE: nm = "leave.s": kind = OK_BR1
        Case &HDF: nm = "stind.i"
        Case &HE0: nm = "conv.u"
        'two-byte 0xFE opcodes
        Case 256 + &H0: nm = "arglist"
        Case 256 + &H1: nm = "ceq"
        Case 256 + &H2: nm = "cgt"
        Case 256 + &H3: nm = "cgt.un"
        Case 256 + &H4: nm = "clt"
        Case 256 + &H5: nm = "clt.un"
        Case 256 + &H6: nm = "ldftn": kind = OK_TOK
        Case 256 + &H7: nm = "ldvirtftn": kind = OK_TOK
        Case 256 + &H9: nm = "ldarg": kind = OK_VAR
        Case 256 + &HA: nm = "ldarga": kind = OK_VAR
        Case 256 + &HB: nm = "starg": kind = OK_VAR
        Case 256 + &HC: nm = "ldloc": kind = OK_VAR
        Case 256 + &HD: nm = "ldloca": kind = OK_VAR
        Case 256 + &HE: nm = "stloc": kind = OK_VAR
        Case 256 + &HF: nm = "localloc"
        Case 256 + &H11: nm = "endfilter"
        Case 256 + &H12: nm = "unaligned.": kind = OK_U1
        Case 256 + &H13: nm = "volatile."
        Case 256 + &H14: nm = "tail."
        Case 256 + &H15: nm = "initobj": kind = OK_TOK
        Case 256 + &H16: nm = "constrained.": kind = OK_TOK
        Case 256 + &H17: nm = "cpblk"
        Case 256 + &H18: nm = "initblk"
        Case 256 + &H1A: nm = "rethrow"
        Case 256 + &H1C: nm = "sizeof": kind = OK_TOK
        Case 256 + &H1D: nm = "refanytype"
        Case 256 + &H1E: nm = "readonly."
        Case Else
            nm = "unknown.0x" & Hex$(full)
    End Select
End Sub

'--- Token resolution ----------------------------------------------------
'Display string for a metadata token used by an IL operand.
Private Function GetTokenDisplay(token As Long, lang As Integer) As String
    On Error GoTo bad
    Dim table As Long, row As Long
    table = (token \ &H1000000) And &HFF
    row = token And &HFFFFFF
    Select Case table
        Case &H6     'MethodDef
            GetTokenDisplay = GetMethodOwnerTypeName(row, lang) & "::" & GetString(MethodStruct(row).name)
        Case &HA     'MemberRef
            GetTokenDisplay = GetMemberRefParentName(MemberRefStruct(row).clas, lang) & "::" & GetString(MemberRefStruct(row).name)
        Case &H1     'TypeRef
            GetTokenDisplay = GetTypeRefFullName(row, lang)
        Case &H2     'TypeDef
            GetTokenDisplay = GetTypeDefFullName(row, lang)
        Case &H4     'Field
            GetTokenDisplay = GetFieldOwnerTypeName(row, lang) & "::" & GetString(FieldStruct(row).name)
        Case &H1B    'TypeSpec
            GetTokenDisplay = "TypeSpec[" & row & "]"
        Case &H2B    'MethodSpec
            GetTokenDisplay = "MethodSpec[" & row & "]"
        Case Else
            GetTokenDisplay = "token_0x" & Hex$(token)
    End Select
    Exit Function
bad:
    GetTokenDisplay = "token_0x" & Hex$(token)
End Function

Private Function GetMethodOwnerTypeName(methodRow As Long, lang As Integer) As String
    On Error Resume Next
    Dim t As Long, mFirst As Long, mLast As Long
    For t = 1 To iRows(2)
        Call GetMethodRange(t, mFirst, mLast)
        If methodRow >= mFirst And methodRow <= mLast Then
            GetMethodOwnerTypeName = GetTypeDefFullName(t, lang)
            Exit Function
        End If
    Next
    GetMethodOwnerTypeName = "?"
End Function

Private Function GetFieldOwnerTypeName(fieldRow As Long, lang As Integer) As String
    On Error Resume Next
    Dim t As Long, fFirst As Long, fLast As Long
    For t = 1 To iRows(2)
        Call GetFieldRange(t, fFirst, fLast)
        If fieldRow >= fFirst And fieldRow <= fLast Then
            GetFieldOwnerTypeName = GetTypeDefFullName(t, lang)
            Exit Function
        End If
    Next
    GetFieldOwnerTypeName = "?"
End Function

'Resolve a 3-bit MemberRefParent coded index.
Private Function GetMemberRefParentName(coded As Long, lang As Integer) As String
    On Error Resume Next
    Dim tag As Long, row As Long
    tag = coded And 7
    row = coded \ 8
    Select Case tag
        Case 0: GetMemberRefParentName = GetTypeDefFullName(row, lang)
        Case 1: GetMemberRefParentName = GetTypeRefFullName(row, lang)
        Case 2: GetMemberRefParentName = GetString(ModuleRefStruct(row).name)
        Case 3: GetMemberRefParentName = GetMethodOwnerTypeName(row, lang)
        Case 4: GetMemberRefParentName = "TypeSpec[" & row & "]"
        Case Else: GetMemberRefParentName = "?"
    End Select
End Function

'Read a #US user string by its 0x70xxxxxx token.  Returns a quoted literal.
Private Function GetUserStringByToken(token As Long) As String
    On Error GoTo bad
    Dim offset As Long
    offset = token And &HFFFFFF
    If offset = 0 Then GetUserStringByToken = """""": Exit Function
    Dim pos As Long
    pos = offset
    Dim length As Long
    length = ReadCompressedUInt(USByteArray, pos)
    If length <= 1 Then GetUserStringByToken = """""": Exit Function
    Dim chars As Long
    chars = (length - 1) \ 2
    Dim s As String, i As Long, code As Long
    For i = 0 To chars - 1
        code = USByteArray(pos + i * 2) + USByteArray(pos + i * 2 + 1) * 256&
        Select Case code
            Case 34: s = s & "\"""
            Case 92: s = s & "\\"
            Case 13: s = s & "\r"
            Case 10: s = s & "\n"
            Case 9: s = s & "\t"
            Case Else
                If code >= 32 And code < 127 Then
                    s = s & Chr$(code)
                ElseIf code < 256 Then
                    s = s & Chr$(code)
                Else
                    s = s & ChrW$(code)
                End If
        End Select
    Next
    GetUserStringByToken = """" & s & """"
    Exit Function
bad:
    GetUserStringByToken = Chr$(34) & Chr$(34)
End Function

'--- Call/field token info for the reconstruction stack machine ----------
Private Function CallInfoFromToken(token As Long, ByRef name As String, ByRef owner As String, ByRef argc As Long, ByRef hasThisC As Boolean, ByRef returnsVal As Boolean, lang As Integer) As Boolean
    On Error GoTo bad
    Dim table As Long, row As Long
    table = (token \ &H1000000) And &HFF
    row = token And &HFFFFFF
    Dim sigBlob As Long
    Select Case table
        Case &H6     'MethodDef
            name = GetString(MethodStruct(row).name)
            owner = GetMethodOwnerTypeName(row, lang)
            sigBlob = MethodStruct(row).signature
        Case &HA     'MemberRef
            name = GetString(MemberRefStruct(row).name)
            owner = GetMemberRefParentName(MemberRefStruct(row).clas, lang)
            sigBlob = MemberRefStruct(row).sig
        Case Else
            CallInfoFromToken = False
            Exit Function
    End Select
    Call GetSigCounts(sigBlob, argc, hasThisC, returnsVal)
    CallInfoFromToken = True
    Exit Function
bad:
    CallInfoFromToken = False
End Function

Private Function FieldRefFromToken(token As Long, ByRef owner As String, ByRef name As String, lang As Integer) As Boolean
    On Error GoTo bad
    Dim table As Long, row As Long
    table = (token \ &H1000000) And &HFF
    row = token And &HFFFFFF
    Select Case table
        Case &H4     'Field
            name = GetString(FieldStruct(row).name)
            owner = GetFieldOwnerTypeName(row, lang)
        Case &HA     'MemberRef
            name = GetString(MemberRefStruct(row).name)
            owner = GetMemberRefParentName(MemberRefStruct(row).clas, lang)
        Case Else
            FieldRefFromToken = False
            Exit Function
    End Select
    FieldRefFromToken = True
    Exit Function
bad:
    FieldRefFromToken = False
End Function

'--- Straight-line body reconstruction -----------------------------------
'A tiny string-valued stack machine.  It models the common load/store/call
'patterns; the first time it meets real control flow (a branch/switch) or an
'opcode it does not model, it stops and leaves a note pointing at the IL.
Private Function ReconstructBody(mi As Long, hasThis As Boolean, pc As Long, pName() As String, retIL As String, ByVal localTok As Long, lang As Integer) As String
    On Error GoTo bad
    gLang = lang
    If lang = LANG_CS Then
        gTerm = ";": gThisKw = "this": gNullKw = "null": gNewKw = "new ": gRetKw = "return": gThrowKw = "throw": gEqOp = " == "
    Else
        gTerm = "": gThisKw = "Me": gNullKw = "Nothing": gNewKw = "New ": gRetKw = "Return": gThrowKw = "Throw": gEqOp = " = "
    End If

    'Collect every branch / switch / leave target offset so we know which
    'instructions need a label.
    Dim targets() As Long, tCount As Long
    ReDim targets(gInsCount + 8)
    tCount = 0
    Dim j As Long
    For j = 0 To gInsCount - 1
        Select Case gInsName(j)
            Case "br", "br.s", "brtrue", "brtrue.s", "brfalse", "brfalse.s", _
                 "beq", "beq.s", "bne.un", "bne.un.s", "bge", "bge.s", "bge.un", "bge.un.s", _
                 "bgt", "bgt.s", "bgt.un", "bgt.un.s", "ble", "ble.s", "ble.un", "ble.un.s", _
                 "blt", "blt.s", "blt.un", "blt.un.s", "leave", "leave.s"
                targets(tCount) = gInsVal(j): tCount = tCount + 1
            Case "switch"
                If Len(gSwitchTargets(j)) > 0 Then
                    Dim sw() As String, q As Long
                    sw = Split(gSwitchTargets(j), ",")
                    For q = 0 To UBound(sw)
                        If tCount > UBound(targets) - 2 Then ReDim Preserve targets(tCount + 16)
                        targets(tCount) = CLng(sw(q)): tCount = tCount + 1
                    Next
                End If
        End Select
        If tCount > UBound(targets) - 2 Then ReDim Preserve targets(tCount + 16)
    Next

    'Reset the structured statement list.
    gLCount = 0
    ReDim gLType(64): ReDim gLText(64): ReDim gLAlt(64)
    ReDim gLTarget(64): ReDim gLSwitch(64): ReDim gLOffset(64): ReDim gLDead(64)

    Dim stack() As String
    ReDim stack(512)
    Dim sp As Long
    sp = 0

    For j = 0 To gInsCount - 1
        If OffsetInTargets(targets, tCount, gInsPos(j)) Then
            sp = 0
            Call LAppend(LT_LABEL, "IL_" & Hex4(gInsPos(j)), "", 0, "", gInsPos(j))
        End If
        gCurOff = gInsPos(j)
        Dim op As String
        op = gInsName(j)
        Dim a1 As String, a2 As String
        Select Case op
            Case "nop", "break", "volatile.", "readonly.", "tail.", "constrained.", "unaligned."
                'prefixes / nops
            Case "ldstr": sp = sp + 1: stack(sp) = gInsText(j)
            Case "ldnull": sp = sp + 1: stack(sp) = gNullKw
            Case "ldc.i4.m1": sp = sp + 1: stack(sp) = "-1"
            Case "ldc.i4.0": sp = sp + 1: stack(sp) = "0"
            Case "ldc.i4.1": sp = sp + 1: stack(sp) = "1"
            Case "ldc.i4.2": sp = sp + 1: stack(sp) = "2"
            Case "ldc.i4.3": sp = sp + 1: stack(sp) = "3"
            Case "ldc.i4.4": sp = sp + 1: stack(sp) = "4"
            Case "ldc.i4.5": sp = sp + 1: stack(sp) = "5"
            Case "ldc.i4.6": sp = sp + 1: stack(sp) = "6"
            Case "ldc.i4.7": sp = sp + 1: stack(sp) = "7"
            Case "ldc.i4.8": sp = sp + 1: stack(sp) = "8"
            Case "ldc.i4.s", "ldc.i4", "ldc.i8": sp = sp + 1: stack(sp) = CStr(gInsVal(j))
            Case "ldc.r4", "ldc.r8": sp = sp + 1: stack(sp) = gInsText(j)
            Case "ldarg.0": sp = sp + 1: stack(sp) = ArgName(0, hasThis, pc, pName, gThisKw)
            Case "ldarg.1": sp = sp + 1: stack(sp) = ArgName(1, hasThis, pc, pName, gThisKw)
            Case "ldarg.2": sp = sp + 1: stack(sp) = ArgName(2, hasThis, pc, pName, gThisKw)
            Case "ldarg.3": sp = sp + 1: stack(sp) = ArgName(3, hasThis, pc, pName, gThisKw)
            Case "ldarg.s", "ldarg", "ldarga.s", "ldarga"
                sp = sp + 1: stack(sp) = ArgName(gInsVal(j), hasThis, pc, pName, gThisKw)
            Case "ldloc.0": sp = sp + 1: stack(sp) = "V_0"
            Case "ldloc.1": sp = sp + 1: stack(sp) = "V_1"
            Case "ldloc.2": sp = sp + 1: stack(sp) = "V_2"
            Case "ldloc.3": sp = sp + 1: stack(sp) = "V_3"
            Case "ldloc.s", "ldloc", "ldloca.s", "ldloca"
                sp = sp + 1: stack(sp) = "V_" & gInsVal(j)
            Case "stloc.0": Call AppendStmt("V_0 = " & PopX(stack, sp))
            Case "stloc.1": Call AppendStmt("V_1 = " & PopX(stack, sp))
            Case "stloc.2": Call AppendStmt("V_2 = " & PopX(stack, sp))
            Case "stloc.3": Call AppendStmt("V_3 = " & PopX(stack, sp))
            Case "stloc.s", "stloc": Call AppendStmt("V_" & gInsVal(j) & " = " & PopX(stack, sp))
            Case "starg.s", "starg"
                Call AppendStmt(ArgName(gInsVal(j), hasThis, pc, pName, gThisKw) & " = " & PopX(stack, sp))
            Case "dup": If sp > 0 Then sp = sp + 1: stack(sp) = stack(sp - 1)
            Case "pop": Call AppendStmt(PopX(stack, sp))
            Case "ldfld", "ldflda"
                a1 = PopX(stack, sp)
                sp = sp + 1: stack(sp) = a1 & "." & FieldNameOnly(gInsVal(j))
            Case "ldsfld", "ldsflda": sp = sp + 1: stack(sp) = FieldFullName(gInsVal(j), lang)
            Case "stfld"
                a2 = PopX(stack, sp): a1 = PopX(stack, sp)
                Call AppendStmt(a1 & "." & FieldNameOnly(gInsVal(j)) & " = " & a2)
            Case "stsfld"
                a2 = PopX(stack, sp)
                Call AppendStmt(FieldFullName(gInsVal(j), lang) & " = " & a2)
            Case "call", "callvirt": Call EmitCall(gInsVal(j), stack, sp, lang)
            Case "newobj": Call EmitNewObj(gInsVal(j), stack, sp, gNewKw, lang)
            Case "castclass"
                a1 = PopX(stack, sp)
                If lang = LANG_CS Then
                    sp = sp + 1: stack(sp) = "((" & gInsText(j) & ")" & a1 & ")"
                Else
                    sp = sp + 1: stack(sp) = "CType(" & a1 & ", " & gInsText(j) & ")"
                End If
            Case "isinst"
                a1 = PopX(stack, sp)
                If lang = LANG_CS Then
                    sp = sp + 1: stack(sp) = "(" & a1 & " as " & gInsText(j) & ")"
                Else
                    sp = sp + 1: stack(sp) = "TryCast(" & a1 & ", " & gInsText(j) & ")"
                End If
            Case "box", "unbox.any", "unbox", "conv.i1", "conv.i2", "conv.i4", "conv.i8", "conv.r4", "conv.r8", "conv.u1", "conv.u2", "conv.u4", "conv.u8", "conv.i", "conv.u", "conv.r.un", "ckfinite"
                'value left on the stack unchanged
            Case "add", "add.ovf", "add.ovf.un": Call BinOp(stack, sp, " + ")
            Case "sub", "sub.ovf", "sub.ovf.un": Call BinOp(stack, sp, " - ")
            Case "mul", "mul.ovf", "mul.ovf.un": Call BinOp(stack, sp, " * ")
            Case "div", "div.un": Call BinOp(stack, sp, " / ")
            Case "rem", "rem.un": Call BinOp(stack, sp, IIf(lang = LANG_VB, " Mod ", " % "))
            Case "and": Call BinOp(stack, sp, IIf(lang = LANG_VB, " And ", " & "))
            Case "or": Call BinOp(stack, sp, IIf(lang = LANG_VB, " Or ", " | "))
            Case "xor": Call BinOp(stack, sp, IIf(lang = LANG_VB, " Xor ", " ^ "))
            Case "shl": Call BinOp(stack, sp, " << ")
            Case "shr", "shr.un": Call BinOp(stack, sp, " >> ")
            Case "ceq": Call BinOp(stack, sp, gEqOp)
            Case "cgt", "cgt.un": Call BinOp(stack, sp, " > ")
            Case "clt", "clt.un": Call BinOp(stack, sp, " < ")
            Case "neg": a1 = PopX(stack, sp): sp = sp + 1: stack(sp) = "(-" & a1 & ")"
            Case "not": a1 = PopX(stack, sp): sp = sp + 1: stack(sp) = IIf(lang = LANG_VB, "(Not " & a1 & ")", "(~" & a1 & ")")
            Case "ldlen": a1 = PopX(stack, sp): sp = sp + 1: stack(sp) = a1 & IIf(lang = LANG_VB, ".Length", ".Length")
            Case "newarr"
                a1 = PopX(stack, sp)
                If lang = LANG_CS Then
                    sp = sp + 1: stack(sp) = "new " & gInsText(j) & "[" & a1 & "]"
                Else
                    sp = sp + 1: stack(sp) = "New " & gInsText(j) & "(" & a1 & " - 1) {}"
                End If
            Case "ldelem.i1", "ldelem.u1", "ldelem.i2", "ldelem.u2", "ldelem.i4", "ldelem.u4", "ldelem.i8", "ldelem.i", "ldelem.r4", "ldelem.r8", "ldelem.ref", "ldelem", "ldelema"
                a2 = PopX(stack, sp): a1 = PopX(stack, sp)
                sp = sp + 1: stack(sp) = a1 & IIf(lang = LANG_VB, "(" & a2 & ")", "[" & a2 & "]")
            Case "stelem.i", "stelem.i1", "stelem.i2", "stelem.i4", "stelem.i8", "stelem.r4", "stelem.r8", "stelem.ref", "stelem"
                Dim ev As String, ei As String, ea As String
                ev = PopX(stack, sp): ei = PopX(stack, sp): ea = PopX(stack, sp)
                Call AppendStmt(ea & IIf(lang = LANG_VB, "(" & ei & ")", "[" & ei & "]") & " = " & ev)
            Case "ldind.i1", "ldind.u1", "ldind.i2", "ldind.u2", "ldind.i4", "ldind.u4", "ldind.i8", "ldind.i", "ldind.r4", "ldind.r8", "ldind.ref"
                'dereference - approximate by leaving the address expression
            Case "stind.i1", "stind.i2", "stind.i4", "stind.i8", "stind.r4", "stind.r8", "stind.i", "stind.ref"
                a2 = PopX(stack, sp): a1 = PopX(stack, sp): Call AppendStmt(a1 & " = " & a2)
            Case "ldtoken": sp = sp + 1: stack(sp) = IIf(lang = LANG_VB, "GetType(" & gInsText(j) & ")", "typeof(" & gInsText(j) & ")")
            Case "ldftn", "ldvirtftn": sp = sp + 1: stack(sp) = gInsText(j)
            Case "br", "br.s", "leave", "leave.s"
                Call LAppend(LT_BR, "", "", gInsVal(j), "", gCurOff): sp = 0
            Case "brtrue", "brtrue.s"
                a1 = PopX(stack, sp)
                Call LAppend(LT_CBR, a1, NegateCond(a1), gInsVal(j), "", gCurOff): sp = 0
            Case "brfalse", "brfalse.s"
                a1 = PopX(stack, sp)
                Call LAppend(LT_CBR, NegateCond(a1), a1, gInsVal(j), "", gCurOff): sp = 0
            Case "beq", "beq.s": Call EmitCBR(stack, sp, "==", "!=", "=", "<>", gInsVal(j))
            Case "bne.un", "bne.un.s": Call EmitCBR(stack, sp, "!=", "==", "<>", "=", gInsVal(j))
            Case "bge", "bge.s", "bge.un", "bge.un.s": Call EmitCBR(stack, sp, ">=", "<", ">=", "<", gInsVal(j))
            Case "bgt", "bgt.s", "bgt.un", "bgt.un.s": Call EmitCBR(stack, sp, ">", "<=", ">", "<=", gInsVal(j))
            Case "ble", "ble.s", "ble.un", "ble.un.s": Call EmitCBR(stack, sp, "<=", ">", "<=", ">", gInsVal(j))
            Case "blt", "blt.s", "blt.un", "blt.un.s": Call EmitCBR(stack, sp, "<", ">=", "<", ">=", gInsVal(j))
            Case "switch"
                a1 = PopX(stack, sp)
                Call LAppend(LT_SWITCH, a1, "", 0, gSwitchTargets(j), gCurOff): sp = 0
            Case "ret"
                If retIL <> "void" And sp > 0 Then
                    Call LAppend(LT_RET, PopX(stack, sp), "", 0, "", gCurOff)
                Else
                    Call LAppend(LT_RET, "", "", 0, "", gCurOff)
                End If
                sp = 0
            Case "throw"
                a1 = PopX(stack, sp): Call LAppend(LT_THROW, a1, "", 0, "", gCurOff): sp = 0
            Case "rethrow"
                Call LAppend(LT_RAW, IIf(lang = LANG_CS, "throw;", "Throw"), "", 0, "", gCurOff)
            Case "endfinally", "endfilter"
                'implicit in the finally / filter block
            Case Else
                Call LAppend(LT_RAW, IIf(lang = LANG_CS, "// il: ", "' il: ") & op, "", 0, "", gCurOff)
        End Select
    Next

    'Wrap try/catch/finally regions, then fold structured control flow.
    Call WrapExceptionHandlers
    Call StructureFlow

    ReconstructBody = GetLocalDecls(localTok, lang) & RenderList()
    Exit Function
bad:
    ReconstructBody = "        " & IIf(lang = LANG_CS, "// ", "' ") & "reconstruction failed; see decompiled.il" & vbCrLf
End Function

Private Function PopX(ByRef stack() As String, ByRef sp As Long) As String
    If sp > 0 Then
        PopX = stack(sp)
        sp = sp - 1
    Else
        PopX = "?"
    End If
End Function

Private Sub BinOp(ByRef stack() As String, ByRef sp As Long, oper As String)
    Dim b As String, a As String
    b = PopX(stack, sp)
    a = PopX(stack, sp)
    sp = sp + 1
    stack(sp) = "(" & a & oper & b & ")"
End Sub

Private Function ArgName(index As Long, hasThis As Boolean, pc As Long, pName() As String, thisKw As String) As String
    If hasThis And index = 0 Then
        ArgName = thisKw
        Exit Function
    End If
    Dim ord As Long
    If hasThis Then ord = index Else ord = index + 1
    If ord >= 1 And ord <= pc Then
        ArgName = pName(ord)
    Else
        ArgName = "arg" & index
    End If
End Function

Private Function FieldNameOnly(token As Long) As String
    Dim owner As String, name As String
    If FieldRefFromToken(token, owner, name, LANG_CS) Then
        FieldNameOnly = name
    Else
        FieldNameOnly = "field"
    End If
End Function

Private Function FieldFullName(token As Long, lang As Integer) As String
    Dim owner As String, name As String
    If FieldRefFromToken(token, owner, name, lang) Then
        FieldFullName = owner & "." & name
    Else
        FieldFullName = "field"
    End If
End Function

Private Sub EmitCall(token As Long, ByRef stack() As String, ByRef sp As Long, lang As Integer)
    Dim name As String, owner As String
    Dim argc As Long, hasThisC As Boolean, returnsVal As Boolean
    If Not CallInfoFromToken(token, name, owner, argc, hasThisC, returnsVal, lang) Then
        Call AppendStmt(IIf(lang = LANG_CS, "/* call */", "' call"))
        Exit Sub
    End If
    Dim a() As String
    ReDim a(argc + 1)
    Dim z As Long
    For z = argc To 1 Step -1
        a(z) = PopX(stack, sp)
    Next
    Dim argstr As String
    For z = 1 To argc
        If z > 1 Then argstr = argstr & ", "
        argstr = argstr & a(z)
    Next
    Dim target As String
    If hasThisC Then
        target = PopX(stack, sp)
    Else
        target = owner
    End If

    'Property accessors read nicer as member access.
    If hasThisC And Left$(name, 4) = "get_" And argc = 0 Then
        sp = sp + 1: stack(sp) = target & "." & Mid$(name, 5)
        Exit Sub
    ElseIf hasThisC And Left$(name, 4) = "set_" And argc = 1 Then
        Call AppendStmt(target & "." & Mid$(name, 5) & " = " & a(1))
        Exit Sub
    End If

    If name = ".ctor" Then
        'base/this constructor call - note it but keep going
        Call AppendStmt(IIf(lang = LANG_CS, "// base ctor ", "' base ctor ") & owner & "(" & argstr & ")")
        Exit Sub
    End If
    Dim callee As String
    If hasThisC Then
        callee = target & "." & name
    Else
        callee = owner & "." & name
    End If
    Dim expr As String
    expr = callee & "(" & argstr & ")"
    If returnsVal Then
        sp = sp + 1: stack(sp) = expr
    Else
        Call AppendStmt(expr)
    End If
End Sub

Private Sub EmitNewObj(token As Long, ByRef stack() As String, ByRef sp As Long, newKw As String, lang As Integer)
    Dim name As String, owner As String
    Dim argc As Long, hasThisC As Boolean, returnsVal As Boolean
    If Not CallInfoFromToken(token, name, owner, argc, hasThisC, returnsVal, lang) Then
        sp = sp + 1: stack(sp) = newKw & "object()"
        Exit Sub
    End If
    Dim a() As String
    ReDim a(argc + 1)
    Dim z As Long
    For z = argc To 1 Step -1
        a(z) = PopX(stack, sp)
    Next
    Dim argstr As String
    For z = 1 To argc
        If z > 1 Then argstr = argstr & ", "
        argstr = argstr & a(z)
    Next
    sp = sp + 1
    stack(sp) = newKw & owner & "(" & argstr & ")"
End Sub

'==========================================================================
'  Per-type accessors (consumed by frmMain's project tree / code viewer)
'==========================================================================
Public Function GetDotNetTypeCount() As Long
    GetDotNetTypeCount = gNetTypeCount
End Function

Public Function GetDotNetTypeName(ByVal i As Long) As String
    If i >= 0 And i < gNetTypeCount Then GetDotNetTypeName = gNetTypeName(i)
End Function

Public Function GetDotNetTypeMethods(ByVal i As Long) As String
    If i >= 0 And i < gNetTypeCount Then GetDotNetTypeMethods = gNetTypeMethods(i)
End Function

'lang: 0 = IL, 1 = C#, 2 = VB.NET
Public Function GetDotNetTypeCode(ByVal i As Long, ByVal lang As Integer) As String
    If i < 0 Or i >= gNetTypeCount Then Exit Function
    Select Case lang
        Case LANG_IL: GetDotNetTypeCode = gNetTypeIL(i)
        Case LANG_VB: GetDotNetTypeCode = gNetTypeVB(i)
        Case Else: GetDotNetTypeCode = gNetTypeCS(i)
    End Select
End Function

'One-line, C#-style method signature list for a type (vbLf separated).
Private Function BuildMethodSigList(t As Long) As String
    Dim mFirst As Long, mLast As Long
    Call GetMethodRange(t, mFirst, mLast)
    Dim s As String, mi As Long
    For mi = mFirst To mLast
        If mi >= 1 And mi <= iRows(6) Then
            s = s & GetMethodDisplaySig(mi, t, LANG_CS) & vbLf
        End If
    Next
    BuildMethodSigList = s
End Function

Private Function GetMethodDisplaySig(mi As Long, t As Long, lang As Integer) As String
On Error GoTo bad
    Dim nm As String
    nm = GetString(MethodStruct(mi).name)
    Dim isCtor As Boolean
    isCtor = (nm = ".ctor" Or nm = ".cctor")
    Dim ret As String
    Dim p() As String
    Dim pc As Long, hasThis As Boolean
    Call ParseMethodSigFull(MethodStruct(mi).signature, ret, p, pc, hasThis, lang)
    Dim pName() As String
    ReDim pName(pc + 1)
    Call GetParamNames(mi, pc, pName)
    Dim head As String
    If isCtor Then
        head = GetString(TypeDefStruct(t).name)
    Else
        head = ret & " " & nm
    End If
    GetMethodDisplaySig = head & "(" & JoinParamsNamed(p, pName, pc, lang) & ")"
    Exit Function
bad:
    GetMethodDisplaySig = GetString(MethodStruct(mi).name) & "()"
End Function

'==========================================================================
'  Build a navigable solution scaffold on disk from the reconstructed code.
'  Writes one .cs and one .vb per class, plus SDK-style project files and a
'  .sln.  The reconstruction is best-effort and may not compile as-is.
'==========================================================================
Public Sub BuildDotNetSolution(ByVal sPath As String)
On Error GoTo bad
    If gNetTypeCount = 0 Then
        MsgBox "No .NET classes are loaded. Open a .NET assembly first.", vbExclamation
        Exit Sub
    End If
    Dim baseName As String
    baseName = SanitizeName(SFile)
    If Len(baseName) = 0 Then baseName = "Decompiled"

    Dim root As String
    root = sPath & "\" & baseName
    Call EnsureDir(sPath)
    Call EnsureDir(root)
    Call EnsureDir(root & "\CSharp")
    Call EnsureDir(root & "\VBNet")

    Dim i As Long
    For i = 0 To gNetTypeCount - 1
        Dim fn As String
        fn = SanitizeName(GetDotNetTypeName(i))
        If Len(fn) = 0 Then fn = "Type" & i
        Call WriteTextFile(root & "\CSharp\" & fn & ".cs", "using System;" & vbCrLf & vbCrLf & GetDotNetTypeCode(i, LANG_CS))
        Call WriteTextFile(root & "\VBNet\" & fn & ".vb", "Imports System" & vbCrLf & vbCrLf & GetDotNetTypeCode(i, LANG_VB))
    Next

    Call WriteTextFile(root & "\CSharp\" & baseName & ".CS.csproj", CsProjText())
    Call WriteTextFile(root & "\VBNet\" & baseName & ".VB.vbproj", VbProjText())
    Call WriteTextFile(root & "\" & baseName & ".sln", SlnText(baseName))
    Call WriteTextFile(root & "\README.txt", ReadmeText())

    If gQuietMode = False Then
        Dim r As VbMsgBoxResult
        r = MsgBox("Solution scaffold written to:" & vbCrLf & root & vbCrLf & vbCrLf & _
                   "Note: the reconstruction is best-effort and may not compile as-is." & vbCrLf & _
                   "Open the folder now?", vbYesNo + vbInformation)
        If r = vbYes Then Shell "explorer.exe " & Chr$(34) & root & Chr$(34), vbNormalFocus
    End If
    Exit Sub
bad:
    If gQuietMode = False Then MsgBox "Error_BuildDotNetSolution: " & err.Description, vbCritical
End Sub

Private Sub EnsureDir(ByVal path As String)
    On Error Resume Next
    MkDir path
End Sub

Private Sub WriteTextFile(ByVal path As String, ByVal content As String)
    Dim ff As Integer
    ff = FreeFile
    Open path For Output As #ff
    Print #ff, content
    Close #ff
End Sub

Private Function SanitizeName(ByVal s As String) As String
    Dim i As Long, c As String, r As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        Select Case c
            Case "\", "/", ":", "*", "?", """", "<", ">", "|"
                r = r & "_"
            Case Else
                r = r & c
        End Select
    Next
    SanitizeName = r
End Function

Private Function CsProjText() As String
    CsProjText = "<Project Sdk=""Microsoft.NET.Sdk"">" & vbCrLf & _
        "  <PropertyGroup>" & vbCrLf & _
        "    <OutputType>Library</OutputType>" & vbCrLf & _
        "    <TargetFramework>net48</TargetFramework>" & vbCrLf & _
        "    <Nullable>disable</Nullable>" & vbCrLf & _
        "    <GenerateAssemblyInfo>false</GenerateAssemblyInfo>" & vbCrLf & _
        "  </PropertyGroup>" & vbCrLf & _
        "</Project>" & vbCrLf
End Function

Private Function VbProjText() As String
    VbProjText = "<Project Sdk=""Microsoft.NET.Sdk"">" & vbCrLf & _
        "  <PropertyGroup>" & vbCrLf & _
        "    <OutputType>Library</OutputType>" & vbCrLf & _
        "    <TargetFramework>net48</TargetFramework>" & vbCrLf & _
        "    <GenerateAssemblyInfo>false</GenerateAssemblyInfo>" & vbCrLf & _
        "  </PropertyGroup>" & vbCrLf & _
        "</Project>" & vbCrLf
End Function

Private Function SlnText(ByVal baseName As String) As String
    Dim csG As String, vbG As String, csP As String, vbP As String
    csG = "{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}"
    vbG = "{F184B08F-C81C-45F6-A57F-5ABD9991F28F}"
    csP = "{11111111-1111-1111-1111-111111111111}"
    vbP = "{22222222-2222-2222-2222-222222222222}"
    Dim s As String
    s = "Microsoft Visual Studio Solution File, Format Version 12.00" & vbCrLf
    s = s & "# Visual Studio Version 17" & vbCrLf
    s = s & "Project(" & Chr$(34) & csG & Chr$(34) & ") = " & Chr$(34) & baseName & ".CS" & Chr$(34) & ", " & Chr$(34) & "CSharp\" & baseName & ".CS.csproj" & Chr$(34) & ", " & Chr$(34) & csP & Chr$(34) & vbCrLf & "EndProject" & vbCrLf
    s = s & "Project(" & Chr$(34) & vbG & Chr$(34) & ") = " & Chr$(34) & baseName & ".VB" & Chr$(34) & ", " & Chr$(34) & "VBNet\" & baseName & ".VB.vbproj" & Chr$(34) & ", " & Chr$(34) & vbP & Chr$(34) & vbCrLf & "EndProject" & vbCrLf
    s = s & "Global" & vbCrLf
    s = s & "  GlobalSection(SolutionConfigurationPlatforms) = preSolution" & vbCrLf
    s = s & "    Debug|Any CPU = Debug|Any CPU" & vbCrLf
    s = s & "    Release|Any CPU = Release|Any CPU" & vbCrLf
    s = s & "  EndGlobalSection" & vbCrLf
    s = s & "  GlobalSection(ProjectConfigurationPlatforms) = postSolution" & vbCrLf
    s = s & ConfigLines(csP)
    s = s & ConfigLines(vbP)
    s = s & "  EndGlobalSection" & vbCrLf
    s = s & "EndGlobal" & vbCrLf
    SlnText = s
End Function

Private Function ConfigLines(ByVal g As String) As String
    Dim s As String
    s = "    " & g & ".Debug|Any CPU.ActiveCfg = Debug|Any CPU" & vbCrLf
    s = s & "    " & g & ".Debug|Any CPU.Build.0 = Debug|Any CPU" & vbCrLf
    s = s & "    " & g & ".Release|Any CPU.ActiveCfg = Release|Any CPU" & vbCrLf
    s = s & "    " & g & ".Release|Any CPU.Build.0 = Release|Any CPU" & vbCrLf
    ConfigLines = s
End Function

Private Function ReadmeText() As String
    ReadmeText = "Generated by Semi VB Decompiler - VisualBasicZone.com" & vbCrLf & _
        "This is a best-effort reconstruction of a .NET assembly." & vbCrLf & _
        "Class structure, fields and method signatures are accurate; method" & vbCrLf & _
        "bodies are simple stack-machine reconstructions and may need edits to" & vbCrLf & _
        "compile.  See the per-class IL listing for the authoritative output." & vbCrLf
End Function

'==========================================================================
'  Control-flow reconstruction support
'==========================================================================
Private Sub LAppend(ty As Integer, txt As String, alt As String, tgt As Long, sw As String, off As Long)
    If gLCount > UBound(gLType) Then
        ReDim Preserve gLType(gLCount + 64)
        ReDim Preserve gLText(gLCount + 64)
        ReDim Preserve gLAlt(gLCount + 64)
        ReDim Preserve gLTarget(gLCount + 64)
        ReDim Preserve gLSwitch(gLCount + 64)
        ReDim Preserve gLOffset(gLCount + 64)
        ReDim Preserve gLDead(gLCount + 64)
    End If
    gLType(gLCount) = ty
    gLText(gLCount) = txt
    gLAlt(gLCount) = alt
    gLTarget(gLCount) = tgt
    gLSwitch(gLCount) = sw
    gLOffset(gLCount) = off
    gLDead(gLCount) = False
    gLCount = gLCount + 1
End Sub

Private Sub AppendStmt(txt As String)
    Call LAppend(LT_STMT, txt & gTerm, "", 0, "", gCurOff)
End Sub

Private Function OffsetInTargets(targets() As Long, count As Long, off As Long) As Boolean
    Dim i As Long
    For i = 0 To count - 1
        If targets(i) = off Then OffsetInTargets = True: Exit Function
    Next
End Function

Private Function NegateCond(c As String) As String
    If gLang = LANG_VB Then
        NegateCond = "Not (" & c & ")"
    Else
        NegateCond = "!(" & c & ")"
    End If
End Function

Private Sub EmitCBR(ByRef stack() As String, ByRef sp As Long, csT As String, csN As String, vbT As String, vbN As String, tgt As Long)
    Dim b As String, a As String
    b = PopX(stack, sp)
    a = PopX(stack, sp)
    Dim taken As String, nottaken As String
    If gLang = LANG_CS Then
        taken = "(" & a & " " & csT & " " & b & ")"
        nottaken = "(" & a & " " & csN & " " & b & ")"
    Else
        taken = "(" & a & " " & vbT & " " & b & ")"
        nottaken = "(" & a & " " & vbN & " " & b & ")"
    End If
    Call LAppend(LT_CBR, taken, nottaken, tgt, "", gCurOff)
    sp = 0
End Sub

'--- Exception-handling clauses (ECMA-335 II.25.4.6) ---------------------
Private Sub ReadEHClauses(f As Integer, foff As Long, codeSize As Long)
    On Error GoTo done
    Dim secStart As Long
    secStart = ((foff + 12 + codeSize + 3) \ 4) * 4
    Seek #f, secStart + 1
    Dim more As Boolean
    more = True
    Do While more
        Dim kind As Long
        kind = GetByteByFile(f)
        more = (kind And &H80) <> 0
        If (kind And &H1) = 0 Then Exit Sub      'not an EH table
        Dim clauses As Long, c As Long
        If (kind And &H40) <> 0 Then
            'Fat: 3-byte data size, 24-byte clauses.
            Dim s1 As Long, s2 As Long, s3 As Long
            s1 = GetByteByFile(f): s2 = GetByteByFile(f): s3 = GetByteByFile(f)
            clauses = ((s1 + s2 * 256 + s3 * 65536) - 4) \ 24
            For c = 0 To clauses - 1
                Call StoreEH(GetDWordByFile(f), GetDWordByFile(f), GetDWordByFile(f), GetDWordByFile(f), GetDWordByFile(f), GetDWordByFile(f))
            Next
        Else
            'Small: 1-byte data size + 2 padding, 12-byte clauses.
            Dim ds As Long
            ds = GetByteByFile(f)
            Call GetWordByFile(f)
            clauses = (ds - 4) \ 12
            For c = 0 To clauses - 1
                Dim fl As Long, t0 As Long, tl As Long, h0 As Long, hl As Long, tk As Long
                fl = GetWordByFile(f)
                t0 = GetWordByFile(f)
                tl = GetByteByFile(f)
                h0 = GetWordByFile(f)
                hl = GetByteByFile(f)
                tk = GetDWordByFile(f)
                Call StoreEH(fl, t0, tl, h0, hl, tk)
            Next
        End If
    Loop
    Exit Sub
done:
End Sub

Private Sub StoreEH(ByVal fl As Long, ByVal t0 As Long, ByVal tl As Long, ByVal h0 As Long, ByVal hl As Long, ByVal tk As Long)
    If gEHCount = 0 Then
        ReDim gEHFlags(16): ReDim gEHTryOff(16): ReDim gEHTryLen(16)
        ReDim gEHHandOff(16): ReDim gEHHandLen(16): ReDim gEHToken(16)
    ElseIf gEHCount > UBound(gEHFlags) Then
        ReDim Preserve gEHFlags(gEHCount + 16)
        ReDim Preserve gEHTryOff(gEHCount + 16)
        ReDim Preserve gEHTryLen(gEHCount + 16)
        ReDim Preserve gEHHandOff(gEHCount + 16)
        ReDim Preserve gEHHandLen(gEHCount + 16)
        ReDim Preserve gEHToken(gEHCount + 16)
    End If
    gEHFlags(gEHCount) = fl
    gEHTryOff(gEHCount) = t0
    gEHTryLen(gEHCount) = tl
    gEHHandOff(gEHCount) = h0
    gEHHandLen(gEHCount) = hl
    gEHToken(gEHCount) = tk
    gEHCount = gEHCount + 1
End Sub

Private Sub WrapExceptionHandlers()
    On Error Resume Next
    If gEHCount = 0 Then Exit Sub
    Dim handled() As Boolean
    ReDim handled(gEHCount)
    Dim c As Long
    For c = 0 To gEHCount - 1
        If Not handled(c) Then
            Dim tOff As Long, tLen As Long
            tOff = gEHTryOff(c): tLen = gEHTryLen(c)
            Dim mem() As Long, mc As Long
            ReDim mem(gEHCount)
            mc = 0
            Dim d As Long
            For d = 0 To gEHCount - 1
                If gEHTryOff(d) = tOff And gEHTryLen(d) = tLen Then
                    mem(mc) = d: mc = mc + 1: handled(d) = True
                End If
            Next
            Call InsertStructBefore(FirstIdxAtOffset(tOff), LT_OPEN, IIf(gLang = LANG_CS, "try {", "Try"), tOff)
            Dim hm As Long
            For hm = 0 To mc - 1
                Call InsertStructBefore(FirstIdxAtOffset(gEHHandOff(mem(hm))), LT_MID, HandlerHeader(mem(hm)), gEHHandOff(mem(hm)))
            Next
            Dim lastEnd As Long
            lastEnd = gEHHandOff(mem(mc - 1)) + gEHHandLen(mem(mc - 1))
            Call InsertStructBefore(FirstIdxAtOffset(lastEnd), LT_CLOSE, IIf(gLang = LANG_CS, "}", "End Try"), lastEnd)
        End If
    Next
End Sub

Private Function HandlerHeader(c As Long) As String
    Dim fl As Long
    fl = gEHFlags(c)
    If (fl And 2) <> 0 Then
        HandlerHeader = IIf(gLang = LANG_CS, "} finally {", "Finally")
    ElseIf (fl And 1) <> 0 Then
        HandlerHeader = IIf(gLang = LANG_CS, "} catch /* filter */ {", "Catch ' filter")
    ElseIf (fl And 4) <> 0 Then
        HandlerHeader = IIf(gLang = LANG_CS, "} catch /* fault */ {", "Catch ' fault")
    Else
        Dim tn As String
        tn = GetTokenDisplay(gEHToken(c), gLang)
        If gLang = LANG_CS Then
            HandlerHeader = "} catch (" & tn & " ex) {"
        Else
            HandlerHeader = "Catch ex As " & tn
        End If
    End If
End Function

Private Function FirstIdxAtOffset(off As Long) As Long
    Dim i As Long
    For i = 0 To gLCount - 1
        If Not gLDead(i) Then
            If gLOffset(i) >= off Then FirstIdxAtOffset = i: Exit Function
        End If
    Next
    FirstIdxAtOffset = gLCount
End Function

Private Sub InsertStructBefore(atIdx As Long, ty As Integer, txt As String, off As Long)
    If gLCount > UBound(gLType) Then
        ReDim Preserve gLType(gLCount + 64)
        ReDim Preserve gLText(gLCount + 64)
        ReDim Preserve gLAlt(gLCount + 64)
        ReDim Preserve gLTarget(gLCount + 64)
        ReDim Preserve gLSwitch(gLCount + 64)
        ReDim Preserve gLOffset(gLCount + 64)
        ReDim Preserve gLDead(gLCount + 64)
    End If
    Dim k As Long
    For k = gLCount To atIdx + 1 Step -1
        gLType(k) = gLType(k - 1)
        gLText(k) = gLText(k - 1)
        gLAlt(k) = gLAlt(k - 1)
        gLTarget(k) = gLTarget(k - 1)
        gLSwitch(k) = gLSwitch(k - 1)
        gLOffset(k) = gLOffset(k - 1)
        gLDead(k) = gLDead(k - 1)
    Next
    gLType(atIdx) = ty
    gLText(atIdx) = txt
    gLAlt(atIdx) = ""
    gLTarget(atIdx) = 0
    gLSwitch(atIdx) = ""
    gLOffset(atIdx) = off
    gLDead(atIdx) = False
    gLCount = gLCount + 1
End Sub

'--- Structuring: fold goto/label spaghetti into if/else/loops ------------
Private Sub StructureFlow()
    On Error Resume Next
    Dim changed As Boolean, iter As Long
    Do
        changed = False
        iter = iter + 1
        If FoldIfElse() Then
            changed = True
        ElseIf FoldIf() Then
            changed = True
        ElseIf FoldDoWhile() Then
            changed = True
        ElseIf FoldWhile() Then
            changed = True
        ElseIf FoldGotoNext() Then
            changed = True
        End If
    Loop While changed And iter < 2000
    Call RemoveUnusedLabels
End Sub

Private Function NextNonDead(i As Long) As Long
    Dim k As Long
    For k = i + 1 To gLCount - 1
        If Not gLDead(k) Then NextNonDead = k: Exit Function
    Next
    NextNonDead = gLCount
End Function

Private Function PrevNonDead(i As Long) As Long
    Dim k As Long
    For k = i - 1 To 0 Step -1
        If Not gLDead(k) Then PrevNonDead = k: Exit Function
    Next
    PrevNonDead = -1
End Function

Private Function FindLabelIdx(off As Long) As Long
    Dim k As Long
    For k = 0 To gLCount - 1
        If Not gLDead(k) Then
            If gLType(k) = LT_LABEL And gLOffset(k) = off Then FindLabelIdx = k: Exit Function
        End If
    Next
    FindLabelIdx = -1
End Function

Private Function RefCount(off As Long) As Long
    Dim k As Long, n As Long
    For k = 0 To gLCount - 1
        If Not gLDead(k) Then
            If (gLType(k) = LT_CBR Or gLType(k) = LT_BR) And gLTarget(k) = off Then n = n + 1
            If gLType(k) = LT_SWITCH And Len(gLSwitch(k)) > 0 Then
                Dim sw() As String, q As Long
                sw = Split(gLSwitch(k), ",")
                For q = 0 To UBound(sw)
                    If CLng(sw(q)) = off Then n = n + 1
                Next
            End If
        End If
    Next
    RefCount = n
End Function

'A region is "straight-line" if it has no labels, gotos or switches - only
'plain statements plus optional early return/throw.  Such a region is a
'single-entry / single-exit block and is safe to wrap in if/while.
Private Function StraightLine(a As Long, b As Long) As Boolean
    Dim k As Long
    StraightLine = True
    For k = a To b
        If k >= 0 And k < gLCount Then
            If Not gLDead(k) Then
                Select Case gLType(k)
                    Case LT_STMT, LT_RAW, LT_RET, LT_THROW
                        'allowed
                    Case Else
                        StraightLine = False: Exit Function
                End Select
            End If
        End If
    Next
End Function

Private Function FoldIf() As Boolean
    Dim i As Long
    For i = 0 To gLCount - 1
        If Not gLDead(i) And gLType(i) = LT_CBR Then
            Dim L As Long
            L = FindLabelIdx(gLTarget(i))
            If L > i Then
                If StraightLine(i + 1, L - 1) Then
                    If NextNonDead(i) < L Then
                        Dim altc As String
                        altc = gLAlt(i)
                        gLType(i) = LT_OPEN
                        gLText(i) = IIf(gLang = LANG_CS, "if (" & altc & ") {", "If " & altc & " Then")
                        Call InsertStructBefore(L, LT_CLOSE, IIf(gLang = LANG_CS, "}", "End If"), gLOffset(L))
                        FoldIf = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

Private Function FoldIfElse() As Boolean
    Dim i As Long
    For i = 0 To gLCount - 1
        If Not gLDead(i) And gLType(i) = LT_CBR Then
            Dim L As Long
            L = FindLabelIdx(gLTarget(i))
            If L > i Then
                Dim k As Long
                k = PrevNonDead(L)
                If k > i Then
                    If gLType(k) = LT_BR Then
                        Dim Eoff As Long, m As Long
                        Eoff = gLTarget(k)
                        m = FindLabelIdx(Eoff)
                        If m > L Then
                            If StraightLine(i + 1, k - 1) And StraightLine(L + 1, m - 1) Then
                                If RefCount(gLTarget(i)) = 1 Then
                                    Dim altc As String
                                    altc = gLAlt(i)
                                    gLType(i) = LT_OPEN
                                    gLText(i) = IIf(gLang = LANG_CS, "if (" & altc & ") {", "If " & altc & " Then")
                                    gLType(k) = LT_MID
                                    gLText(k) = IIf(gLang = LANG_CS, "} else {", "Else")
                                    gLDead(L) = True
                                    Call InsertStructBefore(m, LT_CLOSE, IIf(gLang = LANG_CS, "}", "End If"), gLOffset(m))
                                    FoldIfElse = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
End Function

Private Function FoldDoWhile() As Boolean
    Dim i As Long
    For i = 0 To gLCount - 1
        If Not gLDead(i) And gLType(i) = LT_CBR Then
            Dim H As Long
            H = FindLabelIdx(gLTarget(i))
            If H >= 0 And H < i Then
                If StraightLine(H + 1, i - 1) Then
                    If RefCount(gLTarget(i)) = 1 Then
                        gLType(H) = LT_OPEN
                        gLText(H) = IIf(gLang = LANG_CS, "do {", "Do")
                        gLType(i) = LT_CLOSE
                        If gLang = LANG_CS Then
                            gLText(i) = "} while (" & gLText(i) & ");"
                        Else
                            gLText(i) = "Loop While " & gLText(i)
                        End If
                        FoldDoWhile = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

Private Function FoldWhile() As Boolean
    Dim i As Long
    For i = 0 To gLCount - 1
        If Not gLDead(i) And gLType(i) = LT_BR Then
            Dim Toff As Long, Tidx As Long
            Toff = gLTarget(i)
            Tidx = FindLabelIdx(Toff)
            If Tidx > i Then
                Dim Bidx As Long
                Bidx = NextNonDead(i)
                If Bidx < gLCount Then
                    If gLType(Bidx) = LT_LABEL Then
                        Dim Boff As Long
                        Boff = gLOffset(Bidx)
                        Dim cbrIdx As Long
                        cbrIdx = NextNonDead(Tidx)
                        If cbrIdx < gLCount Then
                            If gLType(cbrIdx) = LT_CBR And gLTarget(cbrIdx) = Boff Then
                                If StraightLine(Bidx + 1, Tidx - 1) Then
                                    If RefCount(Boff) = 1 And RefCount(Toff) = 1 Then
                                        gLDead(i) = True
                                        gLType(Bidx) = LT_OPEN
                                        If gLang = LANG_CS Then
                                            gLText(Bidx) = "while (" & gLText(cbrIdx) & ") {"
                                        Else
                                            gLText(Bidx) = "While " & gLText(cbrIdx)
                                        End If
                                        gLDead(Tidx) = True
                                        gLType(cbrIdx) = LT_CLOSE
                                        gLText(cbrIdx) = IIf(gLang = LANG_CS, "}", "End While")
                                        FoldWhile = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
End Function

Private Function FoldGotoNext() As Boolean
    Dim i As Long
    For i = 0 To gLCount - 1
        If Not gLDead(i) And gLType(i) = LT_BR Then
            Dim nx As Long
            nx = NextNonDead(i)
            If nx < gLCount Then
                If gLType(nx) = LT_LABEL And gLOffset(nx) = gLTarget(i) Then
                    gLDead(i) = True
                    FoldGotoNext = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Sub RemoveUnusedLabels()
    Dim i As Long
    For i = 0 To gLCount - 1
        If Not gLDead(i) And gLType(i) = LT_LABEL Then
            If RefCount(gLOffset(i)) = 0 Then gLDead(i) = True
        End If
    Next
End Sub

'--- Render the structured list to source text ---------------------------
Private Function RenderList() As String
    Dim s As String, i As Long, lvl As Long
    lvl = 2
    For i = 0 To gLCount - 1
        If Not gLDead(i) Then
            Select Case gLType(i)
                Case LT_OPEN
                    s = s & Space$(lvl * 4) & gLText(i) & vbCrLf
                    lvl = lvl + 1
                Case LT_MID
                    s = s & Space$((lvl - 1) * 4) & gLText(i) & vbCrLf
                Case LT_CLOSE
                    lvl = lvl - 1
                    If lvl < 0 Then lvl = 0
                    s = s & Space$(lvl * 4) & gLText(i) & vbCrLf
                Case LT_LABEL
                    If gLang = LANG_CS Then
                        s = s & Space$(lvl * 4) & gLText(i) & ": ;" & vbCrLf
                    Else
                        s = s & Space$(lvl * 4) & gLText(i) & ":" & vbCrLf
                    End If
                Case LT_BR
                    s = s & Space$(lvl * 4) & RenderGoto(gLTarget(i)) & vbCrLf
                Case LT_CBR
                    s = s & Space$(lvl * 4) & RenderIfGoto(gLText(i), gLTarget(i)) & vbCrLf
                Case LT_RET
                    s = s & Space$(lvl * 4) & RenderRet(gLText(i)) & vbCrLf
                Case LT_THROW
                    s = s & Space$(lvl * 4) & gThrowKw & " " & gLText(i) & gTerm & vbCrLf
                Case LT_SWITCH
                    s = s & RenderSwitch(gLText(i), gLSwitch(i), lvl)
                Case Else
                    s = s & Space$(lvl * 4) & gLText(i) & vbCrLf
            End Select
        End If
    Next
    RenderList = s
End Function

Private Function RenderGoto(t As Long) As String
    If gLang = LANG_CS Then
        RenderGoto = "goto IL_" & Hex4(t) & ";"
    Else
        RenderGoto = "GoTo IL_" & Hex4(t)
    End If
End Function

Private Function RenderIfGoto(cond As String, t As Long) As String
    If gLang = LANG_CS Then
        RenderIfGoto = "if (" & cond & ") goto IL_" & Hex4(t) & ";"
    Else
        RenderIfGoto = "If " & cond & " Then GoTo IL_" & Hex4(t)
    End If
End Function

Private Function RenderRet(txt As String) As String
    If Len(txt) = 0 Then
        RenderRet = IIf(gLang = LANG_CS, "return;", "Return")
    Else
        RenderRet = IIf(gLang = LANG_CS, "return " & txt & ";", "Return " & txt)
    End If
End Function

Private Function RenderSwitch(val As String, csv As String, lvl As Long) As String
    If Len(csv) = 0 Then Exit Function
    Dim s As String, sw() As String, k As Long
    sw = Split(csv, ",")
    For k = 0 To UBound(sw)
        Dim lbl As String
        lbl = "IL_" & Hex4(CLng(sw(k)))
        If gLang = LANG_CS Then
            s = s & Space$(lvl * 4) & "if (" & val & " == " & k & ") goto " & lbl & ";" & vbCrLf
        Else
            s = s & Space$(lvl * 4) & "If " & val & " = " & k & " Then GoTo " & lbl & vbCrLf
        End If
    Next
    RenderSwitch = s
End Function

'--- Local variable declarations from the method's LocalVarSig ------------
Private Function GetLocalDecls(localTok As Long, lang As Integer) As String
    On Error GoTo none
    If localTok = 0 Then Exit Function
    Dim table As Long, row As Long
    table = (localTok \ &H1000000) And &HFF
    If table <> &H11 Then Exit Function          'StandAloneSig
    row = localTok And &HFFFFFF
    If row < 1 Or row > iRows(&H11) Then Exit Function
    Dim pos As Long
    If Not BlobDataStart(StandAloneSigStruct(row).index, pos) Then Exit Function
    Dim first As Long
    first = BlobByteArray(pos): pos = pos + 1
    If first <> &H7 Then Exit Function            'LOCAL_SIG
    Dim cnt As Long
    cnt = ReadCompressedUInt(BlobByteArray, pos)
    Dim s As String, k As Long
    For k = 0 To cnt - 1
        Dim ty As String
        ty = ParseTypeSig(BlobByteArray, pos, lang)
        If lang = LANG_CS Then
            s = s & "        " & ty & " V_" & k & ";" & vbCrLf
        Else
            s = s & "        Dim V_" & k & " As " & ty & vbCrLf
        End If
    Next
    GetLocalDecls = s
    Exit Function
none:
End Function
