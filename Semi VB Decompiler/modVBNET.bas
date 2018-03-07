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
    lSignature As Double 'DWORD  “Magic” signature for physical metadata, currently 0x424A5342   - BSJB
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

Private MyDotNet As SemiVBHelper
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
        MsgBox ".Net Runtime is not installed on this machine. Please download and install it.  To Process .Net files", vbCritical
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
    
    Set MyDotNet = New SemiVBHelper
    
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

    Console.SaveConsoleToFile App.Path & "\dump\" & SFile & "\netconsole.txt"
    'Make the console report menu visible
    frmMain.mnuToolsNetConsole.Visible = True
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
    Valid = MyDotNet.BConverterToInt64(MetaDataByteArray, 8)
    'Debug.Print "VAILD: " & Valid
    ReDim iRows(64)
    
    Dim k As Long
    For k = 0 To 63
        Dim lTablepresent As Long
        '#DONE
        'int tablepresent = (int)(valid >> k ) & 1;
        ''lTablepresent = HiLo.DWordShiftR(Valid, k) And 1
        lTablepresent = MyDotNet.isTablePresent(MetaDataByteArray, k)
        'lTablepresent = HiLo.INT64ShiftR(Valid, k) And 1
        If lTablepresent = 1 Then
            '
            '#DONE
            'rows[k] = BitConverter.ToInt32(metadata , tableoffset);
            'iRows(k) = modVBNET.BitConverterToInt32(MetaDataByteArray, TableOffset)
            iRows(k) = MyDotNet.BConverterToInt32(MetaDataByteArray, TableOffset)
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

Public Function BitConverterToInt16(bArray() As Byte, offset As Long) As Long
    BitConverterToInt16 = MyDotNet.BConverterToInt16(bArray, offset)

End Function
Public Function BitConverterToUInt16(bArray() As Byte, offset As Long) As Long
    BitConverterToUInt16 = MyDotNet.BConverterToUInt16(bArray, offset)

End Function
Public Function BitConverterToUInt32(bArray() As Byte, offset As Long) As Long
    BitConverterToUInt32 = MyDotNet.BConverterToUInt32(bArray, offset)

End Function
Public Function BitConverterToInt32(bArray() As Byte, offset As Long) As Long
    BitConverterToInt32 = MyDotNet.BConverterToInt32(bArray, offset)
End Function
Public Function BitConverterToInt64(bArray() As Byte, offset As Long) As Currency
    BitConverterToInt64 = MyDotNet.BConverterToInt64(bArray, offset)
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
    tablebit = MyDotNet.isTablePresent(MetaDataByteArray, tableindex)
    Dim j As Integer
    'For j = 0 To tableindex
    Do While j < tableindex
        Dim o As Integer
        o = Sizes(j) * iRows(j)
        TableOffset = TableOffset + o
        j = j + 1
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
    GetResolutionScopeValue = MyDotNet.DoRightBitShift(rvalue, 2)
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
GetManifestResourceValue = MyDotNet.DoRightBitShift(manifiestvalue, 2)
    
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
        table = MyDotNet.isTablePresent(gVBNETHeader.EntryPointToken, 24)
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
    GetTypeDefOrRefValue = MyDotNet.DoRightBitShift(codedvalue, 2)
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

