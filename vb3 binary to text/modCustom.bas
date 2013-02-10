Attribute VB_Name = "modCustom"
'***********************************
'Jonathan Valentin 2005
'http://www.visualbasiczone.com
'modCustom.bas
'Purpose: To Process external VBX Controls.
'***********************************
Option Explicit

Public Const DOS_SIGNATURE = 23117    '"MZ" = 0x4D5A
Public Const PE_SIGNATURE = 17744    '"PE" + 0x00 = 0x50450000
Public Const NE_SIGNATURE = 17742

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

Private Type SegmentTableEntry
    LogicalSectorOffset As Double 'WORD
    Length As Double 'WORD
    Flag As Double 'WORD
    MinimumAllocationSizeOfTheSegment As Double 'WORD
End Type
Private Type ResourceEntryType
    ResourceDataOffset As Double 'WORD
    Length As Double 'WORD
    Flag As Double 'WORD
    ResourceID As Double 'WORD
    Reserved As Double 'DWORD
End Type

Private Type ResourceTableType
    AlignmentShiftCountForResourceData As Double 'WORD
    TypeID As Double 'WORD
    NumberOfResources  As Double 'WORD
    Reserved As Double 'DWORD
    ResourceEntries() As ResourceEntryType
    Length As Double 'Btye
    Text As String 'Byte
End Type

Dim SegmentTable() As SegmentTableEntry

Private Type ModalType
    usVersion As Double 'VB version used by control
    fl As Double 'Bitfield structure
    pctlproc As Double
    fsClassStyle As Double
    flWndStyle As Double
    cbCtlExtra As Double
    idBmpPalette As Double
    npszDefCtlName As String
    npszClassName As String
    npszParentClassName As String
    nDefProp As Double
    nDefEvent As Double
    nValueProp As Double
    usCtlVersion As Double
End Type
Private VBXModal As ModalType
'typedef struct tagMODEL
'  {
'  USHORT        usVersion;              // VB version used by control
'  FLONG         fl;                     // Bitfield structure
'  PCTLPROC      pctlproc;               // The control proc.
'  FSHORT        fsClassStyle;           // Window class style
'  FLONG         flWndStyle;             // Default window style
'  USHORT        cbCtlExtra;             // # bytes alloc'd for HCTL structure
'  WORD          idBmpPalette;           // BITMAP id for tool palette
'  PSTR          npszDefCtlName;         // Default control name prefix
'  PSTR          npszClassName;          // Visual Basic class name
'  PSTR          npszParentClassName;    // Parent window class if subclassed
'  NPPROPLIST    npproplist;             // Property list
'  NPEVENTLIST   npeventlist;            // Event list
'  BYTE          nDefProp;               // Index of default property
'  BYTE          nDefEvent;              // Index of default event
'  BYTE          nValueProp;             // Index of control value property
'  USHORT        usCtlVersion;           // Identifies the current version of
'                                        // the custom control. The values
'                                        // 1 and 2 are reserved for custom
'                                        // controls created with VB 1.0 and
'                                        // VB 2.0.
'  } MODEL;
Private Type PROPINFOType
    npszName As String
    fl As Double
    offsetData As Double
    infoData As Double
    dataDefault As Double
    npszEnumList As String
    enumMax As Double
End Type
'typedef struct tagPROPINFO
'  {
'  PSTR  npszName;
'  FLONG fl;                     // PF_ flags
'  BYTE  offsetData;             // Offset into static structure
'  BYTE  infoData;               // 0 or _INFO value for bitfield
'  LONG  dataDefault;            // 0 or _INFO value for bitfield
'  PSTR  npszEnumList;           // For type == DT_ENUM, this is
'                                // a near ptr to a string containing
'                                // all the values to be displayed
'                                // in the popup enumeration listbox.
'                                // Each value is an sz, with an
'                                // empty sz indicated the end of list.
'  BYTE  enumMax;                // Maximum legal value for enum.
'  } PROPINFO;
Private Type EVENTINFOType
    npszName As String
    cParms As Double
    cwParms As Double
    npParmTypes As Double
    npszParmProf As String
    fl As Double
End Type
'typedef struct tagEVENTINFO
'  {
'  PSTR          npszName;       // event procedure name suffix
'  USHORT        cParms;         // number of parameters
'  USHORT        cwParms;        // # words of parameters
'  PWORD         npParmTypes;    // list of parameter types
'  PSTR          npszParmProf;   // event parameter profile string
'  FLONG         fl;             // EF_ flags
'  } EVENTINFO;
Dim ResourcesTables() As ResourceTableType
Dim NEHeader As NE_Header
Dim DosHeader As Dos_Header
Dim ErrorFlag As Boolean

Private Sub GetDOSSignature(ByVal F As Integer)

    'Get the first two characters
    DosHeader.Magic = GetWordByFile(F)
    
    'Check for error
    If DosHeader.Magic <> DOS_SIGNATURE Then
        ErrorFlag = True
    End If
    
End Sub
Private Sub GetDOSHeader(ByVal F As Integer)

    'Get DOS header data
        DosHeader.NumBytesLastPage = GetWordByFile(F)
        DosHeader.NumPages = GetWordByFile(F)
        DosHeader.NumRelocates = GetWordByFile(F)
        DosHeader.NumHeaderBlks = GetWordByFile(F)
        DosHeader.NumMinBlks = GetWordByFile(F)
        DosHeader.NumMaxBlks = GetWordByFile(F)
        DosHeader.SSPointer = GetWordByFile(F)
        DosHeader.SPPointer = GetWordByFile(F)
        DosHeader.Checksum = GetWordByFile(F)
        DosHeader.IPPointer = GetWordByFile(F)
        DosHeader.CurrentSeg = GetWordByFile(F)
        DosHeader.RelocTablePointer = GetWordByFile(F)
        DosHeader.Overlay = GetWordByFile(F)
        DosHeader.ReservedW1 = GetWordByFile(F)
        DosHeader.ReservedW2 = GetWordByFile(F)
        DosHeader.ReservedW3 = GetWordByFile(F)
        DosHeader.ReservedW4 = GetWordByFile(F)
        DosHeader.OEMType = GetWordByFile(F)
        DosHeader.OEMData = GetWordByFile(F)
        DosHeader.ReservedW5 = GetWordByFile(F)
        DosHeader.ReservedW6 = GetWordByFile(F)
        DosHeader.ReservedW7 = GetWordByFile(F)
        DosHeader.ReservedW8 = GetWordByFile(F)
        DosHeader.ReservedW9 = GetWordByFile(F)
        DosHeader.ReservedW10 = GetWordByFile(F)
        DosHeader.ReservedW11 = GetWordByFile(F)
        DosHeader.ReservedW12 = GetWordByFile(F)
        DosHeader.ReservedW13 = GetWordByFile(F)
        DosHeader.ReservedW14 = GetWordByFile(F)
        DosHeader.ExeHeaderPointer = GetDWordByFile(F)
        
        'Make sure the potential NE signature location seems reasonable
        If ((DosHeader.ExeHeaderPointer > 4096) Or (DosHeader.ExeHeaderPointer < 64)) Then
            ErrorFlag = True
        End If
        
End Sub
Private Sub GetNESignature(ByVal F As Integer)
    Dim Magic As Double
    'Get the first two characters
    Magic = GetWordByFile(F)
    'Check for error
    If Magic <> NE_SIGNATURE Then
        ErrorFlag = True
    End If
    
End Sub
Private Sub GetNEHeader(ByVal F As Integer)

    NEHeader.signature = NE_SIGNATURE
    NEHeader.VersionLinker = GetByteByFile(F)
    NEHeader.RevisionLinker = GetByteByFile(F)
    NEHeader.EntryTableOffset = GetWordByFile(F)
    NEHeader.SizeOfEntryTable = GetWordByFile(F)
    NEHeader.CRC = GetDWordByFile(F)
    NEHeader.flags = GetWordByFile(F)
    NEHeader.SegmentNumberAutomaticDataSegment = GetWordByFile(F)
    NEHeader.InitialSizeHeap = GetWordByFile(F)
    NEHeader.InitialSizeStack = GetWordByFile(F)
    NEHeader.SegmentNumberOffsetCS = GetDWordByFile(F)
    NEHeader.SegmentNumberOffsetSS = GetDWordByFile(F)
    NEHeader.NumberEntriesSegmentTable = GetWordByFile(F)
    NEHeader.NumberEntriesModuleReferenceTable = GetWordByFile(F)
    NEHeader.SizeOfNonResidentNameTable = GetWordByFile(F)
    NEHeader.SegmentTableOffset = GetWordByFile(F)
    NEHeader.ResourceTableFileOffset = GetWordByFile(F)
    NEHeader.ResidentNameTableOffset = GetWordByFile(F)
    NEHeader.ModuleReferenceTableOffset = GetWordByFile(F)
    NEHeader.ImportedNamesTableOffset = GetWordByFile(F)
    NEHeader.NonResidentNameTableOffset = GetDWordByFile(F)
    NEHeader.NumberMovableEntriesInEntryTable = GetWordByFile(F)
    NEHeader.LogicalSectorAlignmentShiftCount = GetWordByFile(F)
    NEHeader.NumberResourceEntries = GetWordByFile(F)
    NEHeader.ExecutableType = GetByteByFile(F)
    NEHeader.Reserved1 = GetDWordByFile(F)
    NEHeader.Reserved2 = GetDWordByFile(F)
End Sub
Private Sub ProcessNeFile(ByVal F As Integer)
    Dim i As Integer
    ReDim SegmentTable(NEHeader.NumberEntriesSegmentTable - 1)
    For i = 0 To NEHeader.NumberEntriesSegmentTable - 1
        SegmentTable(i).LogicalSectorOffset = GetWordByFile(F)
        SegmentTable(i).Length = GetWordByFile(F)
        SegmentTable(i).Flag = GetWordByFile(F)
        SegmentTable(i).MinimumAllocationSizeOfTheSegment = GetWordByFile(F)
    Next
    Exit Sub
'RESOURCE TABLE
Dim g As Integer
    Seek F, DosHeader.ExeHeaderPointer + NEHeader.ResourceTableFileOffset + 1
    ReDim ResourcesTables(NEHeader.NumberResourceEntries - 1)
    For i = 0 To NEHeader.NumberResourceEntries - 1
        ResourcesTables(i).AlignmentShiftCountForResourceData = GetWordByFile(F)
        ResourcesTables(i).TypeID = GetWordByFile(F)
        'Debug.Print "Typeid: " & ResourcesTables(i).TypeID
        If ResourcesTables(i).TypeID = 0 Then Exit For
        ResourcesTables(i).NumberOfResources = GetWordByFile(F)
        ResourcesTables(i).Reserved = GetDWordByFile(F)
        ReDim ResourcesTables(i).ResourceEntries(ResourcesTables(i).NumberOfResources)
        For g = 0 To ResourcesTables(i).NumberOfResources - 1
            ResourcesTables(i).ResourceEntries(g).ResourceDataOffset = GetWordByFile(F)
            ResourcesTables(i).ResourceEntries(g).Length = GetWordByFile(F)
            ResourcesTables(i).ResourceEntries(g).Flag = GetWordByFile(F)
            ResourcesTables(i).ResourceEntries(g).ResourceID = GetWordByFile(F)
            ResourcesTables(i).ResourceEntries(g).Reserved = GetDWordByFile(F)
        Next g
        ResourcesTables(i).Length = GetWordByFile(F)
      '  Debug.Print "Length: " & ResourcesTables(i).Length
    Next i
End Sub
Public Sub ProccesVBXControl(ByVal strFilename As String)
    Dim F As Integer
    F = FreeFile
    ErrorFlag = False
    Open strFilename For Binary Access Read As #F
        Seek F, 4616
        'Get the Dos signature
        'Call GetDOSSignature(F)
        'If ErrorFlag = True Then
        '    Exit Sub
        'End If
        '*******************************
        'The DOS header follows the DOS signature
        '*******************************
        'Get the Dos header
        'Call GetDOSHeader(F)
        'If ErrorFlag = True Then
        '    Exit Sub
        'End If
        
        'Get NE signature
        'Call GetNESignature(F)
        'If ErrorFlag = True Then
        '    Exit Sub
        'End If
        'Ne Header
        'Call GetNEHeader(F)
        'Call ProcessNeFile(F)
        'Get VBX Modal Type Info
        Call GetModalType(F)
        
    Close #F
End Sub
Private Sub GetModalType(ByVal F As Integer)
    Dim dDouble As Double
    VBXModal.usVersion = GetWordByFile(F)
    MsgBox VBXModal.usVersion
    Get #F, , dDouble
    VBXModal.fl = dDouble
    Get #F, , dDouble
    VBXModal.pctlproc = dDouble
    VBXModal.fsClassStyle = GetDWordByFile(F)
    Get #F, , dDouble
    VBXModal.flWndStyle = dDouble
    VBXModal.cbCtlExtra = GetWordByFile(F)
    VBXModal.idBmpPalette = GetWordByFile(F)
    VBXModal.npszDefCtlName = GetUntilNull(F)
    MsgBox VBXModal.npszDefCtlName
    VBXModal.npszClassName = GetUntilNull(F)
    VBXModal.npszParentClassName = GetUntilNull(F)
    Call GetPropList(F)
    Call GetEventList(F)
    VBXModal.nDefProp = GetByteByFile(F)
    VBXModal.nDefEvent = GetByteByFile(F)
    VBXModal.nValueProp = GetByteByFile(F)
    VBXModal.usCtlVersion = GetWordByFile(F)
End Sub
Private Sub GetPropList(ByVal F As Integer)

End Sub
Private Sub GetEventList(ByVal F As Integer)

End Sub
Public Function GetDWordByFile(ByVal FileNum As Integer) As Double
'**********************************
'Purpose: Gets a DWORD from a file
'**********************************
   GetDWordByFile# = GetWordByFile(FileNum)
   GetDWordByFile# = GetDWordByFile# + 65536# * GetWordByFile(FileNum)
      
End Function

Public Function GetWordByFile(ByVal FileNum As Integer) As Double
'**********************************
'Purpose: Gets a Word from a file
'**********************************
    GetWordByFile# = GetByteByFile(FileNum)
    GetWordByFile# = GetWordByFile# + 256# * GetByteByFile(FileNum)
       
End Function

Public Function GetByteByFile(ByVal FileNum As Integer) As Byte
'**********************************
'Purpose: Gets a byte from a file
'**********************************
    Dim DataByte As Byte
    'Read the data
    Get #FileNum, , DataByte
    
    'Return it
    GetByteByFile = DataByte
      
End Function
Public Function GetUntilNull(FileNum As Integer) As String
    '*****************************
    'Purpose to get a null termintated string
    '*****************************
    Dim aList() As Byte
    Dim K As Byte
    K = 255
    ReDim aList(0)
    Do Until K = 0
        Get FileNum, , K
        ReDim Preserve aList(UBound(aList) + 1)
        aList(UBound(aList)) = K
    Loop
    Dim i As Long
    Dim Final As String
    For i = 1 To UBound(aList) - 1
        Final = Final & Chr(aList(i))
    Next i
    
    GetUntilNull = Final
End Function

Public Function ProccessCustom(ByVal F As Integer, ByVal Opcode As Byte) As Long
    On Error GoTo errHandle
        Dim bData As Byte
        Dim iData As Integer
    Select Case Opcode
        Case 255
       
            Do
                Get #F, , bData

                If bData = 1 Then
                   
                ElseIf bData = 4 Then
                    gIdentSpaces = gIdentSpaces - 1
                    Call frmMain.AddText("End")
                    gFormDone = True
                ElseIf bData = 3 Then
                    gIdentSpaces = gIdentSpaces - 1
                    Call frmMain.AddText("End")
                ElseIf bData = 2 Then
                    gIdentSpaces = gIdentSpaces - 1
                    Call frmMain.AddText("End")
                End If
          
            Loop While bData <> 0 And bData < 6
            Seek F, Loc(F)
        Case Else
            
            Call AddError("Error_Unknown Opcode_ProcessCustomControl: Opcode: " & Opcode & " Offset: " & Loc(F))
            bFirstFF = True
            Seek F, frmMain.ReturnControlEnd
    End Select
    
Exit Function
errHandle:
    Call AddError("Error_ProcessCustomControl: Opcode: " & Opcode & " Offset: " & Loc(F) & " Description= " & Err.Description)
            
End Function
