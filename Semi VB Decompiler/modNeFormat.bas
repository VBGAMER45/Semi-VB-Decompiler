Attribute VB_Name = "modNeFormat"
'*********************************************
'modNeFormat
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************
Option Explicit
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
Dim ResourcesTables() As ResourceTableType
Public Sub GetNEHeader()

    NEHeader.Signature = NE_SIGNATURE
    NEHeader.VersionLinker = modPeSkeleton.GetByte
    NEHeader.RevisionLinker = modPeSkeleton.GetByte
    NEHeader.EntryTableOffset = modPeSkeleton.GetWord
    NEHeader.SizeOfEntryTable = modPeSkeleton.GetWord
    NEHeader.CRC = modPeSkeleton.GetDWord
    NEHeader.Flags = modPeSkeleton.GetWord
    NEHeader.SegmentNumberAutomaticDataSegment = modPeSkeleton.GetWord
    NEHeader.InitialSizeHeap = modPeSkeleton.GetWord
    NEHeader.InitialSizeStack = modPeSkeleton.GetWord
    NEHeader.SegmentNumberOffsetCS = modPeSkeleton.GetDWord
    NEHeader.SegmentNumberOffsetSS = modPeSkeleton.GetDWord
    NEHeader.NumberEntriesSegmentTable = modPeSkeleton.GetWord
    NEHeader.NumberEntriesModuleReferenceTable = modPeSkeleton.GetWord
    NEHeader.SizeOfNonResidentNameTable = modPeSkeleton.GetWord
    NEHeader.SegmentTableOffset = modPeSkeleton.GetWord
    NEHeader.ResourceTableFileOffset = modPeSkeleton.GetWord
    NEHeader.ResidentNameTableOffset = modPeSkeleton.GetWord
    NEHeader.ModuleReferenceTableOffset = modPeSkeleton.GetWord
    NEHeader.ImportedNamesTableOffset = modPeSkeleton.GetWord
    NEHeader.NonResidentNameTableOffset = modPeSkeleton.GetDWord
    NEHeader.NumberMovableEntriesInEntryTable = modPeSkeleton.GetWord
    NEHeader.LogicalSectorAlignmentShiftCount = modPeSkeleton.GetWord
    NEHeader.NumberResourceEntries = modPeSkeleton.GetWord
    NEHeader.ExecutableType = modPeSkeleton.GetByte
    NEHeader.Reserved1 = modPeSkeleton.GetDWord
    NEHeader.Reserved2 = modPeSkeleton.GetDWord
End Sub
Sub ProcessNeFile()
    Dim i As Integer
    ReDim SegmentTable(NEHeader.NumberEntriesSegmentTable - 1)
    For i = 0 To NEHeader.NumberEntriesSegmentTable - 1
        SegmentTable(i).LogicalSectorOffset = modPeSkeleton.GetWord
        SegmentTable(i).Length = modPeSkeleton.GetWord
        SegmentTable(i).Flag = modPeSkeleton.GetWord
        SegmentTable(i).MinimumAllocationSizeOfTheSegment = modPeSkeleton.GetWord
    Next
    Exit Sub
'RESOURCE TABLE
Dim g As Integer
    Seek InFileNumber, DosHeader.ExeHeaderPointer + NEHeader.ResourceTableFileOffset + 1
    ReDim ResourcesTables(NEHeader.NumberResourceEntries - 1)
    For i = 0 To NEHeader.NumberResourceEntries - 1
        ResourcesTables(i).AlignmentShiftCountForResourceData = modPeSkeleton.GetWord
        ResourcesTables(i).TypeID = modPeSkeleton.GetWord
        Debug.Print "Typid: " & ResourcesTables(i).TypeID
        If ResourcesTables(i).TypeID = 0 Then Exit For
        ResourcesTables(i).NumberOfResources = modPeSkeleton.GetWord
        ResourcesTables(i).Reserved = modPeSkeleton.GetDWord
        ReDim ResourcesTables(i).ResourceEntries(ResourcesTables(i).NumberOfResources)
        For g = 0 To ResourcesTables(i).NumberOfResources - 1
            ResourcesTables(i).ResourceEntries(g).ResourceDataOffset = modPeSkeleton.GetWord
            ResourcesTables(i).ResourceEntries(g).Length = modPeSkeleton.GetWord
            ResourcesTables(i).ResourceEntries(g).Flag = modPeSkeleton.GetWord
            ResourcesTables(i).ResourceEntries(g).ResourceID = modPeSkeleton.GetWord
            ResourcesTables(i).ResourceEntries(g).Reserved = modPeSkeleton.GetDWord
        Next g
        ResourcesTables(i).Length = modPeSkeleton.GetWord
        Debug.Print "Length: " & ResourcesTables(i).Length
    Next i
End Sub

