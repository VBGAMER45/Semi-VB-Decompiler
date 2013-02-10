Attribute VB_Name = "modGlobals"
'*********************************************
'modGlobals
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************
'Notes
'*********************************************
'"a" - means it is an Address
'"o" - means it is a relative Offset
'"Unknown" - self explanatory
'"Flag" - Variable Unknown Property
'"Const" - Constant Unknown Property
'"Address" - Unknown Address
'*********************************************
Option Explicit
Const MAX_PATH = 260

Type VBHeader
    signature               As String * 4  '00h 00d
    'VB5! identifier &quot;VB5!&quot;
    
   RuntimeBuild                 As Integer     '04h 04d
     'RuntimeBuild
  LanguageDLL             As String * 14 '06h 06d
    'Language DLL name. _
     0x2A meaning default or null terminated string.
  
  BackupLanguageDLL       As String * 14 '14h 20d
    'Backup Language DLL name. _
     0x7F meaning default or null terminated string. _
     Changing values do not effect working status of an exe.
  
  RuntimeDLLVersion       As Integer     '22h 34d
    'Run-time DLL version
  
  LanguageID              As Long        '24h 36d
  
  BackupLanguageID        As Long        '28h 40d
    'Backup Language ID &#40;only when Language DLL exists&#41;
  
  aSubMain                As Long        '2Ch 44d
    'Address to Sub Main&#40;&#41; code _
     &#40;If 0000 0000 then it's a load form call&#41;
  
  aProjectInfo            As Long        '30h 48d
    
  fMDLIntObjs                 As Long    '34h 52d

  fMDLIntObjs2               As Long        '38h 56d

  ThreadFlags           As Long        '3Ch 60d

  ThreadCount               As Long        '40h 64d
  
  
  FormCount               As Integer     '44h 68d
  
  ExternalComponentCount  As Integer     '46h 70d
    'Number of external components &#40;eg. winsock&#41; referenced
    
  ThunkCount As Long
  
  aGuiTable               As Long        '4Eh 78d
  aExternalComponentTable As Long        '52h 82d
  'aProjectDescription     As Long        '56h 86d
  aComRegisterData         As Long
  
  oProjectExename         As Long        '5Ah 90d
  oProjectTitle           As Long        '5Eh 94d
  oHelpFile               As Long        '62h 98d
  oProjectName            As Long        '66h 102d
End Type

'* Thread Flags:
'+-------+----------------+--------------------------------------------------------+
'| Value | Name           | Description                                            |
'+-------+----------------+--------------------------------------------------------+
'|  0x01 | ApartmentModel | Specifies multi-threading using an apartment model     |
'|  0x02 | RequireLicense | Specifies to do license validation (OCX only)          |
'|  0x04 | Unattended     | Specifies that no GUI elements should be initialized   |
'|  0x08 | SingleThreaded | Specifies that the image is single-threaded            |
'|  0x10 | Retained       | Specifies to keep the file in memory (Unattended only) |
'+-------+----------------+--------------------------------------------------------+
'ex: A value of 0x15 specifies a multi-threaded, memory-resident ActiveX Object with no GUI


'* MDL Internal Object Flags
'+---------+------------+---------------+
'| Ctrl ID |      Value | Object Name   |
'+---------+------------+---------------+
'|                           First Flag |
'+---------+------------+---------------+
'|    0x00 | 0x00000001 | PictureBox    |
'|    0x01 | 0x00000002 | Label         |
'|    0x02 | 0x00000004 | TextBox       |
'|    0x03 | 0x00000008 | Frame         |
'|    0x04 | 0x00000010 | CommandButton |
'|    0x05 | 0x00000020 | CheckBox      |
'|    0x06 | 0x00000040 | OptionButton  |
'|    0x07 | 0x00000080 | ComboBox      |
'|    0x08 | 0x00000100 | ListBox       |
'|    0x09 | 0x00000200 | HScrollBar    |
'|    0x0A | 0x00000400 | VScrollBar    |
'|    0x0B | 0x00000800 | Timer         |
'|    0x0C | 0x00001000 | Print         |
'|    0x0D | 0x00002000 | Form          |
'|    0x0E | 0x00004000 | Screen        |
'|    0x0F | 0x00008000 | Clipboard     |
'|    0x10 | 0x00010000 | Drive         |
'|    0x11 | 0x00020000 | Dir           |
'|    0x12 | 0x00040000 | FileListBox   |
'|    0x13 | 0x00080000 | Menu          |
'|    0x14 | 0x00100000 | MDIForm       |
'|    0x15 | 0x00200000 | App           |
'|    0x16 | 0x00400000 | Shape         |
'|    0x17 | 0x00800000 | Line          |
'|    0x18 | 0x01000000 | Image         |
'|    0x19 | 0x02000000 | Unsupported   |
'|    0x1A | 0x04000000 | Unsupported   |
'|    0x1B | 0x08000000 | Unsupported   |
'|    0x1C | 0x10000000 | Unsupported   |
'|    0x1D | 0x20000000 | Unsupported   |
'|    0x1E | 0x40000000 | Unsupported   |
'|    0x1F | 0x80000000 | Unsupported   |
'+---------+------------+---------------+
'|                          Second Flag |
'+---------+------------+---------------+
'|    0x20 | 0x00000001 | Unsupported   |
'|    0x21 | 0x00000002 | Unsupported   |
'|    0x22 | 0x00000004 | Unsupported   |
'|    0x23 | 0x00000008 | Unsupported   |
'|    0x24 | 0x00000010 | Unsupported   |
'|    0x25 | 0x00000020 | DataQuery     |
'|    0x26 | 0x00000040 | OLE           |
'|    0x27 | 0x00000080 | Unsupported   |
'|    0x28 | 0x00000100 | UserControl   |
'|    0x29 | 0x00000200 | PropertyPage  |
'|    0x2A | 0x00000400 | Document      |
'|    0x2B | 0x00000800 | Unsupported   |
'+---------+------------+---------------+
'ex: A value of 0x30F000 (the so called "static binary constant on most sites") actually means to initialize the Print, Form, Screen, ClipBoard Objects (0xF000) as well as the Drive/Dir Objects (0x30000). This is the default on VB projects because those objects can always be accessed from a module (ie, they are not graphic, except Forms, wich can always be created)

'COM Data Types
'
Type tCOMRegData
  oRegInfo                As Long    ' 0x00 (00d) Offset to COM Interfaces Info
  oNTSProjectName         As Long    ' 0x04 (04d) Offset to Project/Typelib Name
  oNTSHelpDirectory       As Long    ' 0x08 (08d) Offset to Help Directory
  oNTSProjectDescription  As Long    ' 0x0C (12d) Offset to Project Description
  uuidProjectClsId(15)    As Byte    ' 0x10 (16d) CLSID of Project/Typelib
  lTlbLcid                As Long    ' 0x20 (32d) LCID of Type Library
  iPadding1               As Integer ' 0x24 (36d)
  iTlbVerMajor            As Integer ' 0x26 (38d) Typelib Major Version
  iTlbVerMinor            As Integer ' 0x28 (40d) Typelib Minor Version
  iPadding2               As Integer ' 0x2A (42d)
  lPadding3               As Long    ' 0x2C (44d)
                                     ' 0x30 (48d) <- Structure Size
End Type

Type tCOMRegInfo
  oNextObject          As Long    ' 0x00 (00d) Offset to COM Interfaces Info
  oObjectName          As Long    ' 0x04 (04d) Offset to Object Name
  oObjectDescription   As Long    ' 0x08 (08d) Offset to Object Description
  lInstancing          As Long    ' 0x0C (12d) Instancing Mode
  lObjectID            As Long    ' 0x10 (16d) Current Object ID in the Project
  uuidObjectClsID(15)  As Byte    ' 0x14 (20d) CLSID of Object
  fIsInterface         As Long    ' 0x24 (36d) Specifies if the next CLSID is valid
  oObjectClsID         As Long    ' 0x28 (40d) Offset to CLSID of Object Interface
  oControlClsID        As Long    ' 0x2C (44d) Offset to CLSID of Control Interface
  fIsControl           As Long    ' 0x30 (48d) Specifies if the CLSID above is valid
  lMiscStatus          As Long    ' 0x34 (52d) OLEMISC Flags (see MSDN docs)
  fClassType           As Byte    ' 0x38 (56d) Class Type
  fObjectType          As Byte    ' 0x39 (57d) Flag identifying the Object Type
  iToolboxBitmap32     As Integer ' 0x3A (58d) Control Bitmap ID in Toolbox
  iDefaultIcon         As Integer ' 0x3C (60d) Minimized Icon of Control Window
  fIsDesigner          As Integer ' 0x3E (62d) Specifies whether this is a Designer
  oDesignerData        As Long    ' 0x40 (64d) Offset to Designer Data
                                  ' 0x44 (68d) <-- Structure Size
End Type
'Object Type part of tCOMRegInfo
'+-------+---------------+-------------------------------------------+
'| Value | Name          | Description                               |
'+-------+---------------+-------------------------------------------+
'|  0x02 | Designer      | A Visual Basic Designer for an Add.in     |
'|  0x10 | Class Module  | A Visual Basic Class                      |
'|  0x20 | User Control  | A Visual Basic ActiveX User Control (OCX) |
'|  0x80 | User Document | A Visual Basic User Document              |
'+-------+---------------+-------------------------------------------+

Private Type tProjectInfo

  signature As Long                            ' 0x00
  aObjectTable As Long                         ' 0x04
  Null1 As Long                                ' 0x08
  aStartOfCode As Long                         ' 0x0C
  aEndOfCode As Long                           ' 0x10
  Flag1 As Long                                ' 0x14
  ThreadSpace As Long                          ' 0x18
  aVBAExceptionhandler  As Long                ' 0x1C
  aNativeCode As Long                          ' 0x20
  oProjectLocation As Integer                  ' 0x24
  Flag2 As Integer                             ' 0x26
  Flag3 As Integer                             ' 0x28

  OriginalPathName(MAX_PATH * 2) As Byte       ' 0x2A
  NullSpacer As Byte                           ' 0x233
  aExternalTable As Long                       ' 0x234
  ExternalCount As Long                        ' 0x238

' Size 0x23C
End Type

Private Type tObject
    aObjectInfo As Long         ' 0x00
    Const1 As Long              ' 0x04
    aPublicBytes As Long        ' 0x08 (08d) Pointer to Public Variable Size integers
    aStaticBytes As Long        ' 0x0C (12d) Pointer to Static Variables Struct
    aModulePublic As Long       ' 0x10 (16d) Memory Pointer to Public Variables
    aModuleStatic As Long       ' 0x14 (20d) Pointer to Static Variables
    aObjectName As Long         ' 0x18  NTS
    ProcCount As Long           ' 0x1C events, funcs, subs
    aProcNamesArray As Long     ' 0x20 when non-zero
    oStaticVars As Long         ' 0x24 (36d) Offset to Static Vars from aModuleStatic
    ObjectType As Long          ' 0x28
    Null3 As Long               ' 0x2C
                                ' 0x30  <-- Structure Size
End Type

'tObject.ObjectTyper Properties...
'#########################################################
'form&#58;              0000 0001 1000 0000 1000 0011 --&gt; 18083
'                   0000 0001 1000 0000 1010 0011 --&gt; 180A3
'                   0000 0001 1000 0000 1100 0011 --&gt; 180C3
'module&#58;            0000 0001 1000 0000 0000 0001 --&gt; 18001
'                   0000 0001 1000 0000 0010 0001 --&gt; 18021
'class&#58;             0001 0001 1000 0000 0000 0011 --&gt; 118003
'                   0001 0011 1000 0000 0000 0011 --&gt; 138003
'                   0000 0001 1000 0000 0010 0011 --&gt; 18023
'                   0000 0001 1000 1000 0000 0011 --&gt; 18803
'                   0001 0001 1000 1000 0000 0011 --&gt; 118803
'usercontrol&#58;       0001 1101 1010 0000 0000 0011 --&gt; 1DA003
'                  0001 1101 1010 0000 0010 0011 --&gt; 1DA023
'                  0001 1101 1010 1000 0000 0011 --&gt; 1DA803
'propertypage&#58;      0001 0101 1000 0000 0000 0011 --&gt; 158003
'                      | ||     |  |    | |    |
'&#91;moog&#93;                | ||     |  |    | |    |
'HasPublicInterface ---+ ||     |  |    | |    |
'HasPublicEvents --------+|     |  |    | |    |
'IsCreatable/Visible? ----+     |  |    | |    |
'Same as &quot;HasPublicEvents&quot; -----+  |    | |    |
'&#91;aLfa&#93;                         |  |    | |    |
'usercontrol &#40;1&#41; ---------------+  |    | |    |
'ocx/dll &#40;1&#41; ----------------------+    | |    |
'form &#40;1&#41; ------------------------------+ |    |
'vb5 &#40;1&#41; ---------------------------------+    |
'HasOptInfo &#40;1&#41; -------------------------------+
'                                              |
'module&#40;0&#41; ------------------------------------+

Public Type tObjectInfo
    Flag1 As Integer       ' 0x00
    ObjectIndex As Integer ' 0x02
    aObjectTable As Long   ' 0x04
    Null1 As Long          ' 0x08
    aSmallRecord   As Long ' 0x0C  when it is a module this value is -1 [better name?]
    Const1 As Long         ' 0x10
    Null2 As Long          ' 0x14
    aObject As Long        ' 0x18
    RunTimeLoaded  As Long ' 0x1C [can someone verify this?]
    NumberOfProcs  As Long ' 0x20
    aProcTable As Long     ' 0x24
    iConstantsCount As Integer '0x28 (40d) Number of Constants
    iMaxConstants   As Integer '0x2A (42d) Maximum Constants to allocate.
    Flag5 As Long          ' 0x2C
    Flag6 As Integer       ' 0x30
    Flag7 As Integer       ' 0x32
    aConstantPool As Long  ' 0x34
                           ' 0x38 <-- Structure Size
                           'the rest is optional items[OptionalObjectInfo]
End Type
Private Type tObjectTable
    lNull1 As Long          ' 0x00 (00d)
    aExecProj As Long       ' 0x04 (04d) Pointer to a memory structure
    aProjectInfo2 As Long   ' 0x08 (08d) Pointer to Project Info 2
    Const1 As Long          ' 0x0C
    Null2 As Long           ' 0x10
    lpProjectObject As Long ' 0x14
    Flag1 As Long           ' 0x18
    Flag2 As Long           ' 0x1C
    Flag3 As Long           ' 0x20
    Flag4 As Long           ' 0x24
    fCompileType As Integer ' 0x28 (40d) Internal flag used during compilation
    ObjectCount1 As Integer ' 0x2A
    iCompiledObjects As Integer ' 0x2C (44d) Number of objects compiled.
    iObjectsInUse As Integer ' 0x2E (46d) Updated in the IDE to correspond the total number ' but will go up or down when initializing/unloading modules.
    aObject As Long         ' 0x30
    Null3 As Long           ' 0x34
    Null4 As Long           ' 0x38
    Null5 As Long           ' 0x3C
    aProjectName As Long    ' 0x40      NTS
    LangID1  As Long        ' 0x44
    LangID2  As Long        ' 0x48
    Null6  As Long          ' 0x4C
    Const3  As Long         ' 0x50
                            ' 0x54
End Type
Type ExternalTable
   Flag As Long        '0x00
   aExternalLibrary As Long  '0x04
End Type

Type ExternalLibrary
   aLibraryName As Long     '0x00   points to NTS
   aLibraryFunction As Long '0x04   points to NTS
End Type

Public Type tEventLink

    Const1 As Integer        ' 0x00
    CompileType As Byte      ' 0x02
    aEvent As Long           ' 0x03
    PushCmd As Byte          ' 0x07
    pushAddress As Long      ' 0x08
    Const As Byte            ' 0x0C
                             ' 0x0D&lt;-- Structure Size
End Type
Private Type tEventTable
    Null1 As Long                                  ' 0x00
    aControl As Long                               ' 0x04
    aObjectInfo As Long                            ' 0x08
    aQueryInterface As Long                        ' 0x0C
    aAddRef As Long                                ' 0x10
    aRelease As Long                                ' 0x14
    'aEventPointer() As Long
    'aEventPointer(aControl.EventCount - 1) As Long ' 0x18
End Type
Global taEventPointer() As Long

Private Type tOptionalObjectInfo ' if &#40;&#40;tObject.ObjectType AND &amp;H80&#41;=&amp;H80&#41;

    fDesigner As Long              ' 0x00 (0d) If this value is 2 then this object is a designer
    aObjectCLSID As Long           ' 0x04
    Null1 As Long                  ' 0x08
    aGuidObjectGUI As Long         ' 0x0C
    lObjectDefaultIIDCount As Long ' 0x10  01 00 00 00
    aObjectEventsIIDTable As Long  ' 0x14
    lObjectEventsIIDCount As Long  ' 0x18
    aObjectDefaultIIDTable As Long ' 0x1C
    ControlCount As Long           ' 0x20
    aControlArray As Long          ' 0x24
    iEventCount As Integer         ' 0x28 (40d) Number of Events
    iPCodeCount As Integer         ' 0x2C
    oInitializeEvent As Integer    ' 0x2C (44d) Offset to Initialize Event from aMethodLinkTable
    oTerminateEvent As Integer     ' 0x2E (46d) Offset to Terminate Event from aMethodLinkTable
    aEventLinkArray As Long        ' 0x30  Pointer to pointers of MethodLink
    aBasicClassObject As Long      ' 0x34 Pointer to an in-memory
    Null3 As Long                  ' 0x38
    Flag2 As Long                  ' 0x3C usually null
                                   ' 0x40 &lt;-- Structure size
End Type
Public Type tEventPointer
    Const1 As Byte      ' 0x00
    Flag1 As Long       ' 0x01
    Const2 As Long      ' 0x05
    Const3 As Byte      ' 0x09
    aEvent As Long      ' 0x0A
                        ' 0x0E &lt;-- Structure Size
End Type

Public Type tCodeInfo
    aObjectInfo As Long     ' 0x00
    Flag1 As Integer        ' 0x04
    Flag2 As Integer        ' 0x06
    CodeLength As Integer   ' 0x08
    Flag3 As Long           ' 0x0A
    Flag4 As Integer        ' 0x0E
    Null1 As Integer        ' 0x10
    Flag5 As Long           ' 0x12
    Flag6 As Integer        ' 0x16
                            ' 0x18  &lt;-- Structure Size
End Type

Private Type tControl
    Flag1 As Integer        ' 0x00
    EventCount As Integer   ' 0x02
    Flag2 As Long           ' 0x04
    aGUID As Long           ' 0x08
    Index As Integer        ' 0x0C
    Const1 As Integer       ' 0x0E
    Null1 As Long           ' 0x10
    Null2 As Long           ' 0x14
    aEventTable As Long     ' 0x18
    Flag3 As Byte           ' 0x1C
    Const2 As Byte          ' 0x1D
    Const3 As Integer       ' 0x1E
    aName As Long           ' 0x20
    Index2 As Integer       ' 0x24
    Const1Copy As Integer   ' 0x26
                            ' 0x28  &lt;-- Structure Size
End Type



Type tGuiTable
  lStructSize          As Long ' 0x00 (00d) Total size of this structure
  uuidObjectGUI(15)    As Byte ' 0x04 (04d) UUID of Object GUI
  Unknown1             As Long ' 0x14 (20d)
  Unknown2             As Long ' 0x18 (24d)
  Unknown3             As Long ' 0x1C (28d)
  Unknown4             As Long ' 0x20 (32d)
  lObjectID            As Long ' 0x24 (36d) Current Object ID in the Project
  Unknown5             As Long ' 0x28 (40d)
  fOLEMisc             As Long ' 0x2C (44d) OLEMisc Flags
  uuidObject(15)       As Byte ' 0x30 (48d) UUID of Object
  Unknown6             As Long ' 0x40 (64d)
  Unknown7             As Long ' 0x44 (68d)
  aFormPointer         As Long ' 0x48 (72d) Pointer to GUI Object Info
  Unknown8             As Long ' 0x4C (76d)
                               ' 0x50 (80d) <- Structure Size
End Type





Private Type typeProcedureList
    strParent As String
    strProcedureName As String
End Type

'Globals begin
Global gProcedureList() As typeProcedureList
Global gVBHeader As VBHeader



Global gProjectInfo As tProjectInfo
Global gObjectTable As tObjectTable
Global gObject() As tObject
Global gObjectInfo As tObjectInfo


Global gOptionalObjectInfo As tOptionalObjectInfo

Global gControl() As tControl
Global gEventTable() As tEventTable


Global gGuiTable() As tGuiTable
Global gObjectNameArray() As String

Global gObjectInfoHolder() As tObjectInfo
Private Type tObjectOffsetType
    Address As Long
    ObjectName As String
End Type

'Options


Private Type typeControlName
    strParentForm As String
    strControlName As String
    strGuid As String
    bControlImage As Byte
End Type
Global gControlNameArray() As typeControlName


'For Controls



Private Type ImportListtype
    strName As String
    strGuid As String
    strLib As String
End Type
Global ImportList() As ImportListtype



'Variables for .vbp file
Global ProjectExename As String                     ' Project exename. MaxLength: 0x104 (260d)
Global ProjectTitle As String                       ' Project title. MaxLength: 0x28 (40d)
Global HelpFile As String                           ' Helpfile. MaxLength: 0x28 (40d)
Global ProjectName As String                        ' Project name. MaxLength: 0x104 (260d)
Global ProjectDescription As String




Global gDllProject As Boolean 'Is it a dll project?

Global gVB6App As Boolean
Global gVB5App As Boolean
Global gVB4App As Boolean

Public Const cQuote As String = """"


Public Type ControlHeader
   'Length As Integer
    'unknown As Integer
    Length As Long
    cId As Byte 'Used To link events
    cName As String
    un2 As Byte
    cType As Byte
End Type
Public Type ArrayTestType
    Length As Integer
    uni As Byte
    arrayflag As Byte
End Type
Public Type ControlArrayHeader
    Length As Integer
    un1 As Byte
   ' Length As Long
    arrayflag As Integer
    cId As Byte
    un2 As Byte
    cName As String
    un3 As Byte
    cType As Byte
End Type
'Control Sepeartor Constatns
Public Const vbFormNewChildControl = 511 'FF01
Public Const vbFormExistingChildControl = 767 'FF02
Public Const vbFormChildControl = 1023 'FF03
Public Const vbFormEnd = 1279 'FF04
Public Const vbFormMenu = 1535 'FF05
'Used in cType
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
    vbform = 13
    vbDriveListbox = 16
    vbDirectoryListbox = 17
    vbFileListbox = 18
    vbmenu = 19
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


Public Function GetUntilNull(FileNum As Long) As String
    '*****************************
    'Purpose to get a null termintated string
    '*****************************
    Dim aList() As Byte
    Dim k As Byte
    k = 255
    ReDim aList(0)
    Do Until k = 0
        Get FileNum, , k
        ReDim Preserve aList(UBound(aList) + 1)
        aList(UBound(aList)) = k
        'MsgBox k
    Loop
    Dim i As Integer
    Dim Final As String
    For i = 1 To UBound(aList) - 1
        Final = Final & Chr$(aList(i))
      
    Next i
    
    GetUntilNull = Final
End Function

