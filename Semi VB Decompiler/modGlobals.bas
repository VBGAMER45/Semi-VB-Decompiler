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
Public Const Version As String = "0.09" 'Current Version of Semi VB Decompiler

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

Type tDesignerInfo
  uuidDesigner(15)       As Byte    '0x00 (00d)                           CLSID of the Addin/Designer
  lStructSize            As Long    '0x10 (16d)                           Total Size of the next fields
  
  iSizeAddinRegKey       As Integer '0x14 (20d)
  sAddinRegKey           As String  '0x16 (22d)                           Registry Key of the Addin
  
  iSizeAddinName         As Integer '0x16 (22d) + iSizeAddinRegKey
  sAddinName             As String  '0x18 (24d) + iSizeAddinRegKey        Friendly Name of the Addin
  
  iSizeAddinDescription  As Integer '0x18 (24d) + iSizeAddinRegKey _
                                                + iSizeAddinName
  iAddinDescription      As String  '0x1A (26d) + iSizeAddinRegKey _
                                                + iSizeAddinName          Description of Addin
  
  lLoadBehaviour         As Long    '0x1A (26d) + iSizeAddinRegKey _
                                                + iSizeAddinName _
                                                + iSizeAddinDescription   CLSID of Object
  
  iSizeSatelliteDLL      As Integer '0x1E (30d) + iSizeAddinRegKey _
                                                + iSizeAddinName _
                                                + iSizeAddinDescription
  sSatelliteDLL          As String  '0x20 (32d) + iSizeAddinRegKey _
                                                + iSizeAddinName _
                                                + iSizeAddinDescription   SatelliteDLL, if specified
  
  iSizeAdditionalRegKey  As Integer '0x20 (32d) + iSizeAddinRegKey _
                                                + iSizeAddinName _
                                                + iSizeAddinDescription _
                                                + iSizeSatteliteDLL
  sAdditionalRegKey      As String  '0x22 (34d) + iSizeAddinRegKey _
                                                + iSizeAddinName _
                                                + iSizeAddinDescription _
                                                + iSizeSatteliteDLL       Extra Registry Key, if specified
  
  lCommandLineSafe       As Long    '0x22 (34d) + iSizeAddinRegKey _
                                                + iSizeAddinName _
                                                + iSizeAddinDescription _
                                                + iSizeSatteliteDLL _
                                                + iSizeAdditionalRegKey   Specifies a GUI-less Addin if 1
                                    '0x14 + lStructSize  <-- Structure Size
End Type

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
    index As Integer        ' 0x0C
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


Public Type tComponent
  StructLength   As Long
  oUuid         As Long
  l2             As Long
  l3             As Long
  l4             As Long
  l5             As Long
  l6             As Long
  GUIDoffset     As Long
  GUIDlength     As Long
  l7             As Long
  FileNameOffset As Long
  SourceOffset   As Long
  NameOffset     As Long
End Type
'If GUIDlength = -1 then there is no oUUID
'If GUIDlength = 72 then read a unicode UUID

Type tProjectInfo2
  lNull1                  As Long ' 0x00 (00d)
  aObjectTable            As Long ' 0x04 (04d) Pointer to Object Table
  lConst1                 As Long ' 0x08 (08d)
  lNull2                  As Long ' 0x0C (12d)
  aObjectDescriptorTable  As Long ' 0x10 (16d) Pointer to a table of ObjectDescriptors
  lNull3                  As Long ' 0x14 (20d)
  aNTSPrjDescription      As Long ' 0x18 (24d) Pointer to Project Description
  aNTSPrjHelpFile         As Long ' 0x1C (28d) Pointer to Project Help File
  lConst2                 As Long ' 0x20 (32d)
  lHelpContextID          As Long ' 0x24 (36d) Project Help Context ID
                                  ' 0x28 (40d) <- Structure size
End Type

Type ObjectDescriptor
  lNull1      As Long '0x00 (00d)
  aObjectInfo As Long '0x04 (04d) Pointer to Object Info
  lConst1     As Long '0x08 (08d)
  lNull2      As Long '0x0C (12d)
  lFlag1      As Long '0x10 (16d)
  lNull3      As Long '0x14 (20d)
  aUnknown1   As Long '0x18 (24d)
  lNull4      As Long '0x1C (28d)
  aUnknown2   As Long '0x20 (32d)
  aUnknown3   As Long '0x24 (36d)
  aUnknown4   As Long '0x28 (40d)
  lNull5      As Long '0x2C (44d)
  lNull6      As Long '0x30 (48d)
  lNull7      As Long '0x34 (52d)
  lFlag2      As Long '0x38 (56d)
  fObjectType As Long '0x3C (60d) Flags for this Object
                      '0x40 (64d) <- Structure Size
End Type

Type MethodLinkNative
  jmpOpCode As Byte '0x0 (0d)
  jmpoffset As Long '0x1 (1d) jmp <address>  ; <address> = <currentoffset> + <jmpOffset> + 5
                    '0x5 (5d) <-- Structure Size
End Type

Type MethodLinkPCode
  xorOpCode   As Integer '0x0 (00d) xor eax, eax
  movOpCode   As Byte    '0x2 (02d)
  movAddress  As Long    '0x3 (03d) mov edx, <movAddress>
  pushOpCode  As Byte    '0x7 (07d)
  pushAddress As Long    '0x8 (08d) push <pushAddress>
  retOpCode   As Byte    '0xC (12d) ret
                         '0xD (13d) <-- Structure Size
End Type

Type GUIObjectInfo
  lUnknown1            As Long ' 0x00 (00d)
  bUnknown2            As Byte ' 0x04 (04d)
  guidObjectGUI(15)    As Byte ' 0x05 (05d) GUID of this ObjectGUI
  uuidUnknown1(15)     As Byte ' 0x15 (21d)
  guidCOMEventsIID(15) As Byte ' 0x25 (37d) GUID of this object EventsIID
  lUnknown3            As Long ' 0x35 (53d)
  lUnknown4            As Long ' 0x39 (57d)
  lUnknown5            As Long ' 0x3D (61d)
  lUnknown6            As Long ' 0x41 (65d)
  lUnknown7            As Long ' 0x45 (69d)
  lUnknown8            As Long ' 0x49 (73d)
  lUnknown9            As Long ' 0x4D (77d)
  lUnknown10           As Long ' 0x51 (81d)
  lUnknown11           As Long ' 0x55 (85d)
  lPropertiesLength    As Long ' 0x59 (89d) Total Length of Properties
                               ' 0x5D (93d) <-- Structure Size
End Type

Private Type typeApiList
    strLibraryName As String
    strFunctionName As String
End Type

Private Type typeProcedureList
    strParent As String
    strProcedureName As String
End Type

'Globals begin
Global gProcedureList() As typeProcedureList
Global gApiList() As typeApiList
Global gVBHeader As VBHeader
'Com Stuff
Global gCOMRegInfo As tCOMRegInfo
Global gCOMRegData As tCOMRegData
Global gDesignerInfo As tDesignerInfo

Global gProjectInfo As tProjectInfo
Global gObjectTable As tObjectTable
Global gObject() As tObject
Global gObjectInfo As tObjectInfo
Global gExternalTable As ExternalTable
Global gExternalLibrary As ExternalLibrary
Global gOptionalObjectInfo As tOptionalObjectInfo
Global gEventLink As tEventLink
Global gEventPointer As tEventPointer
Global gControl() As tControl
Global gEventTable() As tEventTable
Global gCodeInfo As tCodeInfo
Global gProcedure()  As Long 'As tProcedure
Global gGuiTable() As tGuiTable
Global gObjectNameArray() As String
Global gObjectProcCountArray() As Integer
Global gObjectInfoHolder() As tObjectInfo
Private Type tObjectOffsetType
    Address As Long
    ObjectName As String
End Type
Global gObjectOffsetArray() As tObjectOffsetType

'Options
Global gSkipCom As Boolean
Global gDumpData As Boolean
Global gShowOffsets As Boolean
Global gShowColors As Boolean
Global gPcodeDecompile As Boolean

Private Type typeControlName
    strParentForm As String
    strControlName As String
    strGuid As String
    bControlImage As Byte
End Type
Global gControlNameArray() As typeControlName


'For Controls
Public Type typeStandardControlSize
    cLeft As Integer
    cTop As Integer
    cWidth As Integer
    cHeight As Integer
End Type
Public Type typeStandardControlSize2
    cLeft As Long
    cTop As Long
    cWidth As Long
    cHeight As Long
End Type
'Picture Header
Public Type typePictureHeader
    un1 As Integer
    un2 As Integer
    un3 As Integer
    un4 As Integer
End Type

Private Type ImportListtype
    strName As String
    strGuid As String
    strLib As String
End Type
Global ImportList() As ImportListtype


'Used for Memory Map
Public gVBFile As clsFile
Public gMemoryMap As clsMemoryMap

'Variables for .vbp file
Global ProjectExename As String                     ' Project exename. MaxLength: 0x104 (260d)
Global ProjectTitle As String                       ' Project title. MaxLength: 0x28 (40d)
Global HelpFile As String                           ' Helpfile. MaxLength: 0x28 (40d)
Global ProjectName As String                        ' Project name. MaxLength: 0x104 (260d)
Global ProjectDescription As String

'Determine Object Type
Private Type objectTypeListType
    value As Long
    strType As Byte '1=form 2=module 3=class 4=usercontrol 5=property page 6=user document
End Type
Global gObjectTypeList() As objectTypeListType

'Get File Information File Version Properties
Public Type FILEPROPERTIE
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
    LanguageID As String
    Comments As String
    LegalTradeMark As String
End Type
Global gFileInfo As FILEPROPERTIE
Declare Function GetFileVersionInfo Lib "Version.dll" Alias _
   "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal _
   dwhandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias _
   "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long
Declare Function VerQueryValue Lib "Version.dll" Alias _
   "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, _
   lplpBuffer As Any, puLen As Long) As Long

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Dest As Any, ByVal Source As Long, ByVal length As Long)
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
    ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Const LANG_ENGLISH = &H9

Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Public Type FIRSTCHAR_INFO
    sChar As String
    lCursor As Long
End Type

Public Const SW_SHOWNORMAL = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Code Colors
Global InstrColor As Long
Global FuncColor As Long
Global StringColor As Long
Global CommentColor As Long
        
Global gUpdateText As Boolean 'Update Syntax Coloring?
Global gIdentSpaces As Integer 'For identing code
Global gDllProject As Boolean 'Is it a dll project?
Global CancelDecompile As Boolean

Global gVB6App As Boolean
Global gVB5App As Boolean
Global gVB4App As Boolean

Public Const cQuote As String = """"

'Exports Type
Public Type ExportsType
    Ordinal As Integer
    EntryPoint As Long
    FunctionName As String
End Type
'Events Type Holders
Dim EventCommandButton() As Byte
Dim EventForm() As Byte
Dim EventPictureBox() As Byte
Dim EventImage() As Byte
Dim EventLabel() As Byte
Dim EventTextBox() As Byte
Dim EventFrame() As Byte
Dim EventCheckBox() As Byte
Dim EventRadioButton() As Byte
Dim EventComboBox() As Byte
Dim EventListBox() As Byte
Dim EventHscroll() As Byte
Dim EventVscroll() As Byte
Dim EventDriveListBox() As Byte
Dim EventDirListBox() As Byte
Dim EventFileListBox() As Byte
Dim EventData() As Byte
Dim EventMDIForm() As Byte
Dim EventUserDocument() As Byte

Dim strFormBuffer As String
Dim strErrorLog As String 'Save Errors to File
Public Const MODSIGNATURE = 1#
Public Const FRMHEADERMSB = 52479#
Public Const FRMHEADERLSB = 49#                    '(0xFFCC)3100 ID code for FRM object structure
Public Const FORMHIGHMASK = 32768#
Public Const STARTUPMODOFFSET1 = 76#               '0x4C = offset from VBsignature to StartUP vector #1
Public Const STARTUPMODOFFSET2 = 72#               '0x48 = offset from StartUp vector #1 to StartUp vector #2
Public Const STARTUPMODOFFSET3 = 98#
Public Const FORMSIZEWORDOFFSET = 88#
Public Const TEXTTABLEPTROFFSET = 12#
'tYPEINFO
Public cTypeInfo    As clsTypeLibInfo

Sub SetupEvents()
    ReDim EventCommandButton(16)
    ReDim EventForm(31)
    ReDim EventFrame(12)
    ReDim EventPictureBox(25)
    ReDim EventLabel(17)
    ReDim EventTextBox(23)
    ReDim EventCheckBox(17)
    ReDim EventRadioButton(18)
    ReDim EventComboBox(18)
    ReDim EventListBox(20)
    ReDim EventHscroll(9)
    ReDim EventVscroll(9)
    ReDim EventDriveListBox(15)
    ReDim EventDirListBox(19)
    ReDim EventFileListBox(21)
    ReDim EventImage(12)
    ReDim EventData(14)
    ReDim EventMDIForm(30)
    'Form Events
    EventForm(0) = 5 '#5 DragDrop (Source As Control, X As Single, Y As Single)
    EventForm(1) = 6 '#6 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventForm(2) = 12 '#12 LinkClose
    EventForm(3) = 13 '#13 LinkError (LinkErr As Integer)
    EventForm(4) = 14 '#14 LinkExecute (CmdStr As String, Cancel As Integer)
    EventForm(5) = 15 '#15 LinkOpen (Cancel As Integer)
    EventForm(6) = 16 '#6 = Form_Load
    EventForm(7) = 29 '#29 Resize
    EventForm(9) = 28 '#28 QueryUnload (Cancel As Integer, UnloadMode As Integer)
    EventForm(10) = 1 '#1 Activate
    EventForm(11) = 4 '#4 Deactivate
    EventForm(12) = 2 '#2 Click
    EventForm(13) = 3 '#3 DblClick
    EventForm(14) = 7 '#7 GotFocus
    EventForm(8) = 31 ' #31 Unload (Cancel As Integer)
    EventForm(15) = 9 '#9 KeyDown (KeyCode As Integer, Shift As Integer)
    EventForm(16) = 10 '#10 KeyPress (KeyAscii As Integer)
    EventForm(17) = 11 '#11 KeyUp (KeyCode As Integer, Shift As Integer)
    EventForm(18) = 17 '#17 LostFocus
    EventForm(19) = 18 '#18 MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventForm(20) = 19 '#19 MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventForm(21) = 20 '#20 MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventForm(22) = 27 '#27 Paint
    EventForm(23) = 8 '#8 Initialize
    EventForm(24) = 30 '#30 Terminate
    EventForm(25) = 23 '#23 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventForm(26) = 22 '#22 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventForm(27) = 24 '#24 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventForm(28) = 26 '#26 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventForm(29) = 25 '#25 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventForm(30) = 21 '#21 OLECompleteDrag (Effect As Long)
    
    'Command Button Events
    EventCommandButton(0) = 1 '#1 Click
    EventCommandButton(1) = 2 '#2 DragDrop (Source As Control, X As Single, Y As Single)
    EventCommandButton(2) = 3 '#3 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventCommandButton(3) = 4 '#4 GotFocus
    EventCommandButton(4) = 5 '#5 KeyDown (KeyCode As Integer, Shift As Integer)
    EventCommandButton(5) = 6 '#6 KeyPress (KeyAscii As Integer)
    EventCommandButton(6) = 7 '#7 KeyUp (KeyCode As Integer, Shift As Integer)
    EventCommandButton(7) = 8 '#8 LostFocus
    EventCommandButton(8) = 9 '#9 MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventCommandButton(9) = 10 '#10 MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventCommandButton(10) = 11 '#11 MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventCommandButton(11) = 14 '#14 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventCommandButton(12) = 12 '#12 OLECompleteDrag (Effect As Long)
    EventCommandButton(13) = 15 '#15 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventCommandButton(14) = 17 '#17 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventCommandButton(15) = 16 '#16 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventCommandButton(16) = 13 '#13 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Frame Events
    EventFrame(0) = 3 '#3 DragDrop (Source As Control, X As Single, Y As Single)
    EventFrame(1) = 4 '#4 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventFrame(2) = 5 '#5 MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventFrame(3) = 6 '#6 MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventFrame(4) = 7 '#7 MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventFrame(5) = 1 '#1 Click
    EventFrame(6) = 2 '#2 DblClick
    EventFrame(7) = 10 '#10 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventFrame(8) = 9 '#9 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventFrame(9) = 11 '#11 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventFrame(10) = 13 '#13 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventFrame(11) = 12 '#12 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventFrame(12) = 8 '#8 OLECompleteDrag (Effect As Long)
    
    'PictureBox Events
    EventPictureBox(0) = 1 '#1 Change
    EventPictureBox(1) = 2 '#2 Click
    EventPictureBox(2) = 3 '#3 DblClick
    EventPictureBox(3) = 4 '#4 DragDrop (Source As Control, X As Single, Y As Single)
    EventPictureBox(4) = 5 '#5 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventPictureBox(5) = 6 '#6 GotFocus
    EventPictureBox(6) = 7 '#7 KeyDown (KeyCode As Integer, Shift As Integer)
    EventPictureBox(7) = 8 '#8 KeyPress (KeyAscii As Integer)
    EventPictureBox(8) = 9 '#9 KeyUp (KeyCode As Integer, Shift As Integer)
    EventPictureBox(9) = 10 '#10 LinkClose
    EventPictureBox(10) = 11 '#11 LinkError (LinkErr As Integer)
    EventPictureBox(11) = 13 '#13 LinkOpen (Cancel As Integer)
    EventPictureBox(12) = 14 '#14 LostFocus
    EventPictureBox(13) = 15 '#15 MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventPictureBox(14) = 16 '#16 MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventPictureBox(15) = 17 '#17 MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventPictureBox(16) = 24 '#24 Paint
    EventPictureBox(17) = 12 '#12 LinkNotify
    EventPictureBox(18) = 25 '#25 Resize
    EventPictureBox(19) = 20 '#20 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventPictureBox(20) = 19 '#19 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventPictureBox(21) = 21 '#21 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventPictureBox(22) = 23 '#23 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventPictureBox(23) = 22 '#22 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventPictureBox(24) = 18 '#18 OLECompleteDrag (Effect As Long)
    EventPictureBox(25) = 26 '#26 Validate (Cancel As Boolean)
    
    'Label events
    EventLabel(0) = 1 '#1 Change
    EventLabel(1) = 2 '#2 Click
    EventLabel(2) = 3 '#3 DblClick
    EventLabel(3) = 4 '#4 DragDrop (Source As Control, X As Single, Y As Single)
    EventLabel(4) = 5 '#5 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventLabel(5) = 6 '#6 LinkClose
    EventLabel(6) = 7 '#7 LinkError (LinkErr As Integer)
    EventLabel(7) = 9 '#9 LinkOpen (Cancel As Integer)
    EventLabel(8) = 10 '#10 MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventLabel(9) = 11 '#11 MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventLabel(10) = 12 '#12 MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventLabel(11) = 8 '#8 LinkNotify
    EventLabel(12) = 15 '#15 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventLabel(13) = 14 '#14 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventLabel(14) = 16 '#16 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventLabel(15) = 18 '#18 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventLabel(16) = 17 '#17 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventLabel(17) = 13 '#13 OLECompleteDrag (Effect As Long)
    
    'TextBox Events
    EventTextBox(0) = 1 '#1 Change
    EventTextBox(1) = 4 '#4 DragDrop (Source As Control, X As Single, Y As Single)
    EventTextBox(2) = 5 '#5 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventTextBox(3) = 6 '#6 GotFocus
    EventTextBox(4) = 7 '#7 KeyDown (KeyCode As Integer, Shift As Integer)
    EventTextBox(5) = 8 '#8 KeyPress (KeyAscii As Integer)
    EventTextBox(6) = 9 '#9 KeyUp (KeyCode As Integer, Shift As Integer)
    EventTextBox(7) = 10 '#10 LinkClose
    EventTextBox(8) = 11 '#11 LinkError (LinkErr As Integer)
    EventTextBox(9) = 13 '#13 LinkOpen (Cancel As Integer)
    EventTextBox(10) = 14 '#14 LostFocus
    EventTextBox(11) = 12 '#12 LinkNotify
    EventTextBox(12) = 15 '#15 MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventTextBox(13) = 16 '#16 MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventTextBox(14) = 17 '#17 MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventTextBox(15) = 2 '#2 Click
    EventTextBox(16) = 3 '#3 DblClick
    EventTextBox(17) = 20 '#20 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventTextBox(18) = 19 '#19 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventTextBox(19) = 21 '#21 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventTextBox(20) = 23 '#23 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventTextBox(21) = 22 '#22 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventTextBox(22) = 18 '#18 OLECompleteDrag (Effect As Long)
    EventTextBox(23) = 24 '#24 Validate (Cancel As Boolean)

    'CheckBox Event
    EventCheckBox(0) = 1 '#1 Click
    EventCheckBox(1) = 2 '#2 DragDrop (Source As Control, X As Single, Y As Single)
    EventCheckBox(2) = 3 '#3 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventCheckBox(3) = 4 '#4 GotFocus
    EventCheckBox(4) = 5 '#5 KeyDown (KeyCode As Integer, Shift As Integer)
    EventCheckBox(5) = 6 '#6 KeyPress (KeyAscii As Integer)
    EventCheckBox(6) = 7 '#7 KeyUp (KeyCode As Integer, Shift As Integer)
    EventCheckBox(7) = 8 '#8 LostFocus
    EventCheckBox(8) = 9 '#9 MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventCheckBox(9) = 10 '#10 MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventCheckBox(10) = 11 '#11 MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventCheckBox(11) = 14 '#14 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventCheckBox(12) = 13 '#13 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventCheckBox(13) = 15 '#15 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventCheckBox(14) = 17 '#17 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventCheckBox(15) = 16 '#16 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventCheckBox(16) = 12 '#12 OLECompleteDrag (Effect As Long)
    EventCheckBox(17) = 18 '#18 Validate (Cancel As Boolean)
    
    'Option/Radio Button
    EventRadioButton(0) = 1 '#1 Click
    EventRadioButton(1) = 2 '#2 DblClick
    EventRadioButton(2) = 3 '#3 DragDrop (Source As Control, X As Single, Y As Single)
    EventRadioButton(3) = 4 '#4 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventRadioButton(4) = 5 '#5 GotFocus
    EventRadioButton(5) = 6 '#6 KeyDown (KeyCode As Integer, Shift As Integer)
    EventRadioButton(6) = 7 '#7 KeyPress (KeyAscii As Integer)
    EventRadioButton(7) = 8 '#8 KeyUp (KeyCode As Integer, Shift As Integer)
    EventRadioButton(8) = 9 '#9 LostFocus
    EventRadioButton(9) = 10 '#10 MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventRadioButton(10) = 11 '#11 MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventRadioButton(11) = 12 '#12 MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventRadioButton(12) = 15 '#15 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventRadioButton(13) = 14 '#14 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventRadioButton(14) = 16 '#16 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventRadioButton(15) = 18 '#18 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventRadioButton(16) = 17 '#17 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventRadioButton(17) = 13 '#13 OLECompleteDrag (Effect As Long)
    EventRadioButton(18) = 19 '#19 Validate (Cancel As Boolean)
    
    'ComboBox
    EventComboBox(0) = 1 '#1 Change
    EventComboBox(1) = 2 '#2 Click
    EventComboBox(2) = 3 '#3 DblClick
    EventComboBox(3) = 4 '#4 DragDrop (Source As Control, X As Single, Y As Single)
    EventComboBox(4) = 5 '#5 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventComboBox(5) = 6 '#6 DropDown
    EventComboBox(6) = 7 '#7 GotFocus
    EventComboBox(7) = 8 '#8 KeyDown (KeyCode As Integer, Shift As Integer)
    EventComboBox(8) = 9 '#9 KeyPress (KeyAscii As Integer)
    EventComboBox(9) = 10 '#10 KeyUp (KeyCode As Integer, Shift As Integer)
    EventComboBox(10) = 11 '#11 LostFocus
    EventComboBox(11) = 14 '#14 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventComboBox(12) = 13 '#13 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventComboBox(13) = 15 '#15 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventComboBox(14) = 17 '#17 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventComboBox(15) = 16 '#16 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventComboBox(16) = 12 '#12 OLECompleteDrag (Effect As Long)
    EventComboBox(17) = 18 '#18 Scroll
    EventComboBox(18) = 19 '#19 Validate (Cancel As Boolean)
    
    'ListBox
    EventListBox(0) = 1 '#1 Click
    EventListBox(1) = 2 '#2 DblClick
    EventListBox(2) = 3 '#3 DragDrop (Source As Control, X As Single, Y As Single)
    EventListBox(3) = 4 '#4 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventListBox(4) = 5 '#5 GotFocus
    EventListBox(5) = 7 '#7 KeyDown (KeyCode As Integer, Shift As Integer)
    EventListBox(6) = 8 '#8 KeyPress (KeyAscii As Integer)
    EventListBox(7) = 9 '#9 KeyUp (KeyCode As Integer, Shift As Integer)
    EventListBox(8) = 10 '#10 LostFocus
    EventListBox(9) = 11 '#11 MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventListBox(10) = 12 '#12 MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventListBox(11) = 13 '#13 MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventListBox(12) = 16 '#16 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventListBox(13) = 15 '#15 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventListBox(14) = 17 '#17 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventListBox(15) = 19 '#19 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventListBox(16) = 18 '#18 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventListBox(17) = 14 '#14 OLECompleteDrag (Effect As Long)
    EventListBox(18) = 20 '#20 Scroll
    EventListBox(19) = 6 '#6 ItemCheck (Item As Integer)
    EventListBox(20) = 21 '#21 Validate (Cancel As Boolean)
    
    'Hscroll
    EventHscroll(0) = 1 '#1 Change
    EventHscroll(1) = 2 '#2 DragDrop (Source As Control, X As Single, Y As Single)
    EventHscroll(2) = 3 '#3 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventHscroll(3) = 4 '#4 GotFocus
    EventHscroll(4) = 5 '#5 KeyDown (KeyCode As Integer, Shift As Integer)
    EventHscroll(5) = 6 '#6 KeyPress (KeyAscii As Integer)
    EventHscroll(6) = 7 '#7 KeyUp (KeyCode As Integer, Shift As Integer)
    EventHscroll(7) = 8 '#8 LostFocus
    EventHscroll(8) = 9 '#9 Scroll
    EventHscroll(9) = 10 '#10 Validate (Cancel As Boolean)
    
    'Vscroll
    EventVscroll(0) = 1 '#1 Change
    EventVscroll(1) = 2 '#2 DragDrop (Source As Control, X As Single, Y As Single)
    EventVscroll(2) = 3 '#3 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventVscroll(3) = 4 '#4 GotFocus
    EventVscroll(4) = 5 '#5 KeyDown (KeyCode As Integer, Shift As Integer)
    EventVscroll(5) = 6 '#6 KeyPress (KeyAscii As Integer)
    EventVscroll(6) = 7 '#7 KeyUp (KeyCode As Integer, Shift As Integer)
    EventVscroll(7) = 8 '#8 LostFocus
    EventVscroll(8) = 9 ' #9 Scroll
    EventVscroll(9) = 10 '#10 Validate (Cancel As Boolean)
    
    'Drive List Box
    EventDriveListBox(0) = 1 '#1 Change
    EventDriveListBox(1) = 2 '#2 DragDrop (Source As Control, X As Single, Y As Single)
    EventDriveListBox(2) = 3 '#3 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventDriveListBox(3) = 4 '#4 GotFocus
    EventDriveListBox(4) = 5 '#5 KeyDown (KeyCode As Integer, Shift As Integer)
    EventDriveListBox(5) = 6 '#6 KeyPress (KeyAscii As Integer)
    EventDriveListBox(6) = 7 '#7 KeyUp (KeyCode As Integer, Shift As Integer)
    EventDriveListBox(7) = 8 '#8 LostFocus
    EventDriveListBox(8) = 11 '#11 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventDriveListBox(9) = 10  '#10 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventDriveListBox(10) = 12 '#12 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventDriveListBox(11) = 14 '#14 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventDriveListBox(12) = 13 '#13 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventDriveListBox(13) = 9 '#9 OLECompleteDrag (Effect As Long)
    EventDriveListBox(14) = 15 '#15 Scroll
    EventDriveListBox(15) = 16  '#16 Validate (Cancel As Boolean)
    
    'DirListBox
    EventDirListBox(0) = 1 '#1 Change
    EventDirListBox(1) = 2 '#2 Click
    EventDirListBox(2) = 3 '#3 DragDrop (Source As Control, X As Single, Y As Single)
    EventDirListBox(3) = 4 '#4 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventDirListBox(4) = 5 '#5 GotFocus
    EventDirListBox(5) = 6 '#6 KeyDown (KeyCode As Integer, Shift As Integer)
    EventDirListBox(6) = 7 '#7 KeyPress (KeyAscii As Integer)
    EventDirListBox(7) = 8  '#8 KeyUp (KeyCode As Integer, Shift As Integer)
    EventDirListBox(8) = 9 '#9 LostFocus
    EventDirListBox(9) = 10  '#10 MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventDirListBox(10) = 11 '#11 MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventDirListBox(11) = 12 '#12 MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventDirListBox(12) = 15 '#15 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventDirListBox(13) = 14 '#14 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventDirListBox(14) = 16 '#16 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventDirListBox(15) = 18  '#18 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventDirListBox(16) = 17 '#17 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventDirListBox(17) = 13 '#13 OLECompleteDrag (Effect As Long)
    EventDirListBox(18) = 19 '#19 Scroll
    EventDirListBox(19) = 20 '#20 Validate (Cancel As Boolean)
    
    'File List Box
    EventFileListBox(0) = 1 '#1 Click
    EventFileListBox(1) = 2 '#2 DblClick
    EventFileListBox(2) = 3 '#3 DragDrop (Source As Control, X As Single, Y As Single)
    EventFileListBox(3) = 4 '#4 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventFileListBox(4) = 5 '#5 GotFocus
    EventFileListBox(5) = 6 '#6 KeyDown (KeyCode As Integer, Shift As Integer)
    EventFileListBox(6) = 7 '#7 KeyPress (KeyAscii As Integer)
    EventFileListBox(7) = 8 '#8 KeyUp (KeyCode As Integer, Shift As Integer)
    EventFileListBox(8) = 9 '#9 LostFocus
    EventFileListBox(9) = 10 '#10 MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventFileListBox(10) = 11 '#11 MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventFileListBox(11) = 12 '#12 MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventFileListBox(12) = 19 '#19 PathChange
    EventFileListBox(13) = 20  '#20 PatternChange
    EventFileListBox(14) = 15 '#15 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventFileListBox(15) = 14 '#14 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventFileListBox(16) = 16  '#16 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventFileListBox(17) = 18 '#18 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventFileListBox(18) = 17 ' #17 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventFileListBox(19) = 13 '#13 OLECompleteDrag (Effect As Long)
    EventFileListBox(20) = 21 '#21 Scroll
    EventFileListBox(21) = 22 '#22 Validate (Cancel As Boolean)
    
    'Image
    EventImage(0) = 1 '#1 Click
    EventImage(1) = 2 '#2 DblClick
    EventImage(2) = 3 '#3 DragDrop (Source As Control, X As Single, Y As Single)
    EventImage(3) = 4 '#4 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventImage(4) = 5 '#5 MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventImage(5) = 6 '#6 MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventImage(6) = 7 '#7 MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventImage(7) = 10 '#10 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventImage(8) = 9 '#9 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventImage(9) = 11 '#11 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventImage(10) = 13 '#13 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventImage(11) = 12 '#12 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventImage(12) = 8 '#8 OLECompleteDrag (Effect As Long)
    
    'Data
    EventData(0) = 3 '#3 Error (DataErr As Integer, Response As Integer)
    EventData(1) = 13 '#13 Reposition
    EventData(2) = 15 '#15 Validate (Action As Integer, Save As Integer)
    EventData(3) = 1 '#1 DragDrop (Source As Control, X As Single, Y As Single)
    EventData(4) = 2 '#2 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventData(5) = 4 '#4 MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventData(6) = 5 '#5 MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventData(7) = 6 '#6 MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventData(8) = 14 '#14 Resize
    EventData(9) = 9 '#9 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventData(10) = 8 '#8 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventData(11) = 10 '#10 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventData(12) = 12 '#12 OLEStartDrag (Data As VBRUN.DataObject, AllowedEffects As Long)
    EventData(13) = 11 '#11 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventData(14) = 7  '#7 OLECompleteDrag (Effect As Long)
    
    'MDI Form
    EventMDIForm(0) = 5 '#5 DragDrop (Source As Control, X As Single, Y As Single)
    EventMDIForm(1) = 6 '#6 DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    EventMDIForm(2) = 8 '#8 LinkClose
    EventMDIForm(3) = 9 '#9 LinkError (LinkErr As Integer)
    EventMDIForm(4) = 10 '#10 LinkExecute (CmdStr As String, Cancel As Integer)
    EventMDIForm(5) = 11 '#11 LinkOpen (Cancel As Integer)
    EventMDIForm(6) = 12 '#12 Load
    EventMDIForm(7) = 23 '#23 Resize
    EventMDIForm(8) = 25 '#25 Unload (Cancel As Integer)
    EventMDIForm(9) = 22 '#22 QueryUnload (Cancel As Integer, UnloadMode As Integer)
    EventMDIForm(10) = 1 '#1 Activate
    EventMDIForm(11) = 4 '#4 Deactivate
    EventMDIForm(12) = 2 '#2 Click
    EventMDIForm(13) = 3 '#3 DblClick
    EventMDIForm(19) = 13 '#13 MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventMDIForm(20) = 14 '#14 MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventMDIForm(21) = 15 '#15 MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventMDIForm(23) = 7 '#7 Initialize
    EventMDIForm(24) = 24 '#24 Terminate
    EventMDIForm(25) = 18 '#18 OLEDragOver (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    EventMDIForm(26) = 17 '#17 OLEDragDrop (Data As VBRUN.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventMDIForm(27) = 19 '#19 OLEGiveFeedback (Effect As Long, DefaultCursors As Boolean)
    EventMDIForm(28) = 21
    EventMDIForm(29) = 20 '#20 OLESetData (Data As VBRUN.DataObject, DataFormat As Integer)
    EventMDIForm(30) = 16 '#16 OLECompleteDrag (Effect As Long)
    
    
    
End Sub
Public Function GetEventNumber(ByVal strGuid As String, index As Integer) As Integer
   'Call SetupEvents
'MsgBox strGuid
    'Form Events
    If strGuid = "{33AD4F38-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventForm(index))
        Exit Function
    End If
    'Command Button Events
    If strGuid = "{33AD4EF0-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventCommandButton(index))
        Exit Function
    End If
    'Command Button Array Events
   ' If strGuid = "{33AD4EF1-6699-11CF-B70C-00AA0060D393}" Then
        'GetEventNumber = Int(EventCommandButton(Index))
       ' Exit Function
    'End If

    'Timer
    If strGuid = "{33AD4F28-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(1) 'Only one Event Timer
        Exit Function
    End If
    'Menu
    If strGuid = "{33AD4F68-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(1) 'Only one Event Menu
        Exit Function
    End If
    'Frame
    If strGuid = "{33AD4EE8-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventFrame(index))
        Exit Function
    End If
    'Picture Box
    If strGuid = "{33AD4ED0-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventPictureBox(index))
        Exit Function
    End If
    'Label
    If strGuid = "{33AD4ED8-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventLabel(index))
        Exit Function
    End If
    'TextBox
    If strGuid = "{33AD4EE0-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventTextBox(index))
        Exit Function
    End If
    'Check box
    If strGuid = "{33AD4EF8-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventCheckBox(index))
        Exit Function
    End If
    'Radio Button/Option Button
    If strGuid = "{33AD4F00-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventRadioButton(index))
        Exit Function
    End If
    'Combo Box
    If strGuid = "{33AD4F08-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventComboBox(index))
        Exit Function
    End If
    'Listbox
    If strGuid = "{33AD4F10-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventListBox(index))
        Exit Function
    End If
    'Hscroll
    If strGuid = "{33AD4F18-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventHscroll(index))
        Exit Function
    End If
    'Vscroll
    If strGuid = "{33AD4F20-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventVscroll(index))
        Exit Function
    End If
    'Drive List Box
    If strGuid = "{33AD4F50-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventDriveListBox(index))
        Exit Function
    End If
    'Dir List Box
    If strGuid = "{33AD4F58-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventDirListBox(index))
        Exit Function
    End If
    'File List box
    If strGuid = "{33AD4F60-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventFileListBox(index))
        Exit Function
    End If
    'Image
    If strGuid = "{33AD4F90-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventImage(index))
        Exit Function
    End If
    'Data
    If strGuid = "{33AD4FF8-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventData(index))
        Exit Function
    End If
    'MDI Form
    If strGuid = "{33AD4F70-6699-11CF-B70C-00AA0060D393}" Then
        GetEventNumber = Int(EventMDIForm(index))
        Exit Function
    End If
    
    GetEventNumber = -1
End Function
Sub PrintReadMe()
    '*****************************
    'Prints the ReadMe of the program
    '*****************************
    On Error Resume Next
    Kill (App.Path & "\readme.txt")
    Dim F As Long
    F = FreeFile
    Open App.Path & "\ReadMe.txt" For Output As #F
        Print #F, "-------------------------------"
        Print #F, "Semi VB Decompiler - VisualBasicZone.com"
        Print #F, "Version: " & Version
        Print #F, "Build: " & App.major & "." & App.minor & "." & App.revision
        Print #F, "Website: http://www.visualbasiczone.com/products/semivbdecompiler"
        Print #F, "-------------------------------"
        Print #F, "Contents"
        Print #F, "1. What's New?"
        Print #F, "2. Features"
        Print #F, "3. Questions?"
        Print #F, "4. Bugs"
        Print #F, "5. Contact"
        Print #F, "6. Credits"
        Print #F, ""
        Print #F, "1. What's New?"
        Print #F, ""
        Print #F, "   Version 0.09"
        Print #F, "   Added a new tool. Api Add allows you to add Api's to the Semi VB Decompiler Api Database."
        Print #F, ""
        Print #F, "   Version 0.08 Build 1.0.64"
        Print #F, "   Updated Native Procedure Decompile dissembles faster and added some native dissemble options to the options screen. Also updated decompile from offset, it now verify the files has a VB5! signature."
        Print #F, "   For .Net applications added the view console under the Tools menu.  Added data directories to the PE Optional Header list."
        Print #F, ""
        Print #F, "   Version 0.07 Build 1.0.63"
        Print #F, "   Minor GUI updates and fixes. Fixed VBP external component bug."
        Print #F, ""
        Print #F, "   Version 0.07 Build 1.0.61"
        Print #F, "   Redid the P-Code property name finder procedure for the standard toolbox controls."
        Print #F, "   Now the VTables are now pulled from VB6.olb the typelib file instead of having them hardcoded."
        Print #F, ""
        Print #F, "   Version 0.07 Build 1.0.60"
        Print #F, "   Switched back to my old sytle control property editing. Added saving the old value, so when you switch controls, the changed value is shown."
        Print #F, ""
        Print #F, "   Version 0.07 Build 1.0.59"
        Print #F, "   Added over 100 more P-Code properties to the database"
        Print #F, ""
        Print #F, "   Version 0.07 Build 1.0.58"
        Print #F, "   Fixed a problem loading with some VB Dll's"
        Print #F, "   Included support for Windows XP Styles"
        Print #F, ""
        Print #F, "   Version 0.07 Build 1.0.57"
        Print #F, "   Fixed some VB Detection Problems."
        Print #F, "   Added option to show offset off P-Code String in the P-Code String List"
        Print #F, ""
        Print #F, "   Version 0.07 Build 1.0.53"
        Print #F, "   New icons to indicate the control type in the treeview."
        Print #F, "   Using a property sheet control for property editing and viewing now."
        Print #F, "   Memory Map Generated way faster"
        Print #F, "   Added View Report menu under the Tools Menu."
        Print #F, ""
        Print #F, "   Version 0.06C Build 1.0.52"
        Print #F, "   Improved P-Code decompiling."
        Print #F, "   Api calls are shown as library name . Function Name"
        Print #F, "   For VCallHresult properties are recovered and the object type."
        Print #F, "   Fixed Extra opcodes after end of procedure for P-Code."
        Print #F, ""
        Print #F, "   Version 0.06B Build 1.0.50"
        Print #F, "   Updated VB 1/2/3 Binary Form To Text supports all default controls properties."
        Print #F, "   Redid the menus for all the applications.   The menu's now include Bitmaps for some items, and are now styled."
        Print #F, "   Improved handleing of non vb files, in the treeview."
        Print #F, "   Will detect if the file is protected by the UPX packer."
        Print #F, "   Added Startup Form Patcher, you can choose which form appears first!"
        Print #F, "   Other improvements here and there."
        Print #F, ""
        Print #F, "   Version 0.06A Build 1.0.46"
        Print #F, "   Added support for User Documents."
        Print #F, "   Better Control Processing for unknown opcodes, errors recorded to a file."
        Print #F, "   Added P-Code String List."
        Print #F, "   Added Type Library Explorer program."
        Print #F, ""
        Print #F, "   Verion 0.06A Build 1.0.44"
        Print #F, "   Added a new tool to convert VB 1/2/3 binary forms to text."
        Print #F, "   VBP File output now includes the thread flags."
        Print #F, "   P-Code output is now semi-colored and is in bold."
        Print #F, "   P-Code To VB Code is now colored as well."
        Print #F, "   Redid part of the Control Property editing functions works better."
        Print #F, "   Took out the ComFix.txt file and just included the information in the decompiler."
        Print #F, ""
        Print #F, "   Version 0.06 Build 1.0.21"
        Print #F, "   Detection for all versions of vb from 1 to vb.net"
        'Print #f, "   Added VB2 and VB3 decompiling."
        Print #F, "   Ne Format exe's now shown in FileReport.txt"
        'Print #f, "   The VB2 and VB3 decompiler can be used in your projects via the ocx control."
        Print #F, "   Partial VB4 Support Added."
        Print #F, "   Fixed Backcolor property on labels VB5/6"
        Print #F, "   Now correctly decompiles VB5 dll files."
        Print #F, "   Fixed too many Ends for Menu's."
        Print #F, "   Faster form processing."
        Print #F, "   .Net Structures now shown under .Net Structures."
        Print #F, "   Added more .Net processing now shows Strings, Blobs, Guids, and User Strings"
        Print #F, ""
        Print #F, "   Version 0.05A Build 1.0.20"
        Print #F, "   Faster syntax coloring."
        Print #F, "   More detailed filereports containing pe information."
        Print #F, ""
        Print #F, "   Version 0.05 Build 1.0.19"
        Print #F, "   Added VB.net Detection and shows the CLR header."
        Print #F, "   Better handling of PE imports and VB version detection for other versions."
        Print #F, ""
        Print #F, "   Version 0.05"
        Print #F, "   Correct events identified for all common controls."
        Print #F, "   Control Arrays now handled correctly."
        Print #F, "   Minor fixes here and there."
        Print #F, ""
        Print #F, "   Last Version 0.04C Build 1.0.15"
        Print #F, "   Added Advanced Decompile which you can decompile a vb project"
        Print #F, "   by offset, could be used against packed exes, packed with upx and other compressors."
        Print #F, "   Added Native Procedure Decompile a work in progress."
        Print #F, "   Added Object= includes to forms and project file."
        Print #F, "   Added Export Procedure List, Add Address to P-Code Procedure Decompile."
        Print #F, "   Now shows PE information even if its not a valid vb exe."
        Print #F, ""
        Print #F, "   Version 0.04C"
        Print #F, "   Added update checker."
        Print #F, "   Added vbcode to P-Code procedure decompile."
        Print #F, "   Added memory offset to file offset."
        Print #F, "   Improved Form and control decoding."
        Print #F, "   Fixed Ocx loading bug."
        Print #F, "   Fixed Empty .frx generation. Fixed null subs in form/class generation."
        Print #F, "   Thinking about working on VB4 and VB3 support for the next version."
        Print #F, ""
        Print #F, "   Version 0.04"
        Print #F, "    Improved P-Code Decompiling a lot."
        Print #F, "    Better ObjectType detection."
        Print #F, "    Added control property editing."
        Print #F, "    Added VB5 Support!"
        Print #F, ""
        Print #F, "   Version 0.03"
        Print #F, "     P-Code decoding started and image extraction."
        Print #F, "     Numerous bug fixes."
        Print #F, "     Event detection added."
        Print #F, "     Dll and OCX Support added."
        Print #F, "     External Components added to vbp file."
        Print #F, "     Begun work on a basic antidecompiler."
        Print #F, "     Form property editor, complete with a patch report generator."
        Print #F, "     Procedure names are recovered."
        Print #F, "     Api's used by the program are recovered."
        Print #F, "     Msvbvm60.dll imports are listed in the treeview."
        Print #F, "     Syntax coloring for Forms."
        Print #F, "     Fixed scrolling bug."
        Print #F, ""
        Print #F, "   Version 0.02"
        Print #F, "     Rebuilds the forms"
        Print #F, "     Gets most controls and their properties."
        Print #F, ""
        Print #F, "   Initial Release version 0.01"
        Print #F, ""
        Print #F, "2. Features"
        Print #F, "     Decompiling the P-Code/native vb 4/5/6 exe's, dll's, and ocx's"
        Print #F, "     Form Generation"
        Print #F, "     Resource extraction wmf, ico, cur, gif, bmp, jpg, dib"
        Print #F, "     Control/Form Editor"
        Print #F, "     Startup Form Patcher for VB 5/6"
        Print #F, "     Address to File Offset converter."
        Print #F, "     P-Code Event/Procedure Decompile"
        Print #F, "     Native Event Disassembly"
        Print #F, "     Shows offsets for controls and allows you to edit the control properties."
        Print #F, "     Decompile a file from an offset useful against packed exe's using compression such as upx."
        Print #F, "     Multilanguage support including Dutch, French, German, Italian, spanish and more"
        Print #F, "     Memory Map of the exe file, so you can see what's going on."
        Print #F, "     Advanced decompiling using COM instead of hard coding property opcodes."
        Print #F, ""
        Print #F, "3. Questions?"
        Print #F, "   Q. What about Native Code Decompiling?"
        Print #F, "   A. It is in the works. Right now I have offsets for all the events and can do a disassembly of each event but I need to work on an assembly to VB engine still."
        Print #F, "   Q. What the heck are the P-Code Tokens?"
        Print #F, "   A. P-Code tokens is the last step before turning the P-Code into readable VB Code."
        Print #F, "      All you have to do now is link the imports of the exe with the functions in P-Code."
        Print #F, "   Q. Why does it not show all the controls on my forms?"
        Print #F, "   A. If it is not a common control found in the toolbox then we can not get extra information it, in the future we maybe able to process these controls."
        Print #F, "      Another reason can be because it is a property that is not detected by COM using vb6.olb."
        Print #F, "   Q. Why doesn't it get my procedure names for Modules?"
        Print #F, "   A. Visual Basic only saves procedures names for Form's and Classes.  And it only saves them for forms if they are public."
        Print #F, "   Q. How does this decompiler work?"
        Print #F, "   A. First it gets all the main vb structures from the exe."
        Print #F, "      Next it gets all the controls properties via COM using vb6.olb"
        Print #F, "   Q. What files does this decompiler require?"
        Print #F, "   A. It requires the following files:"
        Print #F, "      TLBINF32.dll"
        Print #F, "      comdlg32.OCX"
        Print #F, "      RICHTX32.OCX"
        Print #F, "      MSCOMCTL.OCX"
        Print #F, "      TABCTL32.OCX"
        Print #F, "      MSFLXGRD.OCX"
        Print #F, "      MSINET32.OCX"
        Print #F, "      Msvbvm60.dll"
        Print #F, "      SSubTmr6.dll"
        Print #F, "      WinSubHook2.tlb"
        Print #F, "      pePropertySheet.ocx"
        Print #F, "      cPopMenu6.ocx"
        Print #F, "      And VB6.olb version 6.0.9"
        Print #F, "      All of the above files need to be registered(the installer should auto register the files.)"
        Print #F, "      If you are examining a .Net file then you need to have the .Net framework installed."
        Print #F, "   Q. Where can I learn more about Visual Basic 5/6 Decompiling?"
        Print #F, "   A. Head over to http://www.vb-decompiler.com  tons of information on vb decompiling."
        Print #F, ""
        Print #F, "4. Bugs"
        Print #F, "     Some properties aren't handled yet such as dataformat"
        Print #F, "     P-Code decoding may hang use the disable P-Code option under options."
        Print #F, "     If you would wish to report a bug email me at"
        Print #F, "     support@visualbasiczone.com"
        Print #F, "     Please include as much information as possible so we can try to fix it and even better send us the file if possible."
        Print #F, ""
        Print #F, "5. Contact/Support"
        Print #F, "     Email=support@visualbasiczone.com"
        Print #F, "     Semi VB Decompiler Website:"
        Print #F, "     http://www.visualbasiczone.com/products/semivbdecompiler/"
        Print #F, ""
        Print #F, "6. Credits"
        Print #F, "     I would like to thank the following people for helping me with this project."
        Print #F, "     Sarge, Mr. Unleaded, Moogman, _aLfa_, Alex Ionescu, Warning and many others."
        
    Close #F
    
End Sub
Public Function sHexStringFromString(ByVal inp As String, Optional Spacing As Boolean = True) As String
Dim hc As String
Dim hs As String
Dim c As Long
While Len(inp)
    
    hc = Hex$(Asc(Mid$(inp, 1, 1)))
    inp = Mid$(inp, 2)
    If Len(hc) = 1 Then hc = "0" & hc
    hs = hs & hc
    c = c + 1
    If Spacing Then
        If c Mod 4 = 0 Then
            hs = hs & "  "
        ElseIf c Mod 2 = 0 Then
            hs = hs & " "
        End If
        
    End If
Wend
sHexStringFromString = hs
End Function
Public Function PadHex(ByVal sHex As String, Optional Pad As Integer = 8) As String
'*****************************
'Purpose: To add extra zero's to a hexadecimal string
'*****************************
    If Len(sHex) > Pad Then
        PadHex = sHex
    Else
        PadHex = String$(Pad - Len(sHex), 48) & sHex
    End If
End Function


Public Function GetUntilNull(FileNum As Variant) As String
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
Public Function GetUnicodeString(FileNum As Variant, length As Integer) As String
    '*****************************
    'Purpose to get a unicode string
    '*****************************
    Dim aList() As Byte

    ReDim aList((length * 2))
    Get FileNum, , aList

    Dim i As Integer
    Dim Final As String
    For i = 1 To UBound(aList) - 1
        If aList(i) <> 0 Then
            Final = Final & Chr$(aList(i))
        End If
    Next i
    
    GetUnicodeString = Final
End Function

Public Function GetGuidString(FileNum As Variant, length As Integer) As String
    '*****************************
    'Purpose to get a guid unicode string
    '*****************************
    Dim aList() As Byte

    ReDim aList((length * 2))
    Get FileNum, , aList

    Dim i As Integer
    Dim Final As String
    For i = 0 To UBound(aList) - 1 Step 2
        'If aList(i) <> 0 Then
            Final = Final & Chr$(aList(i))
            'MsgBox Final
        'End If
    Next i
    
    GetGuidString = Final
End Function
Public Function FileInfo(Optional ByVal PathWithFilename As String) As FILEPROPERTIE
'*****************************
'Purpose: To return file-properties of given file  (EXE , DLL , OCX)
'*****************************
 
Static BACKUP As FILEPROPERTIE   ' backup info for next call without filename
If Len(PathWithFilename) = 0 Then
    FileInfo = BACKUP
    Exit Function
End If

Dim lngBufferlen As Long
Dim lngDummy As Long
Dim lngRc As Long
Dim lngVerPointer As Long
Dim lngHexNumber As Long
Dim bytBuffer() As Byte
Dim bytBuff(255) As Byte
Dim strBuffer As String
Dim strLangCharset As String
Dim strVersionInfo(9) As String
Dim strTemp As String
Dim intTemp As Integer
       
' size
lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
If lngBufferlen > 0 Then
   ReDim bytBuffer(lngBufferlen)
   lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, bytBuffer(0))
   If lngRc <> 0 Then
      lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", _
               lngVerPointer, lngBufferlen)
      If lngRc <> 0 Then
         'lngVerPointer is a pointer to four 4 bytes of Hex number,
         'first two bytes are language id, and last two bytes are code
         'page. However, strLangCharset needs a  string of
         '4 hex digits, the first two characters correspond to the
         'language id and last two the last two character correspond
         'to the code page id.
         MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
         lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + _
                bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
         strLangCharset = Hex$(lngHexNumber)
         'now we change the order of the language id and code page
         'and convert it into a string representation.
         'For example, it may look like 040904E4
         'Or to pull it all apart:
         '04------        = SUBLANG_ENGLISH_USA
         '--09----        = LANG_ENGLISH
         ' ----04E4 = 1252 = Codepage for Windows:Multilingual
         'Do While Len(strLangCharset) < 8
         '    strLangCharset = "0" & strLangCharset
         'Loop
         If Mid(strLangCharset, 2, 2) = LANG_ENGLISH Then
         Dim strLangCharset2 As String
         strLangCharset2 = "English (US)"

         
         End If

         Do While Len(strLangCharset) < 8
             strLangCharset = "0" & strLangCharset
         Loop
         
         ' assign propertienames
         strVersionInfo(0) = "CompanyName"
         strVersionInfo(1) = "FileDescription"
         strVersionInfo(2) = "FileVersion"
         strVersionInfo(3) = "InternalName"
         strVersionInfo(4) = "LegalCopyright"
         strVersionInfo(5) = "OriginalFileName"
         strVersionInfo(6) = "ProductName"
         strVersionInfo(7) = "ProductVersion"
         strVersionInfo(8) = "Comments"
         strVersionInfo(9) = "LegalTrademarks"
         ' loop and get fileproperties
         For intTemp = 0 To 9
            strBuffer = String$(255, 0)
            strTemp = "\StringFileInfo\" & strLangCharset _
               & "\" & strVersionInfo(intTemp)
            lngRc = VerQueryValue(bytBuffer(0), strTemp, _
                  lngVerPointer, lngBufferlen)
            If lngRc <> 0 Then
               ' get and format data
               lstrcpy strBuffer, lngVerPointer
               strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
               strVersionInfo(intTemp) = strBuffer
             Else
               ' property not found
               strVersionInfo(intTemp) = ""
            End If
         Next intTemp
      End If
   End If
End If
' assign array to user-defined-type

    FileInfo.CompanyName = strVersionInfo(0)
    FileInfo.FileDescription = strVersionInfo(1)
    FileInfo.FileVersion = strVersionInfo(2)
    FileInfo.InternalName = strVersionInfo(3)
    FileInfo.LegalCopyright = strVersionInfo(4)
    FileInfo.OrigionalFileName = strVersionInfo(5)
    FileInfo.ProductName = strVersionInfo(6)
    FileInfo.ProductVersion = strVersionInfo(7)
    FileInfo.Comments = strVersionInfo(8)
    FileInfo.LegalTradeMark = strVersionInfo(9)
    FileInfo.LanguageID = strLangCharset2
    BACKUP = FileInfo
End Function
'*****************************
'The following functions are used for COM
'*****************************
Public Function GetBoolean(ByVal FileNum As Variant) As Boolean
'*****************************
'Purpose: Get a boolean value from a file offset
'*****************************
        Dim k As Boolean
        Get FileNum, , k
        GetBoolean = k
End Function
Public Function GetByte2(ByVal FileNum As Variant) As Byte
'*****************************
'Purpose: Get a byte value from a file offset
'*****************************
        Dim k As Byte
        Get FileNum, , k
        GetByte2 = k
End Function
Public Function GetInteger(ByVal FileNum As Variant) As Integer
'*****************************
'Purpose: Get an integer value from a file offset
'*****************************
        Dim k As Integer
        Get FileNum, , k
        
        GetInteger = k
End Function
Public Function GetLong(ByVal FileNum As Variant) As Long
'*****************************
'Purpose: Get a long value from a file offset
'*****************************
        Dim k As Long
        Get FileNum, , k
        GetLong = k
End Function
Public Function GetSingle(ByVal FileNum As Variant) As Single
'*****************************
'Purpose: Get a single value from a file offset
'*****************************
On Error GoTo badsingle:
        Dim k As Single
        Get FileNum, , k
        GetSingle = k
Exit Function
badsingle:
    GetSingle = 0
Exit Function
End Function


Public Function GetAllString(ByVal FileNum As Variant) As String
'*****************************
'Purpose: Get any kind of string Unicode or Ascii
'*****************************
    Dim length As Integer
    Get FileNum, , length
    
    Dim strText As String
    strText = GetUntilNull(FileNum)
    'MsgBox strText
    If Len(strText) < length Then
    'get unicode string
   ' MsgBox "unicode"
        If length < 100 Then
            Seek FileNum, Loc(FileNum) - 2
            strText = GetUnicodeString(FileNum, length)
            Seek FileNum, Loc(FileNum) + 1
        End If
    End If
    GetAllString = strText
End Function

Public Sub AddText(ByVal strText As String)
'*****************************
'Purpose:Adds text to the current form's textbox. And idents it.
'*****************************
    If gIdentSpaces < 0 Then gIdentSpaces = 0
    strFormBuffer = strFormBuffer & Space$(gIdentSpaces * 5) & strText & vbCrLf

End Sub
Sub LoadNewFormHolder(ByVal FormName As String)
'*****************************
'Purpose:To load a new textbox to hold each form's information
'*****************************
    Dim i As Integer

    i = frmMain.txtFinal.UBound + 1
    Load frmMain.txtFinal(i)
    With frmMain.txtFinal(i)
        .tag = FormName
        
    End With
    frmMain.txtFinal(i - 1).Text = strFormBuffer
    strFormBuffer = Space$(50000)
End Sub
Public Sub DoFinalFormBuffer()
    Dim i As Integer
    i = frmMain.txtFinal.UBound
    frmMain.txtFinal(i).Text = strFormBuffer
    strFormBuffer = ""
End Sub

Function ReturnGuid(FileNum As Variant) As String
'*****************************
'Gets a guid from a file, then corrects it into a real guid
'*****************************
Dim bArray(15) As Byte
Dim strArray(15) As String
    Get FileNum, , bArray
    Dim i As Integer
    For i = 0 To 15
        If i = 0 Then
        strArray(0) = Hex$(bArray(0) - 2)
        Else
            strArray(i) = Hex$(bArray(i))
        End If
        If Len(strArray(i)) = 1 Then
            strArray(i) = ("0" & strArray(i))
        End If
        
    Next
    
   
    Dim strFinal As String
   ' strFinal = "{" & Hex(bArray(3)) & Hex(bArray(2)) & Hex(bArray(1)) & Hex(bArray(0) - 2)
   ' strFinal = strFinal & "-" & Hex(bArray(5)) & Hex(bArray(4))
   ' strFinal = strFinal & "-" & Hex(bArray(7)) & Hex(bArray(6))
   ' strFinal = strFinal & "-" & Hex(bArray(8)) & Hex(bArray(9))
   ' strFinal = strFinal & "-" & Hex(bArray(10)) & Hex(bArray(11)) & Hex(bArray(12)) & Hex(bArray(13)) & Hex(bArray(14)) & Hex(bArray(15)) & "}"
   strFinal = "{" & strArray(3) & strArray(2) & strArray(1) & strArray(0)
   strFinal = strFinal & "-" & strArray(5) & strArray(4)
   strFinal = strFinal & "-" & strArray(7) & strArray(6)
   strFinal = strFinal & "-" & strArray(8) & strArray(9)
   strFinal = strFinal & "-" & strArray(10) & strArray(11) & strArray(12) & strArray(13) & strArray(14) & strArray(15) & "}"
    
    ReturnGuid = strFinal
End Function


Sub WriteApiList()
'*****************************
'Purpose: To write the Api's
'*****************************
On Error GoTo errHandle:
    Dim i As Integer
    frmMain.txtCode.Text = vbNullString
    Dim strBuffer As String
    strBuffer = strBuffer & "Number of Api Calls: " & UBound(gApiList) & vbCrLf
    For i = 0 To UBound(gApiList) - 1
        If Left$(gApiList(i).strFunctionName, 7) = "Declare" Then

            strBuffer = strBuffer & gApiList(i).strFunctionName & vbCrLf
        Else
            strBuffer = strBuffer & "Declare " & gApiList(i).strFunctionName & " Lib " & cQuote & gApiList(i).strLibraryName & cQuote & vbCrLf
        End If
    Next
    frmMain.txtCode.Text = strBuffer
Exit Sub
errHandle:
    MsgBox "Error_modGlobals_WriteApiList: " & err.Number & " " & err.Description
End Sub
Public Function GetFirstChar(Start As Long, TextToFind As RichTextBox, ListToLike As String) As FIRSTCHAR_INFO
    Dim i As Long, Cursor As Long, TheChar As String, theCursor As Long, SStart As Long, SLength As Long
    SStart = TextToFind.SelStart
    SLength = TextToFind.SelLength
    Cursor = Len(TextToFind.Text)
    For i = 1 To Len(ListToLike)
        theCursor = TextToFind.Find(Mid$(ListToLike, i, 1), Start - 1) + 1
        If theCursor < Cursor And theCursor > 0 Then
            Cursor = theCursor
            TheChar = Mid$(ListToLike, i, 1)
        End If
    Next i
    TextToFind.SelStart = SStart
    TextToFind.SelLength = SLength
    If Cursor < Start Then
        Cursor = Start
    Else
        GetFirstChar.lCursor = Cursor
    End If
    GetFirstChar.sChar = TheChar
End Function


Sub AddPropertyToTheList(strPropertyName As String, value As Variant, VarType As String, offset As Long, HelpString As String, Optional ByVal PropType As peEditorType = 2)
'*****************************
'Purpose: Used for Form Editor. To add a textbox and label to hold property name and value
'*****************************
    Dim i As Integer
    i = frmMain.txtEditArray.UBound + 1
    
    Load frmMain.txtEditArray(i)
    'Dim pSubMainProp As CPropertyItem
    'With frmMain.pePropTree
    '    .LockWindowUpdate = True
    '    If VarType = "Boolean" Then
    '        Dim oListBoolean      As CListItems
    '        Set oListBoolean = .ListItems.Add("Boolean", peSimpleDropDown)
    '        oListBoolean.ListItemsAdd "True", "True", "True"
    '        oListBoolean.ListItemsAdd "False", "False", "False"
    '       .LockWindowUpdate = True
    '       Set pSubMainProp = .PropertyItems.AddPropertyItem(strPropertyName, strPropertyName, pMain, pelistEditor, oListBoolean)
    '        pSubMainProp.ReadOnly = False
    '        If value = "True" Then
    '            pSubMainProp.value = True
    '        Else
    '            pSubMainProp.value = False
    '        End If
    '        .LockWindowUpdate = False
    '    ElseIf UCase$(strPropertyName) = "FORECOLOR" Or UCase$(strPropertyName) = "BACKCOLOR" Or UCase$(strPropertyName) = "FILLCOLOR" Or UCase$(strPropertyName) = "MASKCOLOR" Then
    '        Set pSubMainProp = .PropertyItems.AddPropertyItem(strPropertyName, strPropertyName, pMain, peColorEditor)
    '        pSubMainProp.value = value
    '    Else
'
'                Set pSubMainProp = .PropertyItems.AddPropertyItem(strPropertyName, strPropertyName, pMain, PropType)
'                pSubMainProp.value = value'
'
  '      End If
 '       pSubMainProp.HelpString = HelpString
'
      '  pSubMainProp.Expanded = True

     '
     '   .LockWindowUpdate = False
    'End With
    
    
    
    With frmMain.txtEditArray(i)
        .Text = value
        .tag = VarType
        .Top = frmMain.txtEditArray(i - 1).Top + 300
        .Left = frmMain.txtEditArray(0).Left
        If VarType = "String" Then
            .MaxLength = Len(value)
        End If
        .ToolTipText = HelpString
        .Visible = True
    End With
    
    Load frmMain.lblArrayEdit(i)
    
    With frmMain.lblArrayEdit(i)
        .Caption = strPropertyName
        .Top = frmMain.lblArrayEdit(i - 1).Top + 300 ' frmMain.lblArrayEdit(i - 1).Top + frmMain.lblArrayEdit(i - 1).Height
        .Left = frmMain.lblArrayEdit(0).Left
        .tag = offset
        If VarType = "String" Then
        .tag = (offset - Len(value))
        End If
        
        .Visible = True
        
    End With
    
    Load frmMain.cmdColor(i)
    With frmMain.cmdColor(i)
        .Top = frmMain.cmdColor(i - 1).Top + 300
        .Left = frmMain.cmdColor(0).Left
        
        If strPropertyName = "BackColor" Or strPropertyName = "ForeColor" Or strPropertyName = "FillColor" Then
            .tag = "c"
            .Visible = True
        Else
            .Visible = False
        End If
        If strPropertyName = "FontName" Then
            .tag = "f"
            .Visible = True
        End If
        If strPropertyName = "Picture" Then
            .tag = "p"
            .Caption = "X"
            .ToolTipText = "Delete Image!"
            .Visible = False 'True
        End If
    End With
    'Check if change has been made?
    On Error GoTo errHandle
    Dim g As Long
    For g = 0 To UBound(ByteChange)
        If frmMain.lblArrayEdit(i).tag = ByteChange(g).offset Then
            frmMain.txtEditArray(i).Text = ByteChange(g).bByte
            Exit Sub
        End If
    Next g
    For g = 0 To UBound(LongChange)
        If frmMain.lblArrayEdit(i).tag = LongChange(g).offset Then
            frmMain.txtEditArray(i).Text = LongChange(g).lLong
            Exit Sub
        End If
    Next
    For g = 0 To UBound(StringChange)
        If frmMain.lblArrayEdit(i).tag = StringChange(g).offset Then
            frmMain.txtEditArray(i).Text = StringChange(g).sString
            Exit Sub
        End If
    Next g
    For g = 0 To UBound(SingleChange)
        If frmMain.lblArrayEdit(i).tag = SingleChange(g).offset Then
            frmMain.txtEditArray(i).Text = SingleChange(g).sSingle
            Exit Sub
        End If
    Next g
    For g = 0 To UBound(BooleanChange)
        If frmMain.lblArrayEdit(i).tag = BooleanChange(g).offset Then
            
            If BooleanChange(g).bBool = True Then
                frmMain.txtEditArray(i).Text = "True"
            Else
                frmMain.txtEditArray(i).Text = "False"
            End If
            Exit Sub
        End If
    Next g
    For g = 0 To UBound(IntegerChange)
        If frmMain.lblArrayEdit(i).tag = IntegerChange(g).offset Then
            frmMain.txtEditArray(i).Text = IntegerChange(g).iInt
            Exit Sub
        End If
    Next g
    
errHandle:
    'MsgBox strPropertyName
End Sub

Public Function FileExists(ByVal Path As String) As Boolean
'*****************************
'Purpose: Checks wether a FileExists or not
'*****************************
  If Len(Path) = 0 Then Exit Function
  If Dir(Path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> vbNullString Then FileExists = True
End Function
Sub LoadObjectType()
'*****************************
'Purpose: Checks wether a FileExists or not
'*****************************
    ReDim gObjectTypeList(15)
    gObjectTypeList(0).value = 98435
    gObjectTypeList(0).strType = 1
    gObjectTypeList(1).value = 98467
    gObjectTypeList(1).strType = 1
    gObjectTypeList(2).value = 98499
    gObjectTypeList(2).strType = 1
    gObjectTypeList(3).value = 98305
    gObjectTypeList(3).strType = 2
    gObjectTypeList(4).value = 98337
    gObjectTypeList(4).strType = 2
    gObjectTypeList(5).value = 1146883
    gObjectTypeList(5).strType = 3
    gObjectTypeList(6).value = 1277955
    gObjectTypeList(6).strType = 3
    gObjectTypeList(7).value = 98339
    gObjectTypeList(7).strType = 3
    gObjectTypeList(8).value = 100355
    gObjectTypeList(8).strType = 3
    gObjectTypeList(9).value = 1148931
    gObjectTypeList(9).strType = 3
   ' gObjectTypeList(10).value = 1148931
    'gObjectTypeList(10).strType = 4
    gObjectTypeList(11).value = 1941507
    gObjectTypeList(11).strType = 4
    gObjectTypeList(12).value = 1941539
    gObjectTypeList(12).strType = 4
    gObjectTypeList(13).value = 1943555
    gObjectTypeList(13).strType = 4
    gObjectTypeList(14).value = 1409027
    gObjectTypeList(14).strType = 5
    gObjectTypeList(15).value = 1411075
    gObjectTypeList(15).strType = 6
End Sub
Public Sub AddToErrorLog(ByVal strError As String)
    strErrorLog = strErrorLog & strError & vbCrLf
End Sub
Public Sub WriteErrorLog(ByVal strPath As String)
On Error GoTo nofile:
    If strErrorLog = "" Then Exit Sub
    Dim F As Long
    F = FreeFile
    Open strPath For Output As #F
        Print #F, "Error Log for " & SFilePath
        Print #F, strErrorLog
    Close #F
Exit Sub
nofile:
    MsgBox "Error_modGlobals_WriteErrorLog: " & err.Description
    Exit Sub
End Sub
Sub ClearErrorLog()
    strErrorLog = ""
End Sub
Public Sub GetStartUpName(FileNumber As Integer)
On Error GoTo errHandle:
    Dim TempData As Double
    Dim TempDouble As Double
    'ReferencePoint is EntryPoint + TEXTTABLEPTROFFSET
    'There are four possibilities:
    '1. The SUform exists at the ReferencePoint; if FFCC3100, SUM is form
    '2. Else, the SUform exists at (ReferencePoint + DWORD) +
    '   offset at (ReferencePoint + 3rd DWORD); if FFCC3100, SUM is form
    '3. Else, SUmodule exists as SubMain if the address at VB5! + $2C is non-zero
    '4. Else, there is no SUM; app is ActiveXDoc or equiv
          
    'Start at the VB start location
    Seek #FileNumber, OptHeader.EntryPoint + 1
    
    'Calculate the ReferencePoint
    TempData# = Seek(FileNumber) + TEXTTABLEPTROFFSET
    
    'Move to ReferencePoint
    Seek #FileNumber, TempData#
        
    'Get first half of flag at ReferencePoint
    TempDouble# = GetWordByFile(FileNumber)
        
    'Check for possibility #1...
    'Check if form exists = check for the MSB of the form header
    If TempDouble# = FRMHEADERMSB Then '1 OK
            
        'Get the second half of flag at ReferencePoint
        TempDouble# = GetWordByFile(FileNumber)
        
        'Check for the LSB of the form header, remove the high bit
        If (TempDouble# And Not FORMHIGHMASK) = FRMHEADERLSB Then '1 OK
        
            'Clear submain
            'SubMainPointer# = 0#
            
            'Save the object type
            AppData.StartUpType = FRMHEADERMSB
            
            'Move to ReferencePoint = start of form
            Seek #FileNumber, TempData#
            
            'Move to the form name length
            Seek #FileNumber, Seek(FileNumber) + STARTUPMODOFFSET3
        
            'Move to the form name
            Seek #FileNumber, Seek(FileNumber) + 2
            
            'Get the string that is there, save it
            AppData.StartUpName = GetDosString()
            
            'Display composite
            'frmRaceMain.lstCompositeFiles(PROJECTLISTINDEX).AddItem "Startup=" & Chr$(34) & AppData.StartUpName & Chr$(34)
            'MsgBox Loc(FileNumber)
            'Append the object type
            AppData.StartUpName = AppData.StartUpName
            
            'Move to ReferencePoint
            Seek #FileNumber, TempData#
            
            'Move to first size word
            Seek #FileNumber, Seek(FileNumber) + FORMSIZEWORDOFFSET
            
            'Get the first size MSB
            TempDouble# = GetByteByFile(FileNumber) * 256#
                        
            'Add in the first size LSB
            TempDouble# = TempDouble# + GetByteByFile(FileNumber)
            
            'Add in the second size MSB
            TempDouble# = TempDouble# + (GetByteByFile(FileNumber) * 256)
                        
            'Add in the second size LSB
            TempDouble# = TempDouble# + GetByteByFile(FileNumber)
            
            'Back up to first size word
            Seek #FileNumber, Seek(FileNumber) - 4
            
            'Move over rest of form
            Seek #FileNumber, Seek(FileNumber) + TempDouble# + 1
                    
        Else
        
            'Bad form flag
           ' ErrorFlag# = MY_MODULE_ID + ERR_BAD_CODE
        
        End If
        
    'Check for possibility #2,3, or 4...
    Else '2,3 OK
    
        'Move back to the ReferencePosition
        Seek #FileNumber, TempData#
        
        'Get 3rd DWORD Offset
        TempDouble# = GetDWordByFile(FileNumber)
        TempDouble# = GetDWordByFile(FileNumber)
        TempDouble# = GetDWordByFile(FileNumber)
        
        'Move back to the ReferencePosition
        Seek #FileNumber, TempData#
        
        'Move ahead 1 DWORD
        Seek #FileNumber, Seek(FileNumber) + 4
        
        'Move ahead by Offset value
        Seek #FileNumber, Seek(FileNumber) + TempDouble#
        
        'Get the data there
        TempDouble# = GetWordByFile(FileNumber)
        
        'Check if form exists = check for the MSB of the form header
        If TempDouble# = FRMHEADERMSB Then
                
            'Get the second half of flag
            TempDouble# = GetWordByFile(FileNumber)
            
            'Check for the LSB of the form header, remove the high bit
            If (TempDouble# And Not FORMHIGHMASK) = FRMHEADERLSB Then '2 OK
            
                'Clear submain
                'SubMainPointer# = 0#
            
                'Save the object type
                AppData.StartUpType = FRMHEADERMSB
                
                'Back up to the start of the form
                Seek #FileNumber, Seek(FileNumber) - 4
                
                'Move to the form name length
                Seek #FileNumber, Seek(FileNumber) + STARTUPMODOFFSET3
            
                'Move to the form name
                Seek #FileNumber, Seek(FileNumber) + 2
                'MsgBox Loc(FileNumber)
                'Get the string that is there, save it
                AppData.StartUpName = GetDosString()
                
                'Append the object type
                AppData.StartUpName = AppData.StartUpName
                
                'Move back to the ReferencePoint
                Seek #FileNumber, TempData#
                                 
            Else
                
                'Bad form flag
               ' ErrorFlag# = MY_MODULE_ID + ERR_BAD_CODE
                
            End If
            
        'Check for possibility #3 or 4....
        Else '3,4 OK
        
            'Point file at the VB signature position
            Seek #FileNumber, AppData.VBVerOffsetMasked + 1
    
            'Move to SUBMAIN pointer
            Seek #FileNumber, Seek(FileNumber) + &H2C
            
            'Get pointer
            TempDouble# = GetDWordByFile(FileNumber)
                        
            'If non-zero, must be SubMain
            If TempDouble# <> 0# Then '3 OK
            
                'Check for valid range
                If ((TempDouble# > DecLoadOffset#) And (TempDouble# < (LOF(FileNumber) + DecLoadOffset#))) Then
            
                    'Save the object type
                    AppData.StartUpType = MODSIGNATURE
                    
                    'Save the object name
                    AppData.StartUpName = "Sub Main"
                
                    'Display composite
                    'frmRaceMain.lstCompositeFiles(PROJECTLISTINDEX).AddItem "Startup=" & Chr$(34) & "SubMain" & Chr$(34)
            
                    'Calculate the APP offset
                    TempDouble# = TempDouble# - DecLoadOffset#
                    
                    'Check for compile type...If Pcode, get pointer;
                    'if Ncode, address is call command
                    If AppData.CompileType = "PCode" Then
                    
                        'Go there
                        Seek FileNumber, TempDouble# + 1
                    
                        'Skip 1 byte
                        TempDouble# = GetByteByFile(FileNumber)
                        
                        'Get pointer
                        TempDouble# = GetDWordByFile(FileNumber)
                        
                        'Calculate the APP offset
                        TempDouble# = TempDouble# - DecLoadOffset#
                        
                    End If
                    
                    'Save it
                   ' SubMainPointer# = TempDouble#
                                    
                    'Move back to the ReferencePosition
                    Seek #FileNumber, TempData#
            
                Else
                
                    'Set error flag
                   ' ErrorFlag# = MY_MODULE_ID + ERR_BAD_SUM
                    
                    'Exit
                    Exit Sub
                
                End If
                
            'Must be possibility #4
            Else '4 OK
                
                'Save the object type
                AppData.StartUpType = 999999 ' NOSIGNATURE
                
                'Save the object name
                AppData.StartUpName = "(NONE)"
                            
                'Move back to the ReferencePosition
                Seek #FileNumber, TempData#
                
            End If
    
        End If
        
    End If
Exit Sub
errHandle:
    Call AddToErrorLog(err.Number & " " & err.Description)
               ' MsgBox AppData.StartUpName
               ' MsgBox AppData.StartUpOffset
               ' MsgBox AppData.StartUpType
End Sub
Public Function GetPart(DataStr As String, DataId As Long, Separator As String) As Variant
    Dim pointer As Long, i As Long
    On Error GoTo errHandler
    For i = 1 To DataId
        pointer = InStr(pointer + 1, DataStr, Separator)
    Next i
    GetPart = Mid$(DataStr, pointer + Len(Separator), IIf(InStr(pointer + 1, DataStr, Separator) = 0, Len(DataStr), InStr(pointer + 1, DataStr, Separator)) - pointer - Len(Separator))
    Exit Function
errHandler:
    GetPart = False
End Function


