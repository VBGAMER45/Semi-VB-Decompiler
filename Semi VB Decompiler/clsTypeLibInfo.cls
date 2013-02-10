VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTypeLibInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Type Library Browser
'
' low level COM project - by [rm_code] 2005


' check if data can be read from a pointer
Private Declare Function IsBadReadPtr Lib "kernel32" ( _
    lp As Any, ByVal ucb As Long) As Long

' Change a Variants type and copy it
Private Declare Function VariantChangeType Lib "oleaut32" ( _
    ByVal pvargDest As Long, ByVal pvarSrc As Long, _
    ByVal wFlags As Integer, ByVal vt As Integer) As Long

' Call a function by its pointer
Private Declare Function CallWindowProcA Lib "user32" ( _
    ByVal adr As Long, ByVal p1 As Long, ByVal p2 As Long, _
    ByVal p3 As Long, ByVal p4 As Long) As Long

Private Declare Sub RtlFillMemory Lib "kernel32" ( _
    pDst As Any, ByVal dlen As Long, ByVal Fill As Byte)

' convert a CLSID string to a data structure
Private Declare Function CLSIDFromString Lib "ole32.dll" ( _
    ByVal lpszProgID As Long, pCLSID As Any) As Long

' find the ProgID of a CLSID
Private Declare Function ProgIDFromCLSID Lib "ole32.dll" ( _
    pCLSID As Any, lpszProgID As Long) As Long

' convert a GUID to a string
Private Declare Function StringFromGUID2 Lib "ole32" ( _
    tGuid As Any, ByVal lpszString As String, ByVal lMax As Long) As Long

' load a type library from a file
Private Declare Function LoadTypeLib Lib "oleaut32" ( _
    ByVal szFile As Long, pptlib As Any) As HRESULT

' copy memory from one location to the other
Private Declare Sub CpyMem Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDst As Any, pSrc As Any, ByVal dwLen As Long)

' free a string allocated by the system
Private Declare Sub SysFreeString Lib "oleaut32" ( _
    ByVal bstr As Long)



' IUnknown interface
' every other interface inherits from IUnknown
Private Type IUnknown
    QueryInterface          As Long
    AddRef                  As Long
    Release                 As Long
End Type

' ITypeInfo interface
' information about a type
Private Type ITypeInfo
    IUnk                    As IUnknown ' inherits from IUnknown
    GetTypeAttr             As Long     ' Retrieves a TYPEATTR structure that contains the attributes of the type description
    GetTypeComp             As Long     ' Retrieves the ITypeComp interface for the type description
    GetFuncDesc             As Long     ' Retrieves the FUNCDESC structure that contains information about a specified function
    GetVarDesc              As Long     ' Retrieves a VARDESC structure that describes the specified variable
    GetNames                As Long     ' Retrieves the variable with the specified member ID
    GetRefTypeOfImplType    As Long     ' retrieves the type description of the implemented interface types
    GetImplTypeFlags        As Long     ' Retrieves the IMPLTYPEFLAGS enumeration for one implemented interface
    GetIDsOfNames           As Long     ' Maps between member names and member IDs, and parameter names and parameter IDs
    Invoke                  As Long     ' Invokes a method, or accesses a property of an object
    GetDocumentation        As Long     ' Retrieves the documentation string, the complete Help file name and path, and the context ID for the Help topic for a specified type description
    GetDllEntry             As Long     ' Retrieves a description or specification of an entry point for a function in a DLL
    GetRefTypeInfo          As Long     ' retrieves the referenced type descriptions
    AddressOfMember         As Long     ' Retrieves the addresses of static functions or variables
    CreateInstance          As Long     ' Creates a new instance of a type that describes a component object class
    GetMops                 As Long     ' Retrieves marshaling information
    GetContainingTypeLib    As Long     ' Retrieves the containing type library and the index of the type description within that type library
    ReleaseTypeAttr         As Long     ' Releases a TYPEATTR previously returned by GetTypeAttr
    ReleaseFuncDesc         As Long     ' Releases a FUNCDESC previously returned by GetFuncDesc
    ReleaseVarDesc          As Long     ' Releases a VARDESC previously returned by GetVarDesc
End Type

' ITypeLib interface
' information about a type library
Private Type ITypeLib
    IUnk                    As IUnknown ' inherits from IUnknown
    GetTypeInfoCount        As Long     ' retrieves the number of type descriptions in the library
    GetTypeInfo             As Long     ' retrieves the specified type description in the library
    GetTypeInfoType         As Long     ' retrieves the type of a type description
    GetTypeInfoOfGuid       As Long     ' retrieves the type description that corresponds to the specified (GUID)
    GetLibAttr              As Long     ' retrieves the structure that contains the library's attributes
    GetTypeComp             As Long     ' retrieves a pointer to the ITypeComp for a type library
    GetDocumentation        As Long     ' retrieves the library's documentation string
    IsName                  As Long     ' indicates whether a passed-in string contains the name of a type or a member described in the library
    FindName                As Long     ' finds occurrences of a type description in a type library
    ReleaseTLibAttr         As Long     ' releases the TLIBATTR originally obtained from GetLibAttr
End Type

' parameter description
Private Type TPARAMDESC
    pPARAMDESCEX            As Long     ' valid if PARAMFLAG_FHASDEFAULT
    wParamFlags             As Integer  ' parameter flags (in,out,...)
End Type

' extended parameter description
Private Type TPARAMDESCEX
    cBytes                  As Long     ' size of structure
    varDefaultValue         As Variant  ' default value of parameter
End Type

' type description
Private Type TTYPEDESC
    pTypeDesc               As Long     ' vt = VT_PTR: points to another TYPEDESC
                                        ' vt = VT_CARRAY: points to another TYPEDESC
                                        ' vt = VT_USERDEFINED: pTypeDesc is a HREFTYPE instead of a pointer
    vt                      As Integer  ' vartype
End Type

' element description
Private Type TELEMDESC
    tdesc                   As TTYPEDESC    ' type description
    pdesc                   As TPARAMDESC   ' parameter description
End Type

' array description
Private Type TARRAYDESC
    tdescElem               As TTYPEDESC    ' type description
    cDims                   As Integer      ' number of dimensions
End Type

Private Type SAFEARRAYBOUND
    cElements               As Long
    lLBound                 As Long
End Type

' type attributes
Private Type TTYPEATTR
    guid(15)                As Byte     ' type GUID
    lcid                    As Long     ' lcid
    dwReserved              As Long     '
    memidConstructor        As Long     ' ID of constructor
    memidDestructor         As Long     ' ID of destructor
    pstrSchema              As Long     '
    cbSizeInstance          As Long     ' size of an instance
    TYPEKIND                As Long     ' kind of type
    cFuncs                  As Integer  ' functions
    cVars                   As Integer  ' variables
    cImplTypes              As Integer  ' implemented types
    cbSizeVft               As Integer  ' VTable size
    cbAlignment             As Integer  ' byte alignment
    wTypeFlags              As Integer  '
    wMajorVerNum            As Integer  ' major version
    wMinorVerNum            As Integer  ' minor version
    tdescAlias              As Long     ' specifies for which type this is an alias
    idldescType             As Long     ' IDL attributes
End Type

' function description
Private Type TFUNCDESC
    memid                   As Long     ' member ID
    lprgscode               As Long     '
    lprgelemdescParam       As Long     '
    funckind                As Long     ' virtual, static, ...
    invkind                 As Long     ' Invocation kind
    callconv                As Long     ' calling convention
    cParams                 As Integer  ' parameter count
    cParamsOpt              As Integer  ' optional parameters
    oVft                    As Integer  ' VTable offset
    cScodes                 As Integer  ' Count of permitted return values
    elemdesc                As TELEMDESC ' Contains the return type of the function
    wFuncFlags              As Integer  ' function flags
End Type

' variable description
Private Type TVARDESC
    memid                   As Long     ' member ID
    lpstrSchema             As Long     '
    uInstVal                As Long     ' vkind = VAR_PERINSTANCE: offset of this variable within the instance
                                        ' vkind = VAR_CONST: value of it as a variant
    elemdescVar             As TELEMDESC ' variable type
    wVarFlags               As Integer  ' variable flags
    vkind                   As Long     ' variable kind
End Type

' variant types
Public Enum Varenum
    VT_EMPTY = 0&                   '
    VT_NULL = 1&                    ' 0
    VT_I2 = 2&                      ' signed 2 bytes integer
    VT_I4 = 3&                      ' signed 4 bytes integer
    VT_R4 = 4&                      ' 4 bytes float
    VT_R8 = 5&                      ' 8 bytes float
    VT_CY = 6&                      ' currency
    VT_DATE = 7&                    ' date
    VT_BSTR = 8&                    ' BStr
    VT_DISPATCH = 9&                ' IDispatch
    VT_ERROR = 10&                  ' error value
    VT_BOOL = 11&                   ' boolean
    VT_VARIANT = 12&                ' variant
    VT_UNKNOWN = 13&                ' IUnknown
    VT_DECIMAL = 14&                ' decimal
    VT_I1 = 16&                     ' signed byte
    VT_UI1 = 17&                    ' unsigned byte
    VT_UI2 = 18&                    ' unsigned 2 bytes integer
    VT_UI4 = 19&                    ' unsigned 4 bytes integer
    VT_I8 = 20&                     ' signed 8 bytes integer
    VT_UI8 = 21&                    ' unsigned 8 bytes integer
    VT_INT = 22&                    ' integer
    VT_UINT = 23&                   ' unsigned integer
    VT_VOID = 24&                   ' 0
    VT_HRESULT = 25&                ' HRESULT
    VT_PTR = 26&                    ' pointer
    VT_SAFEARRAY = 27&              ' safearray
    VT_CARRAY = 28&                 ' carray
    VT_USERDEFINED = 29&            ' userdefined
    VT_LPSTR = 30&                  ' LPStr
    VT_LPWSTR = 31&                 ' LPWStr
    VT_RECORD = 36&                 ' Record
    VT_FILETIME = 64&               ' File Time
    VT_BLOB = 65&                   ' Blob
    VT_STREAM = 66&                 ' Stream
    VT_STORAGE = 67&                ' Storage
    VT_STREAMED_OBJECT = 68&        ' Streamed Obj
    VT_STORED_OBJECT = 69&          ' Stored Obj
    VT_BLOB_OBJECT = 70&            ' Blob Obj
    VT_CF = 71&                     ' CF
    VT_CLSID = 72&                  ' Class ID
    VT_BSTR_BLOB = &HFFF&           ' BStr Blob
    VT_VECTOR = &H1000&             ' Vector
    VT_ARRAY = &H2000&              ' Array
    VT_BYREF = &H4000&              ' ByRef
    VT_RESERVED = &H8000&           ' Reserved
    VT_ILLEGAL = &HFFFF&            ' illegal
End Enum

' a type`s type
Public Enum TYPEKIND
    TKIND_ENUM = 0                  ' enums
    TKIND_RECORD                    ' types
    TKIND_MODULE                    ' modules
    TKIND_INTERFACE                 ' interfaces
    TKIND_DISPATCH                  ' dispatches
    TKIND_COCLASS                   ' CoClasses
    TKIND_ALIAS                     ' aliases
    TKIND_UNION                     ' Unions
    TKIND_MAX                       ' ?
End Enum

' invoke kind of a member
Public Enum INVOKEKIND
    INVOKE_FUNC = &H1               ' function
    INVOKE_PROPERTY_GET = &H2       ' property
    INVOKE_PROPERTY_PUT = &H4       ' property
    INVOKE_PROPERTY_PUTREF = &H8    ' property
End Enum

' variable kind
Public Enum VARKIND
    VAR_PERSISTANCE = 0             '
    VAR_STATIC                      '
    VAR_CONST                       '
    VAR_DISPATCH                    '
End Enum

' parameter flags
Public Enum PARAMFLAGS
    PARAMFLAG_NONE = &H0            ' ...
    PARAMFLAG_FIN = &H1             ' in
    PARAMFLAG_FOUT = &H2            ' out
    PARAMFLAG_FLCID = &H4           ' lcid
    PARAMFLAG_FRETVAL = &H8         ' return value
    PARAMFLAG_FOPT = &H10           ' optional
    PARAMFLAG_FHASDEFAULT = &H20    ' default value
    PARAMFLAG_FHASCUSTDATA = &H40   ' custom data
End Enum

' HRESULT type
Private Enum HRESULT
    S_OK = 0
End Enum

' current file
Private strFile         As String

' ITypeLib pointer and VTable
Private oTypeLib        As Long
Private udtTypeLib      As ITypeLib

' ITypeInfo pointer and VTable
Private oTypeInfo       As Long
Private udtTypeInfo     As ITypeInfo

' counters
Private cVars           As Long
Private cTypeInfos      As Long
Private lngCurInfo      As Long
Private lngCurFnc       As Long
Private lngCurFncParam  As Long
Private lngCurImpl      As Long

' descriptions for the current type/function
Private curTypeAttr     As TTYPEATTR
Private curFuncDesc     As TFUNCDESC
Private curVarDesc      As TVARDESC
Private curTypeDesc     As TTYPEDESC

' function name and parameter names
Private strNames()      As String

Private blnOpen         As Boolean
Private blnTypeOpen     As Boolean

' max allowed code for CallPointer`s ASM
Private Const MAXCODE   As Long = &HEC00&

' currently opened file
Public Property Get Filename() As String
    Filename = strFile
End Property

' only for alias types
Public Property Get AliasName() As String
    If Not TypeInfoKind = TKIND_ALIAS Then Exit Property
    AliasName = StringifyTypeDesc(VarPtr(curTypeAttr.tdescAlias))
End Property

' default value of a parameter
'
' /// something in here is screwed up...
' /// maybe an invalid variant pointer
' /// is returned in PARAMDESCEX
'
'Public Property Get ParameterDefault() As String
'    Dim edesc   As TPARAMDESC
'    Dim descex  As TPARAMDESCEX
'    Dim v       As Variant
'    Dim ptr     As Long
'    Dim ret     As HRESULT
'
'    ' get the ELEMDESC of the current parameter
'    ' from the FUNCDESC.lprgelemdescParam array
'    ptr = curFuncDesc.lprgelemdescParam + (lngCurFncParam * Len(edesc))
'    CpyMem edesc, ByVal ptr, LenB(edesc)
'
'    ' parameter has the DEFAULT flag?
'    If 0 = (edesc.wParamFlags And PARAMFLAG_FHASDEFAULT) Then
'        Exit Property
'    End If
'
'    ' stay away from nasty GPFs
'    If 0 <> IsBadReadPtr(ByVal edesc.pPARAMDESCEX, Len(descex)) Then
'        Exit Property
'    End If
'
'    ' get PARAMDESCEX for current member
'    CpyMem descex, ByVal edesc.pPARAMDESCEX, Len(descex)
'
'    ' convert the variant's type to BSTR (= VB String)
'    VariantChangeType VarPtr(v), VarPtr(descex.varDefaultValue), 0, VT_BSTR
'
'    ' return the default value
'    ParameterDefault = v
'End Property

' flags of the current parameter
Public Property Get ParameterFlags() As PARAMFLAGS
    Dim edesc   As TELEMDESC
    Dim ptr     As Long

    ' get the ELEMDESC of the current parameter
    ' from the FUNCDESC.lprgelemdescParam array
    ptr = curFuncDesc.lprgelemdescParam + (lngCurFncParam * Len(edesc))
    CpyMem edesc, ByVal ptr, LenB(edesc)

    ' return the flags
    ParameterFlags = edesc.pdesc.wParamFlags
End Property

' parameter type
Public Property Get ParameterType() As String
    Dim edesc   As TELEMDESC
    Dim ptr     As Long

    ' get the ELEMDESC of the current parameter
    ' from the FUNCDESC.lprgelemdescParam array
    ptr = curFuncDesc.lprgelemdescParam + (lngCurFncParam * Len(edesc))
    CpyMem edesc, ByVal ptr, Len(edesc)

    ' return the type of the parameter as a string
    ParameterType = StringifyTypeDesc(VarPtr(edesc.tdesc))
End Property

' name of current parameter
Public Property Get ParameterName() As String
    ParameterName = strNames(lngCurFncParam + 1)
End Property

' name of current variable
Public Property Get VariableName() As String
    Dim ret     As HRESULT
    Dim strName As String
    Dim dummy   As Long

    ' call ITypeInfo.GetNames with the
    ' member ID of the current variable
    ret = CallPointer(udtTypeInfo.GetNames, oTypeInfo, curVarDesc.memid, VarPtr(strName), 1, VarPtr(dummy))
    If ret = S_OK Then VariableName = strName
End Property

' variable type
Public Property Get VariableType() As String
    ' convert the type of the variable to a string
    VariableType = StringifyTypeDesc(VarPtr(curVarDesc.elemdescVar.tdesc))
End Property

' variable value
Public Property Get VariableValue() As String
    Dim v   As Variant
    Dim ret As HRESULT

    ' we can only get the value of CONST vars
    ' e.g. those stored in an enum
    If curVarDesc.vkind <> VAR_CONST Then
        VariableValue = "?"
        Exit Property
    End If

    ' convert its value to a BSTR (= VB String)
    ret = VariantChangeType(VarPtr(v), curVarDesc.uInstVal, 0, VT_BSTR)
    If ret <> S_OK Then
        VariableValue = "?"
        Exit Property
    End If

    VariableValue = v
End Property

' kind of the current variable
Public Property Get VariableKind() As VARKIND
    VariableKind = curVarDesc.vkind
End Property

' number of variables in the current type
Public Property Get VariableCount() As Long
    VariableCount = curTypeAttr.cVars
End Property

' number of parameters in the current function
Public Property Get ParameterCount() As Long
    ParameterCount = curFuncDesc.cParams
End Property

' VTable offset of current function
Public Property Get FunctionVTOffset() As Long
    FunctionVTOffset = curFuncDesc.oVft
End Property

' return type of the current function
Public Property Get FunctionReturnType() As String
    ' convert the DESCTYPE to a string
    FunctionReturnType = StringifyTypeDesc(VarPtr(curFuncDesc.elemdesc.tdesc))
End Property

' name of current function
Public Property Get FunctionName() As String
    FunctionName = strNames(0)
End Property

' current function`s invoke kind (function/property)
Public Property Get FunctionInvKind() As INVOKEKIND
    FunctionInvKind = curFuncDesc.invkind
End Property

' number of interfaces the current type implements
Public Property Get TypeInfoImplements() As Long
    TypeInfoImplements = curTypeAttr.cImplTypes
End Property

' number of functions the type implements
Public Property Get TypeInfoFunctions() As Long
    TypeInfoFunctions = curTypeAttr.cFuncs
End Property

' Prog ID of the current type
Public Property Get TypeInfoPrgID() As String
    TypeInfoPrgID = CLSIDToProgID(TypeInfoGUID)
End Property

' GUID of current type
Public Property Get TypeInfoGUID() As String
    TypeInfoGUID = GUID2Str(curTypeAttr.guid)
End Property

' kind of current type
Public Property Get TypeInfoKind() As TYPEKIND
    TypeInfoKind = curTypeAttr.TYPEKIND
End Property

' name of current type
Public Property Get TypeInfoName() As String
    Dim strName As String
    Dim ret     As HRESULT

    ' call ITypeLib.GetDocumentation
    ret = CallPointer(udtTypeLib.GetDocumentation, oTypeLib, lngCurInfo, VarPtr(strName), 0, 0, 0)
    If ret = S_OK Then TypeInfoName = strName
End Property

' number of type infos
Public Property Get TypeInfoCount() As Long
    TypeInfoCount = cTypeInfos
End Property

' type lib name
Public Property Get TypeLibName() As String
    Dim strName As String
    Dim ret     As HRESULT

    ' call ITypeLib.GetDocumentation
    ret = CallPointer(udtTypeLib.GetDocumentation, oTypeLib, -1, VarPtr(strName), 0, 0, 0)
    If ret = S_OK Then TypeLibName = strName
End Property

' GUID of current implement
Public Property Get ImplementGUID() As String
    Dim ret     As HRESULT
    Dim Itnfo   As ITypeInfo
    Dim oITnfo  As Long
    Dim href    As Long
    Dim pVTbl   As Long
    Dim tattr   As TTYPEATTR
    Dim pAttr   As Long

    ' get a new ITypeInfo for the current implement
    ret = CallPointer(udtTypeInfo.GetRefTypeOfImplType, oTypeInfo, lngCurImpl, VarPtr(href))
    ret = CallPointer(udtTypeInfo.GetRefTypeInfo, oTypeInfo, href, VarPtr(oITnfo))

    ' get the VTable
    CpyMem pVTbl, ByVal oITnfo, 4
    CpyMem Itnfo, ByVal pVTbl, Len(Itnfo)

    ' call ITypeInfo.GetTypeAttr to get the
    ' TYPEATTR structure for the current imeplement
    ret = CallPointer(Itnfo.GetTypeAttr, oITnfo, VarPtr(pAttr))
    CpyMem tattr, ByVal pAttr, Len(tattr)
    ret = CallPointer(Itnfo.ReleaseTypeAttr, oITnfo, pAttr)

    ' return the GUID of the implement
    ImplementGUID = GUID2Str(tattr.guid)
End Property

' current implement`s name
Public Property Get ImplementName() As String
    Dim ret     As HRESULT
    Dim Itnfo   As ITypeInfo
    Dim oITnfo  As Long
    Dim href    As Long
    Dim pVTbl   As Long
    Dim strName As String

    ' get a new ITypeInfo for the current implement
    ret = CallPointer(udtTypeInfo.GetRefTypeOfImplType, oTypeInfo, lngCurImpl, VarPtr(href))
    ret = CallPointer(udtTypeInfo.GetRefTypeInfo, oTypeInfo, href, VarPtr(oITnfo))

    ' get the VTable
    CpyMem pVTbl, ByVal oITnfo, 4
    CpyMem Itnfo, ByVal pVTbl, Len(Itnfo)

    ' call ITypeInfo.GetDocumentation
    ret = CallPointer(Itnfo.GetDocumentation, oITnfo, -1, VarPtr(strName), 0, 0, 0)
    If ret = S_OK Then ImplementName = strName
End Property

' select a implement of the current type
Public Function SelectImplement(index As Long) As Boolean
    If index > curTypeAttr.cImplTypes Then Exit Function
    SelectImplement = True
    lngCurImpl = index
End Function

' select a parameter of the current function
Public Function SelectParameter(index As Long) As Boolean
    If index > curFuncDesc.cParams Then Exit Function
    lngCurFncParam = index
    SelectParameter = True
End Function

' select a variable of the current type
Public Function SelectVariable(index As Long) As Boolean
    Dim ret     As HRESULT
    Dim pVDesc  As Long

    If index > curTypeAttr.cVars Then Exit Function

    ' get the VARDESC of the variable
    ret = CallPointer(udtTypeInfo.GetVarDesc, oTypeInfo, index, VarPtr(pVDesc))
    If ret = S_OK Then
        CpyMem curVarDesc, ByVal pVDesc, LenB(curVarDesc)
        ret = CallPointer(udtTypeInfo.ReleaseVarDesc, oTypeInfo, pVDesc)
        SelectVariable = True
    End If
End Function

' select a function of the current type
Public Function SelectFunction(index As Long) As Boolean
    Dim ret         As HRESULT
    Dim pFncDesc    As Long
    Dim cFncs       As Long

    If index >= curTypeAttr.cFuncs Then Exit Function

    ret = CallPointer(udtTypeInfo.GetFuncDesc, oTypeInfo, index, VarPtr(pFncDesc))

    If ret = S_OK Then
        ' get FUNCDESC...
        CpyMem curFuncDesc, ByVal pFncDesc, Len(curFuncDesc)
        ret = CallPointer(udtTypeInfo.ReleaseFuncDesc, oTypeInfo, pFncDesc)

        ReDim strNames(curFuncDesc.cParams) As String
        ' ... the name of the function,
        ' and the name of its parameters
        ret = CallPointer(udtTypeInfo.GetNames, oTypeInfo, curFuncDesc.memid, VarPtr(strNames(0)), 1 + curFuncDesc.cParams, VarPtr(cFncs))
        SelectFunction = True
        lngCurFnc = index
    End If
End Function

' select a type
Public Function SelectTypeInfo(index As Long) As Boolean
    Dim ret     As HRESULT
    Dim pVTbl   As Long
    Dim pAttr   As Long

    If index >= cTypeInfos Then Exit Function

    If blnTypeOpen Then
        ' release the old type info to free memory
        CallPointer udtTypeInfo.IUnk.Release, oTypeInfo
        blnTypeOpen = False
    End If

    ' get a new ITypeInfo
    ret = CallPointer(udtTypeLib.GetTypeInfo, oTypeLib, index, VarPtr(oTypeInfo))

    If ret = S_OK Then
        ' get the VTable of ITypeLib
        CpyMem pVTbl, ByVal oTypeInfo, 4
        CpyMem udtTypeInfo, ByVal pVTbl, Len(udtTypeInfo)

        ' get TYPEATTR of ITypeInfo
        CallPointer udtTypeInfo.GetTypeAttr, oTypeInfo, VarPtr(pAttr)
        CpyMem curTypeAttr, ByVal pAttr, Len(curTypeAttr)
        CallPointer udtTypeInfo.ReleaseTypeAttr, oTypeInfo, pAttr

        SelectTypeInfo = True
        blnTypeOpen = True
        lngCurInfo = index
    End If
End Function

' open a type library
Public Function OpenTypeLib(ByVal file As String) As Boolean
    Dim ret         As HRESULT
    Dim pVTbl       As Long

    If blnOpen Then CloseTypeLib: blnOpen = False
    ret = LoadTypeLib(StrPtr(file), oTypeLib)

    If ret = S_OK Then
        ' get the VTable of ITypeLib
        CpyMem pVTbl, ByVal oTypeLib, 4
        CpyMem udtTypeLib, ByVal pVTbl, Len(udtTypeLib)

        ' get the number of types
        cTypeInfos = CallPointer(udtTypeLib.GetTypeInfoCount, oTypeLib)
        lngCurInfo = -1

        strFile = file
        OpenTypeLib = True

        blnOpen = True
        blnTypeOpen = False
    End If
End Function

Public Sub CloseTypeLib()
    ' release ITypeLib and free memory
    If Not blnOpen Then Exit Sub
    CallPointer udtTypeLib.IUnk.Release, oTypeLib
    strFile = vbNullString
    blnOpen = False
End Sub

' TYPEDESC to string
' http://spec.winprog.org/typeinf2/
Private Function StringifyTypeDesc(ptr As Long) As String
    Dim strName As String
    Dim tdesc   As TTYPEDESC
    Dim arrdesc As TARRAYDESC
    Dim sfb     As SAFEARRAYBOUND
    Dim i       As Long

    ' TYPEDESC from pointer
    CpyMem tdesc, ByVal ptr, Len(tdesc)

    ' is a pointer to another TYPEDESC?
    If tdesc.vt = VT_PTR Then
        ' get the TYPEDESC
        strName = StringifyTypeDesc(tdesc.pTypeDesc) & "*"
        StringifyTypeDesc = strName
        Exit Function
    End If

    ' is a SAFEARRAY?
    If tdesc.vt = VT_SAFEARRAY Then
        ' get the TYPEDESC
        strName = "SAFEARRAY(" & StringifyTypeDesc(tdesc.pTypeDesc) & ")"
        StringifyTypeDesc = strName
        Exit Function
    End If

    ' is a CARRAY?
    If tdesc.vt = VT_CARRAY Then
        ' get the array descriptor
        CpyMem arrdesc, ByVal tdesc.pTypeDesc, Len(arrdesc)
        ' get the type of the array
        strName = StringifyTypeDesc(VarPtr(arrdesc.tdescElem))
        ' get dimensions
        For i = 0 To arrdesc.cDims
            CpyMem sfb, ByVal tdesc.pTypeDesc + ((i + 1) * 8) + 4, Len(sfb)
            strName = strName & "[" & sfb.lLBound & "..." & sfb.cElements + sfb.lLBound - 1 & "]"
        Next
        StringifyTypeDesc = strName
        Exit Function
    End If

    ' is a userdefined type?
    If tdesc.vt = VT_USERDEFINED Then
        ' get its name
        strName = stringifyCustomType(tdesc.pTypeDesc)
        StringifyTypeDesc = strName
        Exit Function
    End If

    Select Case tdesc.vt
        Case VT_NULL: strName = "Long"
        Case VT_I2: strName = "short"
        Case VT_I4: strName = "long"
        Case VT_R4: strName = "float"
        Case VT_R8: strName = "double"
        Case VT_CY: strName = "CY"
        Case VT_DATE: strName = "DATE"
        Case VT_BSTR: strName = "BSTR"
        Case VT_DISPATCH: strName = "IDispatch*"
        Case VT_ERROR: strName = "SCODE"
        Case VT_BOOL: strName = "VARIANT_BOOL"
        Case VT_VARIANT: strName = "VARIANT"
        Case VT_UNKNOWN: strName = "IUnknown*"
        Case VT_UI1: strName = "BYTE"
        Case VT_DECIMAL: strName = "DECIMAL"
        Case VT_I1: strName = "char"
        Case VT_UI2: strName = "USHORT"
        Case VT_UI4: strName = "ULONG"
        Case VT_I8: strName = "__int64"
        Case VT_UI8: strName = "unsigned __int64"
        Case VT_INT: strName = "int"
        Case VT_UINT: strName = "UINT"
        Case VT_HRESULT: strName = "HRESULT"
        Case VT_VOID: strName = "void"
        Case VT_LPSTR: strName = "char*"
        Case VT_LPWSTR: strName = "wchar_t*"
        Case Else: strName = "Any"
    End Select

    StringifyTypeDesc = strName
End Function

' get the name of a custom type
' http://spec.winprog.org/typeinf2/
Private Function stringifyCustomType(hreftype As Long) As String
    Dim ret     As HRESULT
    Dim oTInfo  As Long
    Dim tinfo   As ITypeInfo
    Dim strName As String
    Dim pVTbl   As Long

    ' call ITypeInfo.GetRefTypeInfo
    ret = CallPointer(udtTypeInfo.GetRefTypeInfo, oTypeInfo, hreftype, VarPtr(oTInfo))
    If ret <> S_OK Then
        stringifyCustomType = "?"
        Exit Function
    End If

    ' get the VTable of the new ITypeInfo
    CpyMem pVTbl, ByVal oTInfo, 4
    CpyMem tinfo, ByVal pVTbl, Len(tinfo)

    ' call ITypeInfo.GetDocumentation
    ret = CallPointer(tinfo.GetDocumentation, oTInfo, -1, VarPtr(strName), 0, 0, 0)
    If ret <> S_OK Then
        stringifyCustomType = "?"
        Exit Function
    End If

    stringifyCustomType = strName
End Function

' http://www.aboutvb.de/khw/artikel/khwcreateguid.htm
Private Function GUID2Str(GUIDBytes() As Byte) As String
    Dim nTemp       As String
    Dim nGUID(15)   As Byte
    Dim nLength     As Long

    nTemp = Space$(78)
    CpyMem nGUID(0), GUIDBytes(0), 16
    nLength = StringFromGUID2(nGUID(0), nTemp, Len(nTemp))
    GUID2Str = Left$(StrConv(nTemp, vbFromUnicode), nLength - 1)
End Function

' http://www.devx.com/vb2themax/Tip/18927
Private Function CLSIDToProgID(ByVal CLSID As String) As String
    Dim pResult     As Long, pChar      As Long
    Dim char        As Integer, length  As Long
    Dim guid(15)    As Byte

    CLSIDFromString StrPtr(CLSID), guid(0)
    ProgIDFromCLSID guid(0), pResult
    If pResult = 0 Then Exit Function

    pChar = pResult - 2
    Do
        pChar = pChar + 2
        CpyMem char, ByVal pChar, 2
    Loop While char

    length = pChar - pResult
    CLSIDToProgID = Space$(length \ 2)
    CpyMem ByVal StrPtr(CLSIDToProgID), ByVal pResult, length
End Function

' call a pointer with the help of some assembler
Private Function CallPointer(ByVal fnc As Long, ParamArray params()) As Long
    Dim btASM(MAXCODE - 1)  As Byte
    Dim pASM                As Long
    Dim i                   As Integer

    pASM = VarPtr(btASM(0))

    RtlFillMemory ByVal pASM, MAXCODE, &HCC

    AddByte pASM, &H58                  ' POP EAX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H50                  ' PUSH EAX

    If UBound(params) = 0 Then
        If IsArray(params(0)) Then
            For i = UBound(params(0)) To 0 Step -1
                AddPush pASM, CLng(params(0)(i))    ' PUSH dword
            Next
        Else
            For i = UBound(params) To 0 Step -1
                AddPush pASM, CLng(params(i))       ' PUSH dword
            Next
        End If
    Else
        For i = UBound(params) To 0 Step -1
            AddPush pASM, CLng(params(i))           ' PUSH dword
        Next
    End If

    AddCall pASM, fnc                   ' CALL rel addr
    AddByte pASM, &HC3                  ' RET

    CallPointer = CallWindowProcA(VarPtr(btASM(0)), _
                                  0, 0, 0, 0)
End Function

Private Sub AddPush(pASM As Long, lng As Long)
    AddByte pASM, &H68
    AddLong pASM, lng
End Sub

Private Sub AddCall(pASM As Long, addr As Long)
    AddByte pASM, &HE8
    AddLong pASM, addr - pASM - 4
End Sub

Private Sub AddLong(pASM As Long, lng As Long)
    CpyMem ByVal pASM, lng, 4
    pASM = pASM + 4
End Sub

Private Sub AddByte(pASM As Long, bt As Byte)
    CpyMem ByVal pASM, bt, 1
    pASM = pASM + 1
End Sub

Public Function ParamFlags2String(ByVal P As PARAMFLAGS) As String
    Dim strp    As String

    If P And PARAMFLAG_FIN Then _
        strp = strp & "in, "
    If P And PARAMFLAG_FOUT Then _
        strp = strp & "out, "
    If P And PARAMFLAG_FLCID Then _
        strp = strp & "lcid, "
    If P And PARAMFLAG_FRETVAL Then _
        strp = strp & "retval, "
    If P And PARAMFLAG_FOPT Then _
        strp = strp & "optional, "

    If Right$(strp, 2) = ", " Then
        strp = Left$(strp, Len(strp) - 2)
    End If

    ParamFlags2String = strp
End Function

Public Function VarKind2String(ByVal k As VARKIND) As String
    Select Case k
        Case VAR_CONST:             VarKind2String = "Const"
        Case VAR_DISPATCH:          VarKind2String = "Dispatch"
        Case VAR_PERSISTANCE:       VarKind2String = "Persistance"
        Case VAR_STATIC:            VarKind2String = "Static"
    End Select
End Function

Public Function InvKind2String(ByVal k As INVOKEKIND, Optional ignorefnc As Boolean = False) As String
    Select Case k
        Case INVOKE_FUNC:            InvKind2String = IIf(ignorefnc, "", "Function")
        Case INVOKE_PROPERTY_GET:    InvKind2String = "Property Get"
        Case INVOKE_PROPERTY_PUT:    InvKind2String = "Property Let"
        Case INVOKE_PROPERTY_PUTREF: InvKind2String = "Property Set"
    End Select
End Function

Public Function TKind2String(ByVal k As TYPEKIND) As String
    Select Case k
        Case TKIND_ALIAS:           TKind2String = "Alias"
        Case TKIND_COCLASS:         TKind2String = "CoClass"
        Case TKIND_DISPATCH:        TKind2String = "Dispatch"
        Case TKIND_ENUM:            TKind2String = "Enum"
        Case TKIND_INTERFACE:       TKind2String = "Interface"
        Case TKIND_MAX:             TKind2String = "Max"
        Case TKIND_MODULE:          TKind2String = "Module"
        Case TKIND_RECORD:          TKind2String = "Record"
        Case TKIND_UNION:           TKind2String = "Union"
    End Select
End Function
