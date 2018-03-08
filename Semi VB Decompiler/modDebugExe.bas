Attribute VB_Name = "modDebugExe"
'#############################################
'modDebugExe
'##############################################


Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type



Public Enum ProcessCreationFlags
   DEBUG_PROCESS = &H1
   DEBUG_ONLY_THIS_PROCESS = &H2
   CREATE_SUSPENDED = &H4
   DETACHED_PROCESS = &H8
   CREATE_NEW_CONSOLE = &H10
   NORMAL_PRIORITY_CLASS = &H20
   IDLE_PRIORITY_CLASS = &H40
   HIGH_PRIORITY_CLASS = &H80
   REALTIME_PRIORITY_CLASS = &H100
   CREATE_NEW_PROCESS_GROUP = &H200
   CREATE_UNICODE_ENVIRONMENT = &H400
   CREATE_SEPARATE_WOW_VDM = &H800
   CREATE_SHARED_WOW_VDM = &H1000
   CREATE_FORCEDOS = &H2000
   CREATE_DEFAULT_ERROR_MODE = &H4000000
   CREATE_NO_WINDOW = &H8000000
End Enum

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
   ByVal lpProcessAttributes As Long, _
   ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, _
   ByVal dwCreationFlags As ProcessCreationFlags, _
   ByVal lpEnvironment As Long, _
   ByVal lpCurrentDirectory As String, _
   lpStartupInfo As STARTUPINFO, _
   lpProcessInformation As PROCESS_INFORMATION) As Long

Private Type DEBUGHOOKINFO
    hModuleHook As Long
    Reserved As Long
    lParam As Long
    wParam As Long
    code As Long
End Type



Private Declare Function DebugActiveProcess Lib "kernel32" _
(ByVal dwProcessId As Long) As Long

'Private Declare Function WaitForDebugEvent Lib "kernel32" _
(lpDebugEvent As DEBUG_EVENT_BUFFER, _
 ByVal dwMilliseconds As Long) As Long

Private Type DEBUG_EVENT_HEADER
   dwDebugEventCode As DebugEventTypes
   dwProcessId As Long
   dwThreadId As Long
End Type

Public Enum DebugEventTypes
   EXCEPTION_DEBUG_EVENT = 1&
   CREATE_THREAD_DEBUG_EVENT = 2&
   CREATE_PROCESS_DEBUG_EVENT = 3&
   EXIT_THREAD_DEBUG_EVENT = 4&
   EXIT_PROCESS_DEBUG_EVENT = 5&
   LOAD_DLL_DEBUG_EVENT = 6&
   UNLOAD_DLL_DEBUG_EVENT = 7&
   OUTPUT_DEBUG_STRING_EVENT = 8&
   RIP_EVENT = 9&
End Enum

Public Enum ExceptionCodes
   EXCEPTION_GUARD_PAGE_VIOLATION = &H80000001
   EXCEPTION_DATATYPE_MISALIGNMENT = &H80000002
   EXCEPTION_BREAKPOINT = &H80000003
   EXCEPTION_SINGLE_STEP = &H80000004
   EXCEPTION_ACCESS_VIOLATION = &HC0000005
   EXCEPTION_IN_PAGE_ERROR = &HC0000006
   EXCEPTION_INVALID_HANDLE = &HC0000008
   EXCEPTION_NO_MEMORY = &HC0000017
   EXCEPTION_ILLEGAL_INSTRUCTION = &HC000001D
   EXCEPTION_NONCONTINUABLE_EXCEPTION = &HC0000025
   EXCEPTION_INVALID_DISPOSITION = &HC0000026
   EXCEPTION_ARRAY_BOUNDS_EXCEEDED = &HC000008C
   EXCEPTION_FLOAT_DENORMAL_OPERAND = &HC000008D
   EXCEPTION_FLOAT_DIVIDE_BY_ZERO = &HC000008E
   EXCEPTION_FLOAT_INEXACT_RESULT = &HC000008F
   EXCEPTION_FLOAT_INVALID_OPERATION = &HC0000090
   EXCEPTION_FLOAT_OVERFLOW = &HC0000091
   EXCEPTION_FLOAT_STACK_CHECK = &HC0000092
   EXCEPTION_FLOAT_UNDERFLOW = &HC0000093
   EXCEPTION_INTEGER_DIVIDE_BY_ZERO = &HC0000094
   EXCEPTION_INTEGER_OVERFLOW = &HC0000095
   EXCEPTION_PRIVILEGED_INSTRUCTION = &HC0000096
   EXCEPTION_STACK_OVERFLOW = &HC00000FD
   EXCEPTION_CONTROL_C_EXIT = &HC000013A
End Enum

Public Enum ExceptionFlags
   EXCEPTION_CONTINUABLE = 0
   EXCEPTION_NONCONTINUABLE = 1    '\ Noncontinuable exception
End Enum

Const EXCEPTION_MAXIMUM_PARAMETERS As Long = 15

Private Type DEBUG_EXCEPTION_DEBUG_INFO
   Header            As DEBUG_EVENT_HEADER
   ExceptionCode     As ExceptionCodes
   ExceptionFlags    As ExceptionFlags
   pExceptionRecord  As Long
   ExceptionAddress  As Long
   NumberParameters  As Long
   ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS)    As Long
   dwFirstChance     As Long
End Type

Private Type DEBUG_CREATE_THREAD_DEBUG_INFO
   Header As DEBUG_EVENT_HEADER
   hThread As Long
   lpThreadLocalBase As Long
   lpStartAddress As Long
End Type

Private Type DEBUG_CREATE_PROCESS_DEBUG_INFO
   Header                As DEBUG_EVENT_HEADER
   hfile                 As Long
   hProcess              As Long
   hThread               As Long
   lpBaseOfImage         As Long
   dwDebugInfoFileOffset As Long
   nDebugInfoSize        As Long
   lpThreadLocalBase     As Long
   lpStartAddress        As Long
   lpImageName           As Long
   fUnicode              As Integer
End Type

Private Type DEBUG_EXIT_THREAD_DEBUG_INFO
   Header As DEBUG_EVENT_HEADER
   dwExitCode As Long
End Type

Private Type DEBUG_EXIT_PROCESS_DEBUG_INFO
   Header As DEBUG_EVENT_HEADER
   dwExitCode As Long
End Type

Private Type DEBUG_LOAD_DLL_DEBUG_INFO
   Header As DEBUG_EVENT_HEADER
   hfile As Long
   lpBaseOfDll As Long
   dwDebugInfoFileOffset As Long
   nDebugInfoSize As Long
   lpImageName As Long
   fUnicode As Integer
End Type

Private Type DEBUG_UNLOAD_DLL_DEBUG_INFO
   Header As DEBUG_EVENT_HEADER
   lpBaseOfDll As Long
End Type

Private Type DEBUG_OUTPUT_DEBUG_STRING_INFO
   Header As DEBUG_EVENT_HEADER
   lpDebugStringData As Long
   fUnicode As Integer
   nDebugStringLength As Integer
End Type

Private Type DEBUG_RIP_INFO
    Header As DEBUG_EVENT_HEADER
    dwError As Long
    dwType As Long
End Type

Public Enum DebugStates
   DBG_CONTINUE = &H10002
   DBG_TERMINATE_THREAD = &H40010003
   DBG_TERMINATE_PROCESS = &H40010004
   DBG_CONTROL_C = &H40010005
   DBG_CONTROL_BREAK = &H40010008
   DBG_EXCEPTION_NOT_HANDLED = &H80010001
End Enum

Private Declare Function ContinueDebugEvent Lib "kernel32" _
(ByVal dwProcessId As Long, _
 ByVal dwThreadId As Long, _
 ByVal dwContinueStatus As DebugStates) As Long


'DWORD64 SymLoadModule64(
'  HANDLE hProcess,
'  HANDLE hFile,
'  PSTR ImageName,
'  PSTR ModuleName,
'  DWORD64 BaseOfDll,
'  DWORD SizeOfDll);
Public Declare Function SymLoadModule Lib "dbghelp.dll" ( _
       ByVal bProcess As Long, _
       ByVal hfile As Long, _
       ByVal ImageName As String, _
       ByVal ModuleName As String, _
       ByVal BaseOfDll As Long, _
       ByVal SizeOfDll As Long) As Long
       
'BOOL SymUnloadModule64(
'  HANDLE hProcess,
'  DWORD64 BaseOfDll);
Public Declare Function SymUnloadModule Lib "dbghelp.dll" ( _
        ByVal hProcess As Long, _
        ByVal BaseOfDll As Long) As Long

'BOOL SymInitialize(
'  HANDLE hProcess,
'  PSTR UserSearchPath,
'  BOOL fInvadeProcess
');
Public Declare Function SymInitialize Lib "dbghelp.dll" ( _
     ByVal hProcess As Long, _
     ByVal UserSearchPath As String, _
     ByVal fInvadeProcess As Boolean) As Long
     
'BOOL SymCleanup(
'  Handle hProcess
');
Public Declare Function SymCleanup Lib "dbghelp.dll" (ByVal hProcess As Long) As Long


