VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDisassembler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com

Private Type t_Disasm  '         // Results of disassembling
  ip As Long  '                  // Instrucion pointer
  dump As String * 256        '  // Hexadecimal dump of the command
  result As String * 256
  comment As String * 256     '  // Brief comment
  cmdtype As Long '              // One of C_xxx
  memtype As Long '              // Type of addressed variable in memory
  nprefix As Long '              // Number of prefixes
  indexed As Long '              // Address contains register(s)
  jmpconst As Long '             // Constant jump address
  jmptable As Long '             // Possible address of switch table
  adrconst As Long '             // Constant part of address
  immconst As Long '             // Immediate constant
  zeroconst As Long '            // Whether contains zero constant
  fixupoffset As Long '          // Possible offset of 32-bit fixups
  fixupsize As Long '            // Possible total size of fixups or 0
  error As Long '                // Error while disassembling command
  warnings As Long '             // Combination of DAW_xxx
End Type

Enum disasmMode
    DISASM_SIZE = 0     '            // Determine command size only
    DISASM_DATA = 1     '            // Determine size and analysis data
    DISASM_FILE = 3     '            // Disassembly, no symbols
    DISASM_CODE = 4     '            // Full disassembly
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Private Declare Function disasm Lib "olly.dll" Alias "Disasm" ( _
        ByRef src As Byte, ByVal srcsize As Long, ByVal ip As Long, _
        disasm As t_Disasm, dMode As disasmMode) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Sub Class_Initialize()
    'Make sure olly.dll is found regardless of the current working directory
    '(in the IDE the CWD is VB6's folder, not the project/app folder).
    On Error Resume Next
    If LoadLibraryA(App.Path & "\olly.dll") = 0 Then LoadLibraryA "olly.dll"
End Sub


Function DisasmBytes(b() As Byte, va As Long) As CInstruction
    Dim da As t_Disasm
    Dim x As Long
    Dim src As String
    Dim dump As String
    Dim inst As String
    
    Dim ret As New CInstruction
    
    x = disasm(b(0), UBound(b) + 1, va, da, DISASM_CODE)
    
    ret.instLen = x
    ret.offset = Hex(va)
 


        'MsgBox GetNativeApi(&H401120)

    'Offset Phys: 4384
'4198688
'401120

    
    inst = da.result
    dump = da.dump
    
    x = InStr(inst, Chr(0)) - 1
    ret.command = Mid$(inst, 1, x)

        'If Left$(da.result, "5") = "CALL " Then
        
        'Debug.Print da.result
        'MsgBox ConvertHex("401120")

        'Right$(da.result, Len(da.result) - 7)

        'MsgBox ConvertHex(Right$(da.result, Len(da.result) - 5))
        'ret.command = ret.command & " '" & GetNativeApi(ConvertHex(Mid$(da.result, 8, 6)))
        'End If
    
    x = InStr(dump, Chr(0)) - 1
    ret.dump = Mid$(dump, 1, x)
    'If x < 30 Then dump = dump & Space(30 - x)

    'Capture the analysis data olly.dll already computed (previously discarded).
    'This is what the native->VB engine works from: instruction class, branch
    'and call targets, address displacements, immediates and relocation info.
    ret.va = va
    ret.cmdType = da.cmdtype
    ret.jmpConst = da.jmpconst
    ret.jmpTable = da.jmptable
    ret.adrConst = da.adrconst
    ret.immConst = da.immconst
    ret.fixupOffset = da.fixupoffset
    ret.fixupSize = da.fixupsize
    ret.memType = da.memtype

    Set DisasmBytes = ret

End Function

'Olly command-type classes (C_xxx) are declared Public in modNativeToVB so they
'are visible here and to the native->VB engine.

Public Function ClassOf(inst As CInstruction) As Long
    ClassOf = inst.cmdType And C_TYPEMASK
End Function

Public Function DisasmProc(b() As Byte, ByVal va As Long, ByVal maxLen As Long) As Collection
'*****************************
'Disassemble a procedure linearly from its entry, stopping after the final
'RET of the body (tracks the furthest forward branch target so the function's
'tail is not cut off at the first RET, the way DisasmBlock was).
'*****************************
    Dim tmp() As Byte, pointer As Long, d As CInstruction
    Dim ret As New Collection, cls As Long, farTarget As Long, started As Boolean
    On Error GoTo hell
    pointer = LBound(b)
    farTarget = va
    started = True
    Do While pointer < UBound(b) + 1 And pointer < maxLen
        ReDim tmp(UBound(b) - pointer + 1)
        CopyMemory tmp(0), b(pointer), UBound(tmp)
        Set d = DisasmBytes(tmp, va)
        If d.instLen <= 0 Then Exit Do
        ret.Add d
        cls = d.cmdType And C_TYPEMASK
        'Track the furthest forward branch target inside the proc
        If (cls = C_JMP Or cls = C_JMC) And d.jmpConst > farTarget Then farTarget = d.jmpConst
        va = va + d.instLen
        pointer = pointer + d.instLen
        'Stop once we hit a RET that is at or past every forward branch target
        If cls = C_RET And va > farTarget Then Exit Do
    Loop
hell:
    Set DisasmProc = ret
End Function

Function DisasmBlock(b() As Byte, ByVal va As Long) As Collection
    Dim tmp() As Byte
    Dim disasmLen As Long
    Dim pointer As Long
    Dim inst As String
    Dim dump As String
    
    Dim d As CInstruction
    Dim ret As New Collection
    
    On Error GoTo hell
    
    pointer = LBound(b)
     
    While pointer < UBound(b) + 1
    
        ReDim tmp(UBound(b) - pointer + 1)
        CopyMemory tmp(0), b(pointer), UBound(tmp)
        
        Set d = DisasmBytes(tmp, va)
        'Guard against an undecodable byte (instLen <= 0) which would otherwise
        'spin this loop forever and hang the UI.
        If d.instLen <= 0 Then GoTo hell
        ret.Add d

        va = va + d.instLen
        pointer = pointer + d.instLen

    Wend
    
hell:
    Set DisasmBlock = ret

 
End Function
Public Function GetNativeApi(t As Long)
    'Resolve the runtime/imported API targeted by an indirect call/mov whose
    'opcode is "FF15"/"8B1D" style (operand = absolute IAT address at t+2).
    'Reads the IAT pointer with bounds checks, then defers to GetApiByIatVa.
    Dim a1 As Long, pos As Long, F As Long
    On Error Resume Next
    F = FreeFile
    Open SFilePath For Binary Access Read As F
    pos = t + 3 - OptHeader.ImageBase           '1-based -> file byte at t+2
    If pos >= 1 And pos <= LOF(F) - 3 Then Get F, pos, a1
    Close F
    GetNativeApi = GetApiByIatVa(a1)
End Function

Public Function GetApiByIatVa(ByVal iatVa As Long) As String
    'Given the virtual address of an IAT slot, return the name of the runtime
    'helper / imported API it points at.  Fully bounds-checked so a bad or zero
    'address simply yields "" instead of a 'Bad record number' error.
    Dim getImportRva As Long, J As Long, F As Long, pos As Long
    Dim u As String, TDs As String, ouR As String
    On Error Resume Next
    If iatVa < OptHeader.ImageBase Then Exit Function
    F = FreeFile
    Open SFilePath For Binary Access Read As F
    pos = iatVa + 1 - OptHeader.ImageBase
    If pos < 1 Or pos > LOF(F) - 3 Then Close F: Exit Function
    Get #F, pos, getImportRva
    Close F
    If getImportRva = 0 Then Exit Function

    For J = 1 To UBound(exeIMPORT_APINAME)
        If exeIMPORT_APINAME(J).Address = getImportRva Then
            If Left$(exeIMPORT_APINAME(J).ApiName, 8) = "!ordinal" Then
                u = VBFunction_Description(Val(Mid$(exeIMPORT_APINAME(J).ApiName, 12)), vbNullString, ouR)
            Else
                u = exeIMPORT_APINAME(J).ApiName
            End If
            GetApiByIatVa = u
            Exit Function
        End If
    Next
End Function

 
