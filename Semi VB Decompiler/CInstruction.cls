VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com

Public dump As String
Public offset As String
Public command As String
Public instLen As Long

'--- Analysis data captured from olly.dll t_Disasm (used by the native->VB engine) ---
Public va As Long           'instruction address (numeric, = Val of offset)
Public cmdType As Long      'olly C_xxx class (high nibble: call/jmp/jcc/ret/...)
Public jmpConst As Long     'absolute target of a direct jmp/jcc/call (0 if indirect)
Public jmpTable As Long     'possible switch-table address
Public adrConst As Long     'constant part of a memory address/displacement (e.g. [reg+adrConst])
Public immConst As Long     'immediate constant operand
Public fixupOffset As Long  'offset of a 32-bit relocation within the instruction (0 if none)
Public fixupSize As Long    'size of the relocation, >0 means the address operand is an absolute/fixed-up address (IAT, global)
Public memType As Long      'type of addressed memory variable
