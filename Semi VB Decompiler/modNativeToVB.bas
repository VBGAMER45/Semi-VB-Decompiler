Attribute VB_Name = "modNativeToVB"
'#############################################
'modNativeToVB - Native (machine code) -> VB reconstruction
'Semi VB Decompiler
'
'VB6 native EXEs still drive every object/string/variant operation through
'msvbvm60 runtime helpers and through the form/control vtables, so the machine
'code is highly idiomatic.  This module disassembles a procedure (via the
'olly.dll engine in CDisassembler, which now exposes branch/call/displacement
'analysis) and folds the recognised idioms back into readable VB:
'
'   call [reg + bigOffset]   -> a form control accessor   (Form1.car5)
'   call [obj + smallOffset] -> a control property get/let (.Left)
'   call [IAT]               -> an msvbvm60 runtime helper (__vbaObjSet, RGB...)
'
'Stage 3 adds a light data-flow model so whole expressions are rebuilt:
'   * an FPU expression stack    (fld / fadd / fsub / fmul / fdiv / fstp)
'   * a map of [ebp-X] local slots to the expression they currently hold
'   * a property GET writes its result into the local addressed just before it
'so a following property LET reads e.g. "Form1.car5.Left = (Form1.car5.Left + 125)".
'
'All operand information is taken from the olly t_Disasm analysis fields
'(adrConst, immConst, fixupSize) plus the instruction mnemonic keyword, so the
'engine does not depend on the exact mnemonic text formatting.
'
'Control offsets reuse the P-Code instance map (base + tControl.index*4) and
'property names resolve through modPCode.GetProperty.  Anything not recognised
'is emitted as annotated assembly so no code is lost.
'#############################################
Option Explicit

'Olly command-type classes (high nibble of t_Disasm.cmdtype); see olly disasm.h.
'Declared here in a standard module so the constants are globally visible to
'both CDisassembler and this engine.
Public Enum eCmdType
    C_TYPEMASK = &HF0
    C_CMD = &H0      'ordinary instruction
    C_PSH = &H10     'push
    C_POP = &H20     'pop
    C_MMX = &H30
    C_FLT = &H40     'FPU instruction
    C_JMP = &H50     'unconditional jump
    C_JMC = &H60     'conditional jump
    C_CAL = &H70     'call
    C_RET = &H80     'return
    C_FLG = &H90
    C_RTF = &HA0
    C_REP = &HB0
    C_PRI = &HC0
    C_DAT = &HD0
    C_NOW = &HE0
    C_BAD = &HF0     'unrecognized command
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Private dsmNative As CDisassembler
'Raw disassembly text of the proc most recently passed to
'DecompileNativeProcToVB - built from the same DisasmProc result so the Dism
'tab can be cached without a second disassembly pass.
Public NVLastDisasmText As String

'--- Per-procedure data-flow state ---
Private NVForm As String          'owning form/object (for control resolution)
Private NVHasMe As Boolean        'current proc is a class/form method (receives Me at ebp+8), not a .bas module sub
Private NVBase As Long            'solved control-block base for this form (-1 = unknown)
Private NVLastControl As String   'most recently accessed control name
Private NVLastGuid As String      'its GUID (for property name resolution)
Private NVFpu() As String         'FPU expression stack
Private NVFpuTop As Long
Private NVLocal As Collection      'local stack slot (disp) -> expression
Private NVLocalGuid As Collection  'local stack slot (disp) -> control GUID, when the slot holds a control object
Private NVLastLea As Long          'displacement of the most recent LEA (GET out-param)
Private NVLastLeaSet As Boolean
Private NVPendingArg As String     'value fstp'd into the outgoing argument area
Private NVLastImm As String        'most recent pushed immediate (decoded)
Private NVRegImport(7) As String   'register -> runtime helper cached into it
Private NVPushImm() As String      'recent pushed immediates (call argument list)
Private NVPushDisp() As Long       'for each pushed arg, the by-ref local displacement it addresses (0 = not by-ref)
Private NVPushTop As Long
Private NVLastPushDisp As Long     'set by NativePushOperand: by-ref local disp of the value just decoded (0 = none)
Private NVVSlot As Collection      'variant stack slot (disp) -> last value stored (VT tag / string / expr)
Private NVIndent As Long           'current block-indent level
Private NVIfTarget() As Long       'addresses where open If blocks must close
Private NVIfTop As Long
'--- Select Case (jump-table) reconstruction ---
Private NVSelExprReg As Collection  '"S"&dispatchVA -> index register (the Select expression) - emit "Select Case <reg>"
Private NVSelCaseVal As Collection  '"S"&caseVA -> case value(s) CSV, or "Else" - emit "Case ..."
Private NVSelEnd As Collection      '"S"&endVA -> "1" - emit "End Select"
Private NVSelSkip As Collection     '"S"&va -> "1" - suppress this instruction (the bound Jcc / case-end jumps)
'Open-Select stack (handles multiple / nested Selects per proc). NVSelTop = depth
'(0 = none, index of innermost). NVSelStkBase(level) = the indent level at which that
'Select's "Select Case" line was emitted; Case arms emit at base+1, bodies at base+2,
'End Select restores NVIndent to base. Absolute baselines need no per-arm caseOpen flag.
Private NVSelStkBase() As Long
Private NVSelTop As Long
'--- For loop (counted induction variable) reconstruction ---
Private NVForHdr As Collection     '"F"&headerVA -> counter register index (0..7) - emit "For <var> = <start> To <limit>"
Private NVForJmp As Collection     '"F"&backedgeVA -> headerVA (string) - emit "Next <var>"
Private NVForSkip As Collection    '"F"&va -> "1" - suppress this instruction (the exit Jcc)
Private NVForName As Collection    '"F"&headerVA -> the loop variable name assigned at emit time
Private NVForCnt As Long           'per-proc loop-variable name allocation index
'--- Floating-point comparison idiom (fcom/fnstsw/test ah,mask/jcc -> boolean) ---
Private NVFpCmp As Collection      '"P"&testVA -> "<relOp>|<regIdx>" - set NVReg(reg) to the relational
Private NVFpSkip As Collection     '"P"&va -> "1" - suppress this scaffolding instruction
Private NVStrCmpReg As Collection  '"P"&strcmpCallVA -> regIdx - the StrCmp result boolean-materializes into this register
Private NVCounterSlot As Collection '"C"&disp -> "1" - a stack slot that is a loop induction variable (render by name, not a stale value)
Private NVWhileCond As Collection  '"W"&exitJccVA -> "1" - emit `Do While <cond>` here (top-tested loop header)
Private NVWhileLoop As Collection  '"W"&backedgeVA -> "1" - emit `Loop` here (the back-edge of a Do While)
Private NVArgTok() As String       'per-proc: generic tokens (arg_<offset>) to replace with...
Private NVArgNm() As String        '...their recovered parameter names, at proc finalisation
Private NVArgN As Long             'count of recovered parameter-name substitutions this proc
Private NVLastCmp As String        'hint expression for the next If condition
Private NVStrLits As Collection    'pending string literals (e.g. MsgBox arguments)
Private NVSkipLabels As Collection 'branch targets that belong to dropped error-check guards
Private NVReg(7) As String         'symbolic value currently held in each GP register (eax..edi)
Private NVRegIsAddr(7) As Boolean  'True when a register holds &local (from LEA), for by-ref pushes
Private NVRegAddr(7) As String     'the local name a register's LEA-address points to
Private NVRegAddrDisp(7) As Long   'the local DISPLACEMENT a register's LEA-address points to (for variant resolution)
Private NVRegIsMe(7) As Boolean    'True when a register holds an object pointer (this/Me or a module global)
Private NVRegIsFormVt(7) As Boolean 'True when a register holds an object's vtable ([objPtr])
Private NVRegObjType(7) As String  'object NAME a register's POINTER refers to (App/Screen/Clipboard, or a control like Form1.File1)
Private NVRegObjVt(7) As String    'object NAME whose VTABLE a register holds ([objPtr] deref)
Private NVRegObjGuid(7) As String  'control GUID a register's object POINTER refers to ("" = intrinsic global / none)
Private NVRegObjVtGuid(7) As String 'control GUID whose VTABLE a register holds (resolves control props via GetProperty)
Private NVRegObjInst(7) As String  'receiver expression for a register holding a user-class VTABLE (e.g. global_004230F4), so call [vt+off] -> recv.Method
Private NVObjClass As Collection   'key "G"&globalVA -> user class name of the object instance stored at that global (typed at the __vbaNew auto-instantiation)
Private NVRecentPush(7) As Long    'ring of recent `push imm32/imm8` raw values (to recover __vbaNew's Object Info + @global args)
Private NVRecentTop As Long
Private NVLoopHdr As Collection    'addresses that are loop headers (back-edge targets)
Private NVCallHandled As Boolean   'set by NativeRuntimeCall: True when the call was recognised
Private NVErrHandler As Long       'address of this procedure's On Error handler block (0 = none)
Private NVProcEndWord As String    'closing keyword for this proc: "Sub" / "Function" / "Property"
Private NVRetN As Long             'the proc's `ret imm16` operand (callee-popped arg bytes), -1 if none
Private NVApiStubCache As Collection 'declared-DLL stub address -> resolved API name (global, "" = not a stub)
Private NVCmpL As String           'pending condition: left operand (symbolic)
Private NVCmpR As String           'pending condition: right operand (symbolic)
Private NVCmpIsTest As Boolean     'the pending compare came from TEST (zero-compare)
Private NVCmpIsBool As Boolean     'NVCmpL is already a relational Boolean (render it directly, no "<> 0")
Private NVCmpSet As Boolean        'a GP TEST/CMP condition hint is pending
Private NVFpuChk As Boolean        'an FPU status-word check (fnstsw;test al,imm) is pending -> drop its Jcc
Private NVPendingCall As String    'a "Call X()" deferred until we know if its result is used
Private NVErrObjPending As Boolean  'set by rtcErrObj: the next __vbaObjSet stores the Err object (re-tag eax)
Private NVKeepPushStack As Boolean  'set by a runtime-call handler that manages NVPushTop itself (concat chain) so the dispatch does not wipe the remaining arguments

'Sentinel for a "missing optional argument" Variant (VT_ERROR / DISP_E_PARAMNOTFOUND),
'used while reconstructing rtcMsgBox / rtcInputBox by-reference Variant argument lists.
Private Const NV_MISSING As String = "<<MISSING>>"

Public Function DecompileNativeProcToVB(ByVal addr As Long) As String
'*****************************
'Disassemble and partially decompile one native procedure at memory 'addr'.
'*****************************
    Dim b() As Byte, col As Collection, inst As CInstruction
    Dim output As String, fp As Integer
    Dim labels As Collection
    On Error GoTo fail

    If dsmNative Is Nothing Then Set dsmNative = New CDisassembler
    'Build the As-New private-field class map once, program-wide, so method calls
    'on Me.<clsBitmap field> resolve regardless of which proc auto-instantiated it.
    If gFormFieldClass Is Nothing Then NativeScanFormFieldClasses

    'The procedure list hands back an address a few bytes into the prologue
    '(the SEH setup), so snap back to the real "push ebp / mov ebp,esp" entry.
    addr = NativeSnapEntry(addr)

    'Read up to 8 KB of the procedure from the image
    ReDim b(8191)
    fp = FreeFile
    Open SFilePath For Binary Access Read As #fp
        Get #fp, addr + 1 - OptHeader.ImageBase, b
    Close #fp

    Set col = dsmNative.DisasmProc(b, addr, 8192)
    If col Is Nothing Then DecompileNativeProcToVB = "' (could not disassemble " & Hex$(addr) & ")": Exit Function

    'Build the raw disassembly listing from the SAME collection (no second
    'disassembly) so the Dism tab can be served from cache later.
    NVLastDisasmText = ""
    For Each inst In col
        NVLastDisasmText = NVLastDisasmText & inst.offset & "  " & inst.dump & "  " & inst.command & vbCrLf
    Next


    'Reset per-proc state
    NVForm = NativeFormOf(addr)
    NVHasMe = NativeProcHasMe(addr)
    NVProcEndWord = "Sub"
    NVLastControl = "": NVLastGuid = "": NVLastImm = "": NVPendingArg = ""
    NVLastLea = 0: NVLastLeaSet = False: NVLastCmp = ""
    NVCmpSet = False: NVCmpL = "": NVCmpR = "": NVCmpIsTest = False: NVCmpIsBool = False: NVFpuChk = False
    NVPendingCall = "": NVErrObjPending = False
    ReDim NVFpu(31): NVFpuTop = 0
    ReDim NVPushImm(31): ReDim NVPushDisp(31): NVPushTop = 0: NVLastPushDisp = 0
    ReDim NVIfTarget(31): NVIfTop = 0: NVIndent = 0
    Dim r As Long
    For r = 0 To 7: NVReg(r) = "": NVRegIsAddr(r) = False: NVRegAddr(r) = "": NVRegAddrDisp(r) = 0: NVRegIsMe(r) = False: NVRegIsFormVt(r) = False: NVRegObjType(r) = "": NVRegObjVt(r) = "": NVRegObjGuid(r) = "": NVRegObjVtGuid(r) = "": NVRegObjInst(r) = "": Next
    Set NVObjClass = New Collection
    For r = 0 To 7: NVRecentPush(r) = 0: Next
    NVRecentTop = 0
    Set NVLocal = New Collection
    Set NVLocalGuid = New Collection
    Set NVStrLits = New Collection
    Set NVVSlot = New Collection
    Set NVSkipLabels = New Collection
    Set NVLoopHdr = New Collection
    Set NVSelExprReg = New Collection
    Set NVSelCaseVal = New Collection
    Set NVSelEnd = New Collection
    Set NVSelSkip = New Collection
    Set NVForHdr = New Collection
    Set NVForJmp = New Collection
    Set NVForSkip = New Collection
    Set NVForName = New Collection
    NVForCnt = 0
    Set NVFpCmp = New Collection
    Set NVFpSkip = New Collection
    Set NVStrCmpReg = New Collection
    Set NVCounterSlot = New Collection
    Set NVWhileCond = New Collection
    Set NVWhileLoop = New Collection
    NVArgN = 0
    ReDim NVSelStkBase(31): NVSelTop = 0
    NVBase = NativeSolveControlBase(col)

    'Collect branch targets (for labels), and detect the On Error handler:
    'VB's error epilogue is "PUSH <resume-addr> ; JMP <exit>" and the handler
    'block starts at the instruction immediately after that JMP.
    Set labels = New Collection
    NVErrHandler = 0
    NVRetN = -1
    Dim prevWasResumePush As Boolean
    prevWasResumePush = False
    For Each inst In col
        Dim cls As Long
        cls = inst.cmdType And C_TYPEMASK
        If (cls = C_JMP Or cls = C_JMC) And inst.jmpConst <> 0 Then
            NativeAddUnique labels, inst.jmpConst
        End If
        If cls = C_JMP And prevWasResumePush And NVErrHandler = 0 Then
            NVErrHandler = inst.va + inst.instLen        'handler = fall-through of the JMP
        End If
        'A `ret imm16` (C2) pops the callee's argument bytes - the parameter count.
        'Take the FIRST one: the disassembly window runs past this proc into the
        'next, so a later ret belongs to a different procedure.  All of one proc's
        'exits pop the same N, so the first ret is this proc's.
        If cls = C_RET And NVRetN = -1 Then
            Dim rdmp As String
            rdmp = UCase$(Replace(inst.dump, " ", ""))
            If Left$(rdmp, 2) = "C2" And Len(rdmp) >= 6 Then NVRetN = NativeDumpByte(rdmp, 1) + NativeDumpByte(rdmp, 2) * 256
        End If
        prevWasResumePush = NativeIsResumePush(inst, addr)
    Next
    If NVErrHandler <> 0 Then NativeAddUnique labels, NVErrHandler

    'Fallback for procs the linear disassembly mis-aligns on (e.g. SEH/On Error
    'handlers): byte-scan the proc's own range for its epilogue `ret N`.  The range
    'ends at the next discovered procedure; the last C2 before it (padding is not
    'C2) is the epilogue, and every exit pops the same N.
    If NVRetN = -1 Then
        Dim nextAddr As Long, pp As Long, scanLen As Long, jj As Long, nn As Long
        nextAddr = addr + 8190
        For pp = 0 To UBound(gNativeProcArray) - 1
            If gNativeProcArray(pp).offset > addr And gNativeProcArray(pp).offset < nextAddr Then nextAddr = gNativeProcArray(pp).offset
        Next
        scanLen = nextAddr - addr
        If scanLen > 8190 Then scanLen = 8190
        For jj = scanLen - 3 To 0 Step -1
            If b(jj) = &HC2 Then
                nn = b(jj + 1) + b(jj + 2) * 256&
                If nn >= 0 And nn <= 256 Then NVRetN = nn: Exit For
            End If
        Next
    End If

    'Detect Select Case jump tables so the dispatch + case bodies reconstruct as a
    'Select Case block rather than a bogus indirect GoTo.
    NativeDetectSelects col, b, addr

    'Detect counted For loops (mov ctr,start / cmp ctr,limit / jg exit / ... /
    'inc ctr / jmp header) so they reconstruct as For...Next with a named counter
    'rather than an If <cond> ... GoTo back-edge.
    NativeDetectForLoops col

    'Detect top-tested Do While loops (a header `cmp/jcc` exit guarding a body that
    'ends in an unconditional back-edge jmp) so they reconstruct as Do While <cond>
    '/ Loop rather than `loc: If <cond> ... GoTo loc / End If`.
    NativeDetectWhileLoops col

    'Detect floating-point comparison idioms (fcom/fnstsw/test ah,mask/jcc + the
    'boolean materialization) so they yield a real relational instead of <cond>.
    NativeDetectFpCompares col

    'Detect VB6's string-comparison relational: __vbaStrCmp(a,b) whose tri-state
    'result is boolean-materialised (neg/sbb/inc/neg) into a register, then tested.
    'Bind the equality relational to that register so the branch reads `a = b`
    'instead of a blank <cond>.
    NativeDetectStrCmpCompares col

    'Mark stack slots that are loop induction variables (written AND read inside a
    'backward-branch loop) so they render by their variable name rather than a stale
    'per-iteration value.  Loop-type-agnostic (For / Do While / Do Until / While).
    NativeDetectCounterSlots col

    'Suppress VB6's SAFEARRAY element-access bounds-check guards (the je/jne/jb
    'jumps around __vbaGenerateBoundsError) so an array access stops rendering as
    'bogus nested `If arr <> 0 / If arr = 1 / If (idx - lb) >= cEls` blocks.
    NativeDetectBoundsChecks col

    output = NativeProcHeader(addr) & vbCrLf

    'A Function/Property Get returns its value through a hidden retbuf pointer (the
    'last stack arg); the local copied into it at the epilogue is the return value.
    'Rename that local to the procedure name so assignments read `FuncName = value`.
    'Runs AFTER NativeProcHeader so it appends to the parameter-name substitutions.
    NativeDetectReturnSlot col, addr

    For Each inst In col
        'Resolve a deferred call (from the previous instruction) based on how
        'THIS instruction uses eax, the call's result.  Done before any block
        'close/label so a flushed "Call X()" stays inside the right block.
        If Len(NVPendingCall) > 0 Then
            Dim eu As Long
            eu = NativeEaxUse(inst)
            If eu = 1 Then
                NVPendingCall = ""                          'result consumed -> folded
            ElseIf eu = 2 Then
                output = output & NativeIndentStr() & NVPendingCall & vbCrLf
                NVPendingCall = "": NVReg(0) = ""           'unused -> emit; result spent
            End If
        End If
        'Close any structured If blocks that end at this address
        NativeCloseIfs output, inst.va

        '--- Select Case transitions (after If-close so case-body Ifs close first) ---
        Dim svaKey As String, selCV As String, selExpr As String
        svaKey = "S" & inst.va
        'End Select: close the innermost open Select, restoring its indent baseline.
        If NVSelTop > 0 And Len(NativeColGet(NVSelEnd, svaKey)) > 0 Then
            NVIndent = NVSelStkBase(NVSelTop)
            output = output & NativeIndentStr() & "End Select" & vbCrLf
            NVSelTop = NVSelTop - 1
        End If
        'Case / Case Else: a new arm of the innermost open Select.
        selCV = NativeColGet(NVSelCaseVal, svaKey)
        If NVSelTop > 0 And Len(selCV) > 0 Then
            NVIndent = NVSelStkBase(NVSelTop) + 1
            If selCV = "Else" Then
                output = output & NativeIndentStr() & "Case Else" & vbCrLf
            Else
                output = output & NativeIndentStr() & "Case " & selCV & vbCrLf
            End If
            NVIndent = NVSelStkBase(NVSelTop) + 2
        End If
        'Select Case open: push a new block onto the stack.
        selExpr = NativeColGet(NVSelExprReg, svaKey)
        If Len(selExpr) > 0 Then
            output = output & NativeIndentStr() & "Select Case " & NativeRegVal(CLng(selExpr)) & vbCrLf
            NVSelTop = NVSelTop + 1
            If NVSelTop > UBound(NVSelStkBase) Then ReDim Preserve NVSelStkBase(NVSelTop + 8)
            NVSelStkBase(NVSelTop) = NVIndent
            NVIndent = NVIndent + 1
            NVCmpSet = False                'the bound cmp's flags fed the suppressed bound jump
        End If
        'Suppress the dispatch jump / bound Jcc / case-end jumps (replaced above).
        If Len(NativeColGet(NVSelSkip, svaKey)) > 0 Then GoTo nextInst

        '--- Floating-point comparison idiom ---
        'At the `test ah,mask` head: bind the recovered relational into the target
        'register (using the fcom operands captured in NVLastCmp) and cancel the
        'overflow-check flag the preceding fnstsw set.  The je/mov/jmp/xor/neg
        'scaffolding is suppressed via NVFpSkip.
        Dim fpRec As String
        fpRec = NativeColGet(NVFpCmp, "P" & inst.va)
        If Len(fpRec) > 0 Then
            Dim fpBar As Long, fpOp As String, fpReg As Long, fpExpr As String
            fpBar = InStr(fpRec, "|")
            fpOp = Left$(fpRec, fpBar - 1)
            fpReg = CLng(Mid$(fpRec, fpBar + 1))
            If Len(NVLastCmp) > 0 Then
                fpExpr = Replace(NVLastCmp, " ? ", " " & fpOp & " ")
            Else
                fpExpr = "(st0 " & fpOp & " 0)"
            End If
            NVReg(fpReg) = fpExpr
            NVRegIsAddr(fpReg) = False: NVRegIsMe(fpReg) = False: NVRegIsFormVt(fpReg) = False
            NVRegObjType(fpReg) = "": NVRegObjVt(fpReg) = "": NVRegObjGuid(fpReg) = "": NVRegObjVtGuid(fpReg) = ""
            NVFpuChk = False: NVLastCmp = ""
            GoTo nextInst
        End If
        'Suppress the fp-compare / bounds-check scaffolding.  Still emit the label
        'when this address is a real jump target (e.g. an On Error handler that lands
        'on a suppressed bounds-error call) so the branch to it does not dangle; a
        'label nothing references is removed later by the orphan-label strip.
        If Len(NativeColGet(NVFpSkip, "P" & inst.va)) > 0 Then
            If NativeHas(labels, inst.va) And Not NativeIsIfTarget(inst.va) _
               And Not NativeHas(NVSkipLabels, inst.va) And Not NativeHas(NVLoopHdr, inst.va) Then
                output = output & "loc_" & Right$("00000000" & Hex$(inst.va), 8) & ":" & vbCrLf
            End If
            GoTo nextInst
        End If

        '--- For loop transitions ---
        'Suppress a recognised For loop's exit Jcc (the For header below replaces it).
        If Len(NativeColGet(NVForSkip, "F" & inst.va)) > 0 Then GoTo nextInst
        'Close a For loop at its back-edge: emit "Next <var>" instead of GoTo header.
        Dim fHdr As String
        fHdr = NativeColGet(NVForJmp, "F" & inst.va)
        If Len(fHdr) > 0 Then
            If NVIndent > 0 Then NVIndent = NVIndent - 1
            output = output & NativeIndentStr() & "Next " & NativeColGet(NVForName, "F" & fHdr) & vbCrLf
            GoTo nextInst
        End If

        'Open a Do loop when this address is the target of a back-edge
        If NativeHas(NVLoopHdr, inst.va) Then
            output = output & NativeIndentStr() & "Do" & vbCrLf
            NVIndent = NVIndent + 1
        End If
        'Open a For loop at its header `cmp ctr,limit`: emit For ctr = start To limit.
        Dim fCtrS As String
        fCtrS = NativeColGet(NVForHdr, "F" & inst.va)
        If Len(fCtrS) > 0 Then
            Dim fCtr As Long, fName As String, fStart As String, fLimit As String
            fCtr = CLng(fCtrS)
            NativeDecodeCompare inst, "CMP"          'fills NVCmpL (start=counter init) / NVCmpR (limit)
            fStart = NVCmpL: fLimit = NVCmpR
            NVCmpSet = False: NVCmpL = "": NVCmpR = ""
            If Len(fStart) = 0 Then fStart = NativeRegVal(fCtr)
            If Len(fStart) = 0 Then fStart = "1"
            If Len(fLimit) = 0 Then fLimit = "?"
            fName = NativeLoopVarName(NVForCnt): NVForCnt = NVForCnt + 1
            NativeColPut NVForName, "F" & inst.va, fName
            output = output & NativeIndentStr() & "For " & fName & " = " & fStart & " To " & fLimit & vbCrLf
            NVIndent = NVIndent + 1
            'The counter now reads as the loop variable everywhere in the body, so
            'index expressions (esi = i-1) and SAFEARRAY element compares resolve.
            NVReg(fCtr) = fName
            NVRegIsAddr(fCtr) = False: NVRegIsMe(fCtr) = False: NVRegIsFormVt(fCtr) = False
            NVRegObjType(fCtr) = "": NVRegObjVt(fCtr) = "": NVRegObjGuid(fCtr) = "": NVRegObjVtGuid(fCtr) = ""
            GoTo nextInst
        End If
        'A label is only needed when it is a real jump target (not an If close
        'point, a loop header, a dropped error-check guard target, or a Select
        'Case / Case / End Select point already emitted above)
        If NativeHas(labels, inst.va) And Not NativeIsIfTarget(inst.va) _
           And Not NativeHas(NVSkipLabels, inst.va) And Not NativeHas(NVLoopHdr, inst.va) _
           And Len(selCV) = 0 And Len(NativeColGet(NVSelEnd, svaKey)) = 0 Then
            output = output & "loc_" & Right$("00000000" & Hex$(inst.va), 8) & ":" & vbCrLf
        End If
        output = output & NativeProcessInst(inst)
        'A call clobbers the CALLER-SAVED registers (ecx, edx; eax holds the return
        'and is re-tagged by the call handler), so their tracked object/control
        'identity is invalid afterwards - drop it.  The CALLEE-SAVED registers
        '(ebx, esi, edi, ebp) are preserved across the call by the callee, and the
        'mov tracker already clears their identity on any in-proc reload, so their
        'tag stays valid - keeping it lets a control-property LET resolve when the
        'value is built by intervening helper calls (e.g.
        'Label1.Caption = "a" & vbCrLf & "b": ebx holds Label1's vtable across the
        '__vbaStrCat calls between the load and put_Caption).
        If (inst.cmdType And C_TYPEMASK) = C_CAL Then
            NVRegObjType(1) = "": NVRegObjVt(1) = "": NVRegObjGuid(1) = "": NVRegObjVtGuid(1) = ""   'ecx
            NVRegObjType(2) = "": NVRegObjVt(2) = "": NVRegObjGuid(2) = "": NVRegObjVtGuid(2) = ""   'edx
            'eax now holds the call's RETURN value, never a `lea`-captured &local,
            'so drop any stale address-of tag on it - else a later push of eax leaks
            'the by-reference local: a string built after `lea eax,[var_20]` (for
            '__vbaObjSet) was pushing var_20 instead of the computed value.  Only eax
            'is cleared here: ecx/edx by-ref pushes happen before their call, and
            'clearing them mis-rendered live by-ref argument locals.
            NVRegIsAddr(0) = False
        End If
nextInst:
    Next

    'Close any Select Cases left open (their End fell outside the decoded range).
    Do While NVSelTop > 0
        NVIndent = NVSelStkBase(NVSelTop)
        output = output & NativeIndentStr() & "End Select" & vbCrLf
        NVSelTop = NVSelTop - 1
    Loop
    'Any call still deferred is the proc's last statement - emit it.
    If Len(NVPendingCall) > 0 Then output = output & NativeIndentStr() & NVPendingCall & vbCrLf: NVPendingCall = ""
    NativeCloseIfs output, &H7FFFFFFF
    output = output & "End " & NVProcEndWord & vbCrLf
    DecompileNativeProcToVB = NativeStripOrphanLabels(NativeSubstituteArgNames(output))
    Exit Function
fail:
    DecompileNativeProcToVB = "' Error decompiling " & Hex$(addr) & ": " & Err.Description & vbCrLf
End Function

Private Function NativeStripOrphanLabels(ByVal src As String) As String
    'Remove `loc_XXXXXXXX:` label lines that nothing branches to.  VB labels are
    'referenced only by GoTo / GoSub / Resume / On Error GoTo (all contain
    '"GoTo loc_"/"GoSub loc_"/"Resume loc_"), so a label-definition line not named
    'by any of those is dead noise from dropped/restructured branches - safe to cut.
    Dim lines() As String, i As Long, ln As String, lt As String, refs As Collection
    Dim out As String, lbl As String
    On Error Resume Next
    Set refs = New Collection
    lines = Split(src, vbCrLf)
    For i = 0 To UBound(lines)
        ln = lines(i)
        NativeAddRefLabel refs, ln, "GoTo loc_"
        NativeAddRefLabel refs, ln, "GoSub loc_"
        NativeAddRefLabel refs, ln, "Resume loc_"
    Next
    For i = 0 To UBound(lines)
        lt = Trim$(lines(i))
        If Left$(lt, 4) = "loc_" And Right$(lt, 1) = ":" And InStr(lt, " ") = 0 Then
            lbl = Left$(lt, Len(lt) - 1)
            If Not NativeHasKey(refs, lbl) Then GoTo skipLine    'orphan label -> drop
        End If
        If Len(out) > 0 Then out = out & vbCrLf
        out = out & lines(i)
skipLine:
    Next
    NativeStripOrphanLabels = out
End Function

Private Sub NativeAddRefLabel(refs As Collection, ByVal ln As String, ByVal kw As String)
    'Add the label named after every occurrence of kw (e.g. "GoTo loc_") in ln.
    Dim p As Long, lbl As String
    On Error Resume Next
    p = InStr(ln, kw)
    Do While p > 0
        lbl = Mid$(ln, p + Len(kw) - 4, 12)        'kw ends with "loc_"; back up 4 to include it
        If Left$(lbl, 4) = "loc_" Then refs.Add 1, lbl
        p = InStr(p + 1, ln, kw)
    Loop
End Sub

Private Function NativeHasKey(c As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = c(key)
    NativeHasKey = (Err.Number = 0)
End Function

Private Function NativeColGet(c As Collection, ByVal key As String) As String
    On Error Resume Next
    NativeColGet = c(key)
End Function

Private Sub NativeColPut(c As Collection, ByVal key As String, ByVal v As String)
    On Error Resume Next
    c.Remove key
    On Error GoTo 0
    On Error Resume Next
    c.Add v, key
End Sub

Private Function NativeBDword(ByRef b() As Byte, ByVal addr As Long, ByVal va As Long) As Long
    'Read a little-endian dword from the proc byte buffer b() (loaded at addr).
    Dim o As Long
    o = va - addr
    On Error GoTo bad
    If o < 0 Or o + 3 > UBound(b) Then GoTo bad
    NativeBDword = b(o) + b(o + 1) * &H100& + b(o + 2) * &H10000 + b(o + 3) * &H1000000
    Exit Function
bad:
End Function

Private Sub NativeDetectSelects(col As Collection, ByRef b() As Byte, ByVal addr As Long)
    'Recognise VB6's Select Case jump table:
    '    mov reg,<expr> ; [sub reg,base] ; cmp reg,RANGE ; ja DEFAULT ; jmp [reg*4 + TBL]
    'TBL is RANGE+1 dword case-body addresses; each case body ends with jmp END; the
    'ja target is the Case Else body (which falls through to END).  Populate the
    'NVSel* maps so the main loop emits Select Case / Case / Case Else / End Select.
    On Error GoTo done
    Dim n As Long, k As Long, i As Long, inst As CInstruction
    n = col.Count
    If n < 3 Then Exit Sub
    Dim arr() As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next
    For k = 2 To n - 1
        Dim idxReg As Long, tbl As Long
        If Not NativeJmpTableInfo(arr(k), idxReg, tbl) Then GoTo nextk
        If (arr(k - 1).cmdType And C_TYPEMASK) <> C_JMC Then GoTo nextk   'a conditional bound jump precedes
        Dim defAddr As Long, rangeC As Long
        defAddr = arr(k - 1).jmpConst
        If Not NativeCmpRegImm(arr(k - 2), idxReg, rangeC) Then GoTo nextk
        If rangeC < 0 Or rangeC > 4096 Then GoTo nextk
        Dim baseV As Long
        baseV = 0
        If k >= 3 Then
            Dim sv As Long
            If NativeSubRegImm(arr(k - 3), idxReg, sv) Then baseV = sv
        End If
        'Find END = target of the first case-body jump after the dispatch.
        Dim endVA As Long
        endVA = 0
        For i = k + 1 To n - 1
            If (arr(i).cmdType And C_TYPEMASK) = C_JMP And arr(i).jmpConst <> 0 Then endVA = arr(i).jmpConst: Exit For
        Next i
        If endVA = 0 Then GoTo nextk
        'Read the jump table; require every entry inside the proc.
        Dim caseVA As Long, ok As Boolean, prevCV As String
        ok = True
        For i = 0 To rangeC
            caseVA = NativeBDword(b, addr, tbl + i * 4)
            If caseVA < addr Or caseVA >= NVProcEndApprox(addr) Then ok = False: Exit For
        Next i
        If Not ok Then GoTo nextk
        'Populate.  Same target at consecutive indices -> a value list (Case 0, 1, 2).
        NativeColPut NVSelExprReg, "S" & arr(k).va, CStr(idxReg)
        NativeColPut NVSelSkip, "S" & arr(k).va, "1"          'the dispatch jmp itself
        NativeColPut NVSelSkip, "S" & arr(k - 1).va, "1"      'the bound ja
        For i = 0 To rangeC
            caseVA = NativeBDword(b, addr, tbl + i * 4)
            If caseVA = defAddr Then GoTo nexti               'a gap case -> Else, skip
            prevCV = NativeColGet(NVSelCaseVal, "S" & caseVA)
            If Len(prevCV) > 0 Then
                NativeColPut NVSelCaseVal, "S" & caseVA, prevCV & ", " & CStr(baseV + i)
            Else
                NativeColPut NVSelCaseVal, "S" & caseVA, CStr(baseV + i)
            End If
nexti:
        Next i
        NativeColPut NVSelCaseVal, "S" & defAddr, "Else"
        NativeColPut NVSelEnd, "S" & endVA, "1"
        'Suppress the case-body jumps to END (between dispatch and END).
        For i = k + 1 To n - 1
            If arr(i).va >= endVA Then Exit For
            If (arr(i).cmdType And C_TYPEMASK) = C_JMP And arr(i).jmpConst = endVA Then
                NativeColPut NVSelSkip, "S" & arr(i).va, "1"
            End If
        Next i
nextk:
    Next k
done:
End Sub

Private Sub NativeDetectCounterSlots(col As Collection)
    'Mark ebp-relative stack slots that are loop induction variables: a slot WRITTEN
    'and READ within the span of a backward branch (any loop - For / Do While / Do
    'Until / While...Wend).  Such a slot holds a different value each iteration, so the
    'store handler renders it by name (var_X) with the assignment surfaced, instead of
    'binding a single stale value that then leaks into the loop header / body as a
    'constant (e.g. `If 1 <= 10` for a register counter mirrored to var_24).
    On Error GoTo done
    Dim n As Long, k As Long, bi As Long, i As Long, inst As CInstruction
    n = col.Count
    If n < 3 Then Exit Sub
    Dim arr() As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next
    For bi = 1 To n - 1
        Dim cls As Long
        cls = arr(bi).cmdType And C_TYPEMASK
        If (cls = C_JMP Or cls = C_JMC) And arr(bi).jmpConst <> 0 And arr(bi).jmpConst < arr(bi).va Then
            'A backward branch: the loop body spans [hdrVA, this branch].
            Dim hdrVA As Long, wrote As Collection, readd As Collection
            hdrVA = arr(bi).jmpConst
            Set wrote = New Collection: Set readd = New Collection
            For i = 0 To n - 1
                If arr(i).va >= hdrVA And arr(i).va <= arr(bi).va Then
                    Dim wd As Long, rd As Long
                    wd = NativeStackStoreDisp(arr(i))
                    If wd < 0 Then NativeColPut wrote, "S" & wd, "1"
                    rd = NativeStackLoadDisp(arr(i))
                    If rd < 0 Then NativeColPut readd, "S" & rd, "1"
                End If
            Next
            'A slot written AND read in the loop is an induction variable.
            Dim ki As Long
            For ki = 0 To n - 1
                If arr(ki).va >= hdrVA And arr(ki).va <= arr(bi).va Then
                    Dim sd As Long
                    sd = NativeStackStoreDisp(arr(ki))
                    If sd < 0 Then
                        If Len(NativeColGet(readd, "S" & sd)) > 0 Then NativeColPut NVCounterSlot, "C" & sd, "1"
                    End If
                End If
            Next
        End If
    Next
done:
End Sub

Private Function NativeStackStoreDisp(inst As CInstruction) As Long
    'ebp-relative displacement (negative) written by `mov [ebp-X], r32` (89 /r),
    '`inc [ebp-X]` (FF /0) or `add/sub [ebp-X], imm` (83 or 81 /0 /5); else 1.
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long, md As Long, reg As Long, disp As Long, isAbs As Boolean
    NativeStackStoreDisp = 1
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    modrm = NativeDumpByte(dump, i + 1)
    md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7
    If md = 3 Then Exit Function                        'register destination, not a slot
    Select Case op
        Case &H89                                        'mov r/m32, r32
        Case &HFF: If reg <> 0 Then Exit Function         'inc -> /0
        Case &H83, &H81: If reg <> 0 And reg <> 5 Then Exit Function   'add /0 or sub /5
        Case Else: Exit Function
    End Select
    If NativeDecodeDisp(dump, disp, isAbs) Then
        If Not isAbs And disp < 0 Then NativeStackStoreDisp = disp
    End If
End Function

Private Function NativeStackLoadDisp(inst As CInstruction) As Long
    'ebp-relative displacement (negative) READ by an instruction whose memory operand
    'is [ebp-X] and is a source (mov r32,[ebp-X] / cmp / add / or ... reg,[ebp-X]);
    'else 1.  A broad read test - any non-store memory reference to a stack slot.
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long, md As Long, disp As Long, isAbs As Boolean
    NativeStackLoadDisp = 1
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    modrm = NativeDumpByte(dump, i + 1)
    md = (modrm \ &H40) And 3
    If md = 3 Then Exit Function
    'Reading opcodes with a r32, r/m32 form (the r/m is the memory source).
    Select Case op
        Case &H8B, &H3B, &H03, &HB, &H23, &H2B, &H33, &H39, &H3D, &H83, &H81
        Case Else: Exit Function
    End Select
    If NativeDecodeDisp(dump, disp, isAbs) Then
        If Not isAbs And disp < 0 Then NativeStackLoadDisp = disp
    End If
End Function

Private Sub NativeDetectForLoops(col As Collection)
    'Recognise VB6's counted For loop:
    '    mov <ctr>,<start> ; H: cmp <ctr>,<limit> ; jg <exit> ; ... ; <ctr>=<ctr>+1 ; jmp H
    'Populate NVFor* so the main loop emits For <var> = <start> To <limit> / Next
    'in place of the If <cond> ... GoTo back-edge.  Strict: every check must pass
    'or the loop is left as-is (a missed loop is harmless; a false match is not).
    On Error GoTo done
    Dim n As Long, k As Long, hi As Long, bi As Long, i As Long, inst As CInstruction
    n = col.Count
    If n < 4 Then Exit Sub
    Dim arr() As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next
    For bi = 3 To n - 1
        'Back-edge: an unconditional jmp to an earlier header address.
        If (arr(bi).cmdType And C_TYPEMASK) <> C_JMP Then GoTo nextb
        Dim hdrVA As Long
        hdrVA = arr(bi).jmpConst
        If hdrVA = 0 Or hdrVA > arr(bi).va Then GoTo nextb
        If Len(NativeColGet(NVForHdr, "F" & hdrVA)) > 0 Then GoTo nextb   'header already claimed
        'Locate the header instruction in the array.
        hi = -1
        For i = 0 To bi - 1
            If arr(i).va = hdrVA Then hi = i: Exit For
        Next i
        If hi < 0 Or hi + 1 >= bi Then GoTo nextb
        'Header = `cmp <ctr>, <limit>` immediately followed by an exit Jcc that
        'leaves the loop forward (past the back-edge).
        Dim ctrReg As Long
        ctrReg = NativeForCounterReg(arr(hi))
        If ctrReg < 0 Then GoTo nextb
        If (arr(hi + 1).cmdType And C_TYPEMASK) <> C_JMC Then GoTo nextb
        If Not NativeIsExitJcc(NativeMnem(arr(hi + 1))) Then GoTo nextb
        If arr(hi + 1).jmpConst <= arr(bi).va Then GoTo nextb
        'The counter must be incremented within the body.
        If Not NativeForHasIncrement(arr, hi + 2, bi - 1, ctrReg) Then GoTo nextb
        'The counter must be initialised to a CONSTANT before the header (a real
        'For starts at 0/1/2...).  Rejects false matches where the "counter"
        'register was just loaded from memory, which produced degenerate output
        'like `For i = var_38 To var_38` or `For i = ebx To edi`.
        If Not NativeForStartIsConst(arr, hi, ctrReg) Then GoTo nextb
        'The header must be reached ONLY by this back-edge (a single, clean
        'counted loop - no extra continue-jumps that would need a second Next).
        Dim refs As Long
        refs = 0
        For i = 0 To n - 1
            If ((arr(i).cmdType And C_TYPEMASK) = C_JMP Or (arr(i).cmdType And C_TYPEMASK) = C_JMC) _
               And arr(i).jmpConst = hdrVA Then refs = refs + 1
        Next i
        If refs <> 1 Then GoTo nextb
        'Record.
        NativeColPut NVForHdr, "F" & hdrVA, CStr(ctrReg)
        NativeColPut NVForJmp, "F" & arr(bi).va, CStr(hdrVA)
        NativeColPut NVForSkip, "F" & arr(hi + 1).va, "1"     'the exit Jcc (replaced by For)
        NativeAddUnique NVSkipLabels, hdrVA                   'header label -> the For line
nextb:
    Next bi
done:
End Sub

Private Sub NativeDetectWhileLoops(col As Collection)
    'Reconstruct a top-tested loop as `Do While <cond> ... Loop`.  Shape:
    '    H: [setup] ; cmp/test <x>,<y> ; jcc EXIT ; ... body ... ; jmp H ; EXIT:
    'where the back-edge is an UNCONDITIONAL jmp to the header H and the exit Jcc
    'leaves the loop forward (past the back-edge).  This is the safe, general loop
    'form (the increment - if any - stays visible in the body), covering both counted
    'loops the strict For detector misses and genuine Do While / While...Wend loops.
    'Emit Do While at the exit Jcc, Loop at the back-edge, and suppress the header
    'label.  Strict single-back-edge match; a miss just leaves the GoTo form.
    On Error GoTo done
    Dim n As Long, k As Long, bi As Long, hi As Long, j As Long, i As Long, inst As CInstruction
    n = col.Count
    If n < 4 Then Exit Sub
    Dim arr() As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next
    For bi = 1 To n - 1
        'Unconditional back-edge to an earlier header.
        If (arr(bi).cmdType And C_TYPEMASK) <> C_JMP Then GoTo nextw
        Dim hdrVA As Long
        hdrVA = arr(bi).jmpConst
        If hdrVA = 0 Or hdrVA >= arr(bi).va Then GoTo nextw
        'Not already claimed by the For detector.
        If Len(NativeColGet(NVForHdr, "F" & hdrVA)) > 0 Then GoTo nextw
        'Single back-edge / entry: the header is targeted by exactly this one jump
        '(extra jumps to it would need multiple Loop ends - leave those as GoTo).
        Dim refs As Long
        refs = 0
        For i = 0 To n - 1
            If ((arr(i).cmdType And C_TYPEMASK) = C_JMP Or (arr(i).cmdType And C_TYPEMASK) = C_JMC) _
               And arr(i).jmpConst = hdrVA Then refs = refs + 1
        Next i
        If refs <> 1 Then GoTo nextw
        'Locate the header instruction.
        hi = -1
        For i = 0 To bi - 1
            If arr(i).va = hdrVA Then hi = i: Exit For
        Next i
        If hi < 0 Then GoTo nextw
        'Find the header's exit test: the FIRST cmp/test in the header, immediately
        'followed by a forward Jcc whose target is past the back-edge.
        For j = hi To bi - 2
            Dim mnj As String
            mnj = NativeMnem(arr(j))
            If mnj = "CMP" Or mnj = "TEST" Then
                If (arr(j + 1).cmdType And C_TYPEMASK) = C_JMC Then
                    'Not an already-consumed scaffolding Jcc (For exit / bounds / fp).
                    If Len(NativeColGet(NVForSkip, "F" & arr(j + 1).va)) = 0 _
                       And Len(NativeColGet(NVFpSkip, "P" & arr(j + 1).va)) = 0 Then
                        If arr(j + 1).jmpConst > arr(bi).va Then
                            NativeColPut NVWhileCond, "W" & arr(j + 1).va, "1"
                            NativeColPut NVWhileLoop, "W" & arr(bi).va, "1"
                            NativeAddUnique NVSkipLabels, hdrVA       'header label -> Do While line
                        End If
                    End If
                End If
                Exit For                                              'only the first compare heads the loop
            End If
        Next j
nextw:
    Next bi
done:
End Sub

Private Function NativeForCounterReg(inst As CInstruction) As Long
    'Counter register of a header `cmp <reg>, <limit>`: 3B (cmp r32,r/m32) -> reg
    'field; 83/81 (cmp r/m32,imm with r/m a register) -> rm; 3D -> eax.  A 16-bit
    '(0x66) compare juggles low-word partials we cannot model, so it is rejected.
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long, md As Long
    NativeForCounterReg = -1
    If NativeMnem(inst) <> "CMP" Then Exit Function
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    If NativeHas66(dump) Then Exit Function
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    Select Case op
        Case &H3D                       'cmp eax, imm32
            NativeForCounterReg = 0
        Case &H3B                       'cmp r32, r/m32 -> counter is the reg field
            modrm = NativeDumpByte(dump, i + 1)
            NativeForCounterReg = (modrm \ 8) And 7
        Case &H83, &H81                 'cmp r/m32, imm -> counter is the r/m (register only)
            modrm = NativeDumpByte(dump, i + 1)
            md = (modrm \ &H40) And 3
            If md = 3 And ((modrm \ 8) And 7) = 7 Then NativeForCounterReg = modrm And 7
    End Select
End Function

Private Function NativeIsExitJcc(ByVal mn As String) As Boolean
    'A counted For loop exits when the counter exceeds its limit: jg/jge (signed)
    'or ja/jae (unsigned) and their synonyms.
    Select Case mn
        Case "JG", "JNLE", "JGE", "JNL", "JA", "JNBE", "JAE", "JNB", "JNC": NativeIsExitJcc = True
    End Select
End Function

Private Function NativeForHasIncrement(arr() As CInstruction, ByVal lo As Long, ByVal hi As Long, ByVal ctrReg As Long) As Boolean
    'True when the counter register is incremented within arr(lo..hi): a direct
    'inc/add, or VB6's three-step `mov tmp,k ; add tmp,ctr ; mov ctr,tmp` form
    '(tracked via a per-register "derived from counter" flag).
    Dim i As Long, dump As String, nn As Long, p As Long, op As Long, modrm As Long, md As Long, rg As Long, rm As Long
    Dim derived(7) As Boolean
    On Error Resume Next
    For i = lo To hi
        dump = Replace(arr(i).dump, " ", "")
        nn = Len(dump) \ 2
        If nn < 1 Then GoTo nexti
        p = NativeOpStart(dump, nn)
        op = NativeDumpByte(dump, p)
        If op = (&H40 + ctrReg) Then NativeForHasIncrement = True: Exit Function   'inc ctrReg (1-byte)
        modrm = NativeDumpByte(dump, p + 1)
        md = (modrm \ &H40) And 3: rg = (modrm \ 8) And 7: rm = modrm And 7
        If op = &HFF And md = 3 And rg = 0 And rm = ctrReg Then NativeForHasIncrement = True: Exit Function          'inc ctrReg (FF /0)
        If (op = &H83 Or op = &H81) And md = 3 And rg = 0 And rm = ctrReg Then NativeForHasIncrement = True: Exit Function  'add ctrReg, imm
        If op >= &HB8 And op <= &HBF Then derived(op - &HB8) = False               'mov reg,imm clears any derivation
        If op = &H3 And md = 3 And rm = ctrReg Then derived(rg) = True             'add rg, ctrReg
        If op = &H1 And md = 3 And rg = ctrReg Then derived(rm) = True             'add rm, ctrReg
        If op = &H8D Then                                                          'lea rg,[ctrReg + k] (no index)
            If NativeMemBase(dump) = ctrReg And NativeMemIndex(dump) < 0 Then derived(rg) = True
        End If
        If op = &H8B And md = 3 Then                                              'mov rg, rm
            If rg = ctrReg And derived(rm) Then NativeForHasIncrement = True: Exit Function
            If rm = ctrReg Then derived(rg) = True
        End If
        If op = &H89 And md = 3 Then                                              'mov rm, rg
            If rm = ctrReg And derived(rg) Then NativeForHasIncrement = True: Exit Function
            If rg = ctrReg Then derived(rm) = True
        End If
nexti:
    Next i
End Function

Private Function NativeForStartIsConst(arr() As CInstruction, ByVal hi As Long, ByVal ctrReg As Long) As Boolean
    'A counted For initialises its counter to a constant (mov ctr,imm or xor
    'ctr,ctr) before the header.  Scan backward to the NEAREST writer of the
    'counter and require it to be such an init; any other writer first (or none
    'within range) means this is not a clean counted loop, so reject it.
    Dim i As Long, lo As Long, dump As String, nn As Long, p As Long, op As Long, modrm As Long, md As Long, rg As Long, rm As Long
    On Error Resume Next
    lo = hi - 12: If lo < 0 Then lo = 0
    For i = hi - 1 To lo Step -1
        dump = Replace(arr(i).dump, " ", "")
        nn = Len(dump) \ 2
        If nn < 1 Then GoTo prev
        p = NativeOpStart(dump, nn)
        op = NativeDumpByte(dump, p)
        If op = (&HB8 + ctrReg) Then NativeForStartIsConst = True: Exit Function       'mov ctr, imm
        modrm = NativeDumpByte(dump, p + 1)
        md = (modrm \ &H40) And 3: rg = (modrm \ 8) And 7: rm = modrm And 7
        If (op = &H33 Or op = &H31 Or op = &H2B Or op = &H29) And md = 3 And rg = ctrReg And rm = ctrReg Then NativeForStartIsConst = True: Exit Function  'xor/sub ctr,ctr -> 0
        'Any OTHER writer of the counter seen first -> not a constant-init loop.
        If op = (&H40 + ctrReg) Or op = (&H48 + ctrReg) Then Exit Function             'inc/dec ctr
        If op = (&H58 + ctrReg) Then Exit Function                                      'pop ctr
        If op = &H8B And rg = ctrReg Then Exit Function                                 'mov ctr, r/m
        If op = &H8D And rg = ctrReg Then Exit Function                                 'lea ctr, [..]
        If op = &H89 And md = 3 And rm = ctrReg Then Exit Function                      'mov ctr, reg
        If (op = &H1 Or op = &H3 Or op = &H21 Or op = &H23 Or op = &H9 Or op = &HB) And md = 3 And rm = ctrReg Then Exit Function  'add/and/or ctr, reg
        If (op = &H83 Or op = &H81) And md = 3 And rm = ctrReg Then Exit Function       'add/sub/.. ctr, imm
prev:
    Next i
End Function

Private Function NativeLoopVarName(ByVal idx As Long) As String
    Const LV As String = "ijklmnopqrstuvwxyz"
    If idx >= 0 And idx < Len(LV) Then
        NativeLoopVarName = Mid$(LV, idx + 1, 1)
    Else
        NativeLoopVarName = "i" & CStr(idx)
    End If
End Function

Private Sub NativeDetectFpCompares(col As Collection)
    'Recognise VB6's floating-point relational, which materialises a Boolean:
    '    fcom(p) B ; fnstsw ax ; test ah,MASK ; j(e|ne) L1 ; mov REG,1 ;
    '    jmp L2 ; L1: xor REG,REG ; L2: [neg REG]
    'Record at the `test ah` VA the recovered relational operator + target REG,
    'and mark the je/jne/mov/jmp/xor/neg scaffolding for suppression.  The fcom's
    'operands are captured at run time (NVLastCmp); only the operator + structure
    'come from here.  Strict shape match - anything off leaves the code as-is.
    On Error GoTo done
    Dim n As Long, k As Long, i As Long, inst As CInstruction
    n = col.Count
    If n < 6 Then Exit Sub
    Dim arr() As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next
    For k = 1 To n - 6
        'fnstsw ax / fstsw ax
        Dim fd As String
        fd = UCase$(Replace(arr(k).dump, " ", ""))
        If Left$(fd, 4) <> "DFE0" And Left$(fd, 6) <> "9BDFE0" Then GoTo nextk2
        'test ah, imm8  (F6 C4 imm)
        Dim mask As Long
        If Not NativeIsTestAh(arr(k + 1), mask) Then GoTo nextk2
        'j(e|ne)
        If (arr(k + 2).cmdType And C_TYPEMASK) <> C_JMC Then GoTo nextk2
        Dim jm As String
        jm = NativeMnem(arr(k + 2))
        If jm <> "JE" And jm <> "JZ" And jm <> "JNE" And jm <> "JNZ" Then GoTo nextk2
        'mov REG, 1
        Dim regI As Long
        regI = NativeMovReg1(arr(k + 3))
        If regI < 0 Then GoTo nextk2
        'jmp
        If (arr(k + 4).cmdType And C_TYPEMASK) <> C_JMP Then GoTo nextk2
        'xor REG, REG  (same register that mov targeted)
        If Not NativeIsXorSelf(arr(k + 5), regI) Then GoTo nextk2
        Dim op As String
        op = NativeFpRelation(mask, jm)
        If Len(op) = 0 Then GoTo nextk2
        'Record the relational at the `test ah` instruction; suppress the rest.
        NativeColPut NVFpCmp, "P" & arr(k + 1).va, op & "|" & regI
        NativeColPut NVFpSkip, "P" & arr(k + 2).va, "1"      'je/jne
        NativeColPut NVFpSkip, "P" & arr(k + 3).va, "1"      'mov REG,1
        NativeColPut NVFpSkip, "P" & arr(k + 4).va, "1"      'jmp
        NativeColPut NVFpSkip, "P" & arr(k + 5).va, "1"      'xor REG,REG
        NativeAddUnique NVSkipLabels, arr(k + 2).jmpConst    'L1 (xor) label
        NativeAddUnique NVSkipLabels, arr(k + 4).jmpConst    'L2 label
        'Optional `neg REG` immediately after the xor's join point.
        If k + 6 <= n - 1 Then
            If NativeIsNegReg(arr(k + 6), regI) Then NativeColPut NVFpSkip, "P" & arr(k + 6).va, "1"
        End If
nextk2:
    Next k
done:
End Sub

Private Sub NativeDetectBoundsChecks(col As Collection)
    'VB6 SAFEARRAY element access emits a bounds-check guard before computing the
    'element address:
    '    test ARR,ARR / je ERR ; cmp word[ARR],1 / jne ERR ; ... index math ... ;
    '    cmp IDX,cElements / jb OK ; call __vbaGenerateBoundsError ; OK: lea ... ;
    '    jmp NEXT ; ERR: call __vbaGenerateBoundsError ; NEXT: ...
    'None of it is user code, but the je/jne/jb guards otherwise become bogus nested
    'If blocks (If arr <> 0 / If arr = 1 / If (idx - lb) >= cEls).  Suppress the guard
    'jumps, their feeding cmp/test, the skip-over jmp, and the bounds-error calls; the
    'kept mov/sub/lea still compute the element, and the orphaned ERR/OK/NEXT labels
    'strip out.  Anchored strictly on __vbaGenerateBoundsError calls (no other
    'construct branches to one), so it cannot misfire on real conditionals.
    On Error GoTo done
    Dim n As Long, k As Long, inst As CInstruction
    n = col.Count
    If n < 2 Then Exit Sub
    Dim arr() As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next
    'Pass 1: flag the __vbaGenerateBoundsError calls and drop them.  The error helper
    'is called either directly (`call [iat]`) or, when the proc calls it more than
    'once, through a register VB caches it into (`mov ebx,[iat]; call ebx`); track
    'which register currently holds it so the indirect form is recognised too.
    Dim isBerr() As Boolean, disp As Long, isAbs As Boolean, hasMem As Boolean
    Dim regBerr(7) As Boolean, di As Long, ci As Long, cr As String
    ReDim isBerr(n - 1)
    For k = 0 To n - 1
        If NativeMnem(arr(k)) = "MOV" Then
            di = NativeRegIndex(UCase$(NativeFirstReg(arr(k).command)))
            If di >= 0 And di <= 7 Then _
                regBerr(di) = (InStr(NativeApiName(arr(k)), "__vbaGenerateBoundsError") > 0)
        ElseIf (arr(k).cmdType And C_TYPEMASK) = C_CAL Then
            hasMem = NativeDecodeDisp(arr(k).dump, disp, isAbs)
            Dim berr As Boolean
            berr = False
            If hasMem And isAbs Then
                berr = (InStr(dsmNative.GetApiByIatVa(disp), "__vbaGenerateBoundsError") > 0)
            ElseIf Not hasMem Then
                cr = UCase$(Trim$(Mid$(Trim$(arr(k).command), 5)))   'call REG
                ci = NativeRegIndex(cr)
                If ci >= 0 And ci <= 7 Then berr = regBerr(ci)
            End If
            'Flag it as a bounds-error call so guard jumps targeting it are dropped,
            'but DON'T suppress the call itself: let it render through NativeRuntimeCall
            '(which silently drops __vbaGenerateBoundsError AND clears the push stack,
            'so a stray arg can't leak into the next real call).
            If berr Then isBerr(k) = True
        End If
    Next
    'Pass 2: suppress the guard jumps (+ feeding compare) and the skip-over jmp.
    'A jump points at / skips over a bounds-error STUB, which is the error call
    'optionally preceded by a `mov REG,[iat]` that loads the helper - so test for the
    'stub, not just the bare call.
    For k = 0 To n - 1
        Dim cls As Long, guard As Boolean
        cls = arr(k).cmdType And C_TYPEMASK
        If cls = C_JMC Then
            guard = False
            If NativeBerrStubFromIdx(arr, n, isBerr, NativeIdxOfVa(arr, n, arr(k).jmpConst)) Then guard = True  'je/jne -> ERR
            If NativeBerrStubFromIdx(arr, n, isBerr, k + 1) Then guard = True   'jb over a fall-through ERR
            If guard Then
                NativeColPut NVFpSkip, "P" & arr(k).va, "1"
                If k - 1 >= 0 Then
                    Dim pm As String
                    pm = NativeMnem(arr(k - 1))
                    If pm = "CMP" Or pm = "TEST" Then NativeColPut NVFpSkip, "P" & arr(k - 1).va, "1"
                End If
            End If
        ElseIf cls = C_JMP Then
            'The unconditional jmp that skips over an ERR stub (jmp NEXT around ERR).
            If NativeBerrStubFromIdx(arr, n, isBerr, k + 1) Then NativeColPut NVFpSkip, "P" & arr(k).va, "1"
        End If
    Next
done:
End Sub

Private Function NativeIdxOfVa(arr() As CInstruction, ByVal n As Long, ByVal va As Long) As Long
    Dim k As Long
    NativeIdxOfVa = -1
    If va = 0 Then Exit Function
    For k = 0 To n - 1
        If arr(k).va = va Then NativeIdxOfVa = k: Exit Function
    Next
End Function

Private Function NativeBerrStubFromIdx(arr() As CInstruction, ByVal n As Long, isBerr() As Boolean, ByVal idx As Long) As Boolean
    'True when a bounds-error call begins at idx, allowing the `mov REG,[iat]` helper
    'load(s) VB places just before an indirect `call REG` (only movs may precede).
    Dim j As Long, lim As Long
    If idx < 0 Then Exit Function
    lim = idx + 3: If lim > n - 1 Then lim = n - 1
    For j = idx To lim
        If isBerr(j) Then NativeBerrStubFromIdx = True: Exit Function
        If NativeMnem(arr(j)) <> "MOV" Then Exit Function
    Next
End Function

Private Sub NativeDetectStrCmpCompares(col As Collection)
    'Recognise VB6's string relational:
    '    call __vbaStrCmp ; [mov REG,eax] ; neg REG ; sbb REG,REG ; inc REG ; neg REG
    'The neg/sbb/inc/neg turns the strcmp tri-state into the Boolean (strcmp = 0),
    'i.e. (a = b).  Record the target REG against the call VA (so the runtime
    'handler binds "(a = b)" into REG using the strcmp operands), and suppress the
    'four materialisation instructions so they don't clobber REG's tracked value.
    'The optional `mov REG,eax` is LEFT in place - it copies the Boolean from eax
    'into REG.  When absent, the materialisation is on eax itself (REG = 0).
    'Strict shape match - anything off leaves the strcmp rendered as a Call.
    On Error GoTo done
    Dim n As Long, k As Long
    n = col.Count
    If n < 5 Then Exit Sub
    Dim arr() As CInstruction, inst As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next
    For k = 0 To n - 5
        If (arr(k).cmdType And C_TYPEMASK) <> C_CAL Then GoTo nexts
        Dim disp As Long, isAbs As Boolean
        If Not NativeDecodeDisp(arr(k).dump, disp, isAbs) Then GoTo nexts
        If Not isAbs Then GoTo nexts
        Dim cnm As String
        cnm = dsmNative.GetApiByIatVa(disp)
        If InStr(cnm, "__vbaStrCmp") = 0 And InStr(cnm, "__vbaStrComp") = 0 Then GoTo nexts
        'Optional `mov REG, eax` right after the call.
        Dim t As Long, reg As Long
        t = k + 1
        reg = NativeMovFromEax(arr(t))
        If reg >= 0 Then t = t + 1 Else reg = 0      'materialise on eax when no mov
        'neg REG ; sbb REG,REG ; inc REG ; neg REG
        If t + 3 > n - 1 Then GoTo nexts
        If Not NativeIsNegReg2(arr(t), reg) Then GoTo nexts
        If Not NativeIsSbbSelf(arr(t + 1), reg) Then GoTo nexts
        If Not NativeIsIncReg(arr(t + 2), reg) Then GoTo nexts
        If Not NativeIsNegReg2(arr(t + 3), reg) Then GoTo nexts
        NativeColPut NVStrCmpReg, "P" & arr(k).va, CStr(reg)
        NativeColPut NVFpSkip, "P" & arr(t).va, "1"          'neg
        NativeColPut NVFpSkip, "P" & arr(t + 1).va, "1"      'sbb
        NativeColPut NVFpSkip, "P" & arr(t + 2).va, "1"      'inc
        NativeColPut NVFpSkip, "P" & arr(t + 3).va, "1"      'neg
nexts:
    Next k
done:
End Sub

Private Function NativeMovFromEax(inst As CInstruction) As Long
    'Match `mov REG, eax` -> dest register index, else -1.  Encodings:
    '  8B /r  mod=3 rm=000(eax)      -> dest = reg field
    '  89 /r  mod=3 reg=000(eax)     -> dest = rm field
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long
    NativeMovFromEax = -1
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    modrm = NativeDumpByte(dump, i + 1)
    If (modrm \ &H40) <> 3 Then Exit Function
    If op = &H8B Then
        If (modrm And 7) = 0 Then NativeMovFromEax = (modrm \ 8) And 7
    ElseIf op = &H89 Then
        If ((modrm \ 8) And 7) = 0 Then NativeMovFromEax = modrm And 7
    End If
End Function

Private Function NativeIsNegReg2(inst As CInstruction, ByVal reg As Long) As Boolean
    'Match `neg reg` (F7 /3: ModRM md=3, reg-field=3, rm=reg).
    Dim dump As String, nn As Long, i As Long, modrm As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    If NativeDumpByte(dump, i) <> &HF7 Then Exit Function
    modrm = NativeDumpByte(dump, i + 1)
    NativeIsNegReg2 = ((modrm \ &H40) = 3) And (((modrm \ 8) And 7) = 3) And ((modrm And 7) = reg)
End Function

Private Function NativeIsSbbSelf(inst As CInstruction, ByVal reg As Long) As Boolean
    'Match `sbb reg,reg` (1B /r or 19 /r, md=3, reg field = rm = reg).
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op <> &H1B And op <> &H19 Then Exit Function
    modrm = NativeDumpByte(dump, i + 1)
    If (modrm \ &H40) <> 3 Then Exit Function
    NativeIsSbbSelf = (((modrm \ 8) And 7) = reg) And ((modrm And 7) = reg)
End Function

Private Function NativeIsIncReg(inst As CInstruction, ByVal reg As Long) As Boolean
    'Match `inc reg` (single-byte 0x40+reg, or FF /0 md=3 rm=reg).
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op = &H40 + reg Then NativeIsIncReg = True: Exit Function
    If op = &HFF Then
        modrm = NativeDumpByte(dump, i + 1)
        If (modrm \ &H40) = 3 And (((modrm \ 8) And 7) = 0) And ((modrm And 7) = reg) Then NativeIsIncReg = True
    End If
End Function

Private Function NativeIsTestAh(inst As CInstruction, ByRef mask As Long) As Boolean
    'Match `test ah, imm8` (F6 /0 with ModRM 0xC4 -> md=3, reg=0 test, rm=4 = AH).
    Dim dump As String, nn As Long, i As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    If NativeDumpByte(dump, i) <> &HF6 Then Exit Function
    If NativeDumpByte(dump, i + 1) <> &HC4 Then Exit Function
    mask = NativeDumpByte(dump, i + 2)
    NativeIsTestAh = True
End Function

Private Function NativeMovReg1(inst As CInstruction) As Long
    'Match `mov r32, 1` (B8+reg, imm32 = 1) -> the register index, else -1.
    Dim dump As String, nn As Long, i As Long, op As Long
    NativeMovReg1 = -1
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op < &HB8 Or op > &HBF Then Exit Function
    If NativeDumpInt32(dump, i + 1) <> 1 Then Exit Function
    NativeMovReg1 = op - &HB8
End Function

Private Function NativeIsXorSelf(inst As CInstruction, ByVal reg As Long) As Boolean
    'Match `xor reg,reg` / `sub reg,reg` (md=3, reg field = rm = reg).
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op <> &H33 And op <> &H31 And op <> &H2B And op <> &H29 Then Exit Function
    modrm = NativeDumpByte(dump, i + 1)
    If (modrm \ &H40) <> 3 Then Exit Function
    NativeIsXorSelf = (((modrm \ 8) And 7) = reg) And ((modrm And 7) = reg)
End Function

Private Function NativeIsNegReg(inst As CInstruction, ByVal reg As Long) As Boolean
    'Match `neg reg` (F7 /3, ModRM md=3 reg-field=3 rm=reg).
    Dim dump As String, nn As Long, i As Long, modrm As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    If NativeDumpByte(dump, i) <> &HF7 Then Exit Function
    modrm = NativeDumpByte(dump, i + 1)
    NativeIsNegReg = ((modrm \ &H40) = 3) And (((modrm \ 8) And 7) = 3) And ((modrm And 7) = reg)
End Function

Private Function NativeIsSetcc(inst As CInstruction, ByRef op As String, ByRef reg As Long) As Boolean
    'Match SETcc r/m8 (0F 9x, reg-direct).  Returns the relational the condition
    'represents (reg = 1 when true) and the 32-bit parent register index.
    Dim dump As String, nn As Long, i As Long, cc As Long, modrm As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    If NativeDumpByte(dump, i) <> &HF Then Exit Function
    cc = NativeDumpByte(dump, i + 1)
    Select Case cc
        Case &H9C, &H92: op = "<"        'setl / setb
        Case &H9F, &H97: op = ">"        'setg / seta
        Case &H9E, &H96: op = "<="       'setle / setbe
        Case &H9D, &H93: op = ">="       'setge / setae
        Case &H94: op = "="              'sete
        Case &H95: op = "<>"             'setne
        Case Else: Exit Function
    End Select
    modrm = NativeDumpByte(dump, i + 2)
    If (modrm \ &H40) <> 3 Then Exit Function       'register operand only
    reg = (modrm And 7) And 3                        'al/cl/dl/bl + ah/ch/dh/bh -> eax/ecx/edx/ebx
    NativeIsSetcc = True
End Function

Private Function NativeIsBoolOrAnd(inst As CInstruction, ByRef op As String, ByRef dst As Long, ByRef src As Long) As Boolean
    'Match `or/and r32, r32` (reg-direct).  Returns "Or"/"And" and the dest + source
    'register indices (dest is the register read by the following test/Jcc).
    Dim dump As String, nn As Long, i As Long, opc As Long, modrm As Long, reg As Long, rm As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    opc = NativeDumpByte(dump, i)
    modrm = NativeDumpByte(dump, i + 1)
    If (modrm \ &H40) <> 3 Then Exit Function
    reg = (modrm \ 8) And 7: rm = modrm And 7
    Select Case opc
        Case &HB: op = "Or": dst = reg: src = rm           'or  r32, r/m32
        Case &H9: op = "Or": dst = rm: src = reg           'or  r/m32, r32
        Case &H23: op = "And": dst = reg: src = rm         'and r32, r/m32
        Case &H21: op = "And": dst = rm: src = reg         'and r/m32, r32
        Case Else: Exit Function
    End Select
    NativeIsBoolOrAnd = True
End Function

Private Function NativeFpRelation(ByVal mask As Long, ByVal jcc As String) As String
    'Relational the materialised Boolean represents (REG True <=> jcc NOT taken).
    'fcom sets C0(0x01)=A<B, C3(0x40)=A=B, C2(0x04)=unordered.  je selects the
    'base relation (ah&mask)!=0; jne selects its negation.
    Dim base As String
    Select Case mask
        Case &H1, &H5: base = "<"        'C0 [|C2]
        Case &H40, &H44: base = "="      'C3 [|C2]
        Case &H41, &H45: base = "<="     'C0|C3 [|C2]
        Case Else: Exit Function
    End Select
    If jcc = "JE" Or jcc = "JZ" Then
        NativeFpRelation = base
    Else
        NativeFpRelation = NativeNegOp(base)
    End If
End Function

Private Function NativeLooksRelational(ByVal s As String) As Boolean
    'True when s contains a relational operator (so it is a Boolean value, not a
    'plain arithmetic expression) - used to render `If <bool>` instead of the
    'redundant `If (<bool>) <> 0`.
    NativeLooksRelational = (InStr(s, " > ") > 0) Or (InStr(s, " < ") > 0) _
        Or (InStr(s, " >= ") > 0) Or (InStr(s, " <= ") > 0) _
        Or (InStr(s, " <> ") > 0) Or (InStr(s, " = ") > 0) _
        Or (InStr(s, " Is ") > 0)
End Function

Private Function NVProcEndApprox(ByVal addr As Long) As Long
    'Upper bound for case-target validation: the next discovered proc, else +8KB.
    Dim pe As Long, e As Long
    NVProcEndApprox = addr + 8190
    On Error Resume Next
    For pe = 0 To UBound(gNativeProcArray) - 1
        e = gNativeProcArray(pe).offset
        If e > addr And e < NVProcEndApprox Then NVProcEndApprox = e
    Next
End Function

Private Function NativeJmpTableInfo(inst As CInstruction, ByRef idxReg As Long, ByRef tbl As Long) As Boolean
    'Detect `jmp dword ptr [idxReg*4 + disp32]` (FF /4, ModRM 0x24, SIB scale=4 base=none).
    Dim dump As String, nn As Long, i As Long, modrm As Long, sib As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    If i + 7 > nn Then Exit Function           'need FF + ModRM + SIB + disp32
    If NativeDumpByte(dump, i) <> &HFF Then Exit Function
    modrm = NativeDumpByte(dump, i + 1)
    If modrm <> &H24 Then Exit Function                 'md=0, reg=4 (jmp), rm=4 (SIB)
    sib = NativeDumpByte(dump, i + 2)
    If (sib \ &H40) <> 2 Then Exit Function             'scale must be *4
    If (sib And 7) <> 5 Then Exit Function              'base = 5 -> disp32 (no base reg)
    idxReg = (sib \ 8) And 7
    If idxReg = 4 Then Exit Function                    'no index
    tbl = NativeDumpInt32(dump, i + 3)
    NativeJmpTableInfo = (tbl >= OptHeader.ImageBase)
End Function

Private Function NativeCmpRegImm(inst As CInstruction, ByVal reg As Long, ByRef imm As Long) As Boolean
    'Match `cmp reg, imm` (83 /7 imm8, 81 /7 imm32, or 3D imm32 for eax).
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op = &H3D Then
        If reg <> 0 Then Exit Function
        imm = NativeDumpInt32(dump, i + 1): NativeCmpRegImm = True: Exit Function
    End If
    If op <> &H83 And op <> &H81 Then Exit Function
    modrm = NativeDumpByte(dump, i + 1)
    If (modrm \ &H40) <> 3 Then Exit Function            'md=3 (register)
    If ((modrm \ 8) And 7) <> 7 Then Exit Function       'reg field = 7 (cmp)
    If (modrm And 7) <> reg Then Exit Function
    If op = &H83 Then imm = NativeDumpInt8(dump, i + 2) Else imm = NativeDumpInt32(dump, i + 2)
    NativeCmpRegImm = True
End Function

Private Function NativeSubRegImm(inst As CInstruction, ByVal reg As Long, ByRef imm As Long) As Boolean
    'Match `sub reg, imm` (83 /5 imm8 or 81 /5 imm32) for a Select Case lower-bound base.
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op <> &H83 And op <> &H81 Then Exit Function
    modrm = NativeDumpByte(dump, i + 1)
    If (modrm \ &H40) <> 3 Then Exit Function
    If ((modrm \ 8) And 7) <> 5 Then Exit Function       'reg field = 5 (sub)
    If (modrm And 7) <> reg Then Exit Function
    If op = &H83 Then imm = NativeDumpInt8(dump, i + 2) Else imm = NativeDumpInt32(dump, i + 2)
    NativeSubRegImm = True
End Function

Private Function NativeProcessInst(inst As CInstruction) As String
'*****************************
'Turn one instruction into a VB statement (when it completes a recognised
'idiom) or an annotated-assembly line.
'*****************************
    Dim cls As Long, ind As String, ann As String, vb As String, mn As String
    cls = inst.cmdType And C_TYPEMASK
    mn = NativeMnem(inst)
    ind = NativeIndentStr()

    'TEST/CMP set the flags consumed by the next conditional jump.  Record the
    'operands now (the relational operator is resolved later from the Jcc).
    If mn = "TEST" Or mn = "CMP" Then
        NativeDecodeCompare inst, mn
        Exit Function
    End If

    'SETcc materialises a relational into a register: bind "(L <op> R)" from the
    'pending compare so the value flows into a following boolean combine / test.
    Dim ccOp As String, ccReg As Long
    If NativeIsSetcc(inst, ccOp, ccReg) Then
        If NVCmpSet And Len(NVCmpL) > 0 Then
            Dim ccR As String
            If NVCmpIsTest Then ccR = "0" Else ccR = NVCmpR
            If Len(ccR) > 0 Then NVReg(ccReg) = "(" & NVCmpL & " " & ccOp & " " & ccR & ")" Else NVReg(ccReg) = ""
        Else
            NVReg(ccReg) = ""
        End If
        NVRegIsAddr(ccReg) = False: NVRegIsMe(ccReg) = False: NVRegIsFormVt(ccReg) = False
        NVRegObjType(ccReg) = "": NVRegObjVt(ccReg) = "": NVRegObjGuid(ccReg) = "": NVRegObjVtGuid(ccReg) = ""
        NVCmpSet = False: NVCmpL = "": NVCmpR = "": NVCmpIsTest = False
        Exit Function
    End If

    'OR / AND of two relational-Boolean registers is a compound condition
    '(`x <= 0 And x >= -10000` compiles to setl/setg/or; the OR/AND also sets the
    'flags the following Jcc reads).  Combine them and arm the Boolean condition.
    Dim boolOp As String, bDst As Long, bSrc As Long
    If NativeIsBoolOrAnd(inst, boolOp, bDst, bSrc) Then
        Dim bv1 As String, bv2 As String
        bv1 = NVReg(bDst): bv2 = NVReg(bSrc)
        If NativeLooksRelational(bv1) And NativeLooksRelational(bv2) Then
            Dim combined As String
            combined = "(" & bv1 & " " & boolOp & " " & bv2 & ")"
            NVReg(bDst) = combined
            NVRegIsAddr(bDst) = False: NVRegIsMe(bDst) = False: NVRegIsFormVt(bDst) = False
            NVRegObjType(bDst) = "": NVRegObjVt(bDst) = "": NVRegObjGuid(bDst) = "": NVRegObjVtGuid(bDst) = ""
            NVCmpL = combined: NVCmpIsBool = True: NVCmpSet = True: NVCmpIsTest = False
            Exit Function
        End If
    End If

    Select Case cls

        Case C_CAL
            Dim disp As Long, isAbs As Boolean, hasMem As Boolean, rn As String
            hasMem = NativeDecodeDisp(inst.dump, disp, isAbs)
            If hasMem And isAbs Then
                'call [abs] -> msvbvm60 runtime helper / imported API (IAT)
                vb = NativeRuntimeCall(inst, dsmNative.GetApiByIatVa(disp))
                If Not NVKeepPushStack Then NVPushTop = 0
                If NVCallHandled Then
                    If Len(vb) > 0 Then NativeProcessInst = ind & vb & vbCrLf
                    Exit Function
                End If
                ann = "call " & dsmNative.GetApiByIatVa(disp)
            ElseIf hasMem Then
                'Property access on a resolved intrinsic object (e.g. App.Path):
                'when the call's base register holds an intrinsic object's vtable
                '(tagged by the getter chain `mov reg,[App_local]; mov vt,[reg]`),
                'map the call offset via the intrinsic property table.  Checked
                'first - a tagged object vtable is a strong, specific signal.
                Dim ocb As Long, oprop As String
                ocb = NativeMemBase(inst.dump)
                'Property GET on the VB6 Err object (call [ErrVt + offset]).  The Err
                'vtable was tagged by the rtcErrObj -> __vbaObjSet -> deref chain.
                'Resolve to Err.Number / Err.Description and flow the value quietly to
                'the out-param local (it is normally consumed by an error-message
                'concatenation, so no standalone statement is emitted).
                If ocb >= 0 And ocb <= 7 Then
                    If NVRegObjVt(ocb) = "Err" Then
                        Dim eprop As String
                        eprop = NativeErrPropByOffset(disp)
                        If Len(eprop) > 0 Then
                            NVPushTop = 0
                            NVReg(0) = "Err." & eprop
                            NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                            NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjGuid(0) = "": NVRegObjVtGuid(0) = ""
                            If NVLastLeaSet Then
                                NativeSetLocalExpr NVLastLea, "Err." & eprop
                                NVLastLeaSet = False
                            End If
                            Exit Function
                        End If
                    End If
                End If
                If ocb >= 0 And ocb <= 7 Then
                    If Len(NVRegObjVt(ocb)) > 0 Then
                        oprop = NativeIntrinsicPropByOffset(NVRegObjVt(ocb), disp)
                        If Len(oprop) > 0 Then
                            Dim ochain As String
                            ochain = NVRegObjVt(ocb) & "." & oprop
                            NVPushTop = 0
                            NVReg(0) = ochain
                            NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                            NVRegObjType(0) = "": NVRegObjVt(0) = ""
                            If NVLastLeaSet Then
                                Dim oln As String
                                oln = "var_" & Hex$(Abs(NVLastLea))
                                NativeSetLocalExpr NVLastLea, ochain
                                NVLastLeaSet = False
                                NativeProcessInst = ind & oln & " = " & ochain & vbCrLf
                            End If
                            Exit Function           'value flows to the consumer
                        End If
                    End If
                End If
                'Property access on a tracked CONTROL object: its vtable register
                'carries the control GUID (set when the control was fetched via a
                'form accessor and stored through __vbaObjSet).  Reuse NativeProperty
                'by seeding the last-control/guid from the per-register tracking, so
                'a Get folds into `var = Form1.File1.Prop` and a Let/Set into
                '`Form1.File1.Prop = value`.
                If ocb >= 0 And ocb <= 7 Then
                    If Len(NVRegObjVtGuid(ocb)) > 0 Then
                        vb = NativeControlProp(NVRegObjVt(ocb), NVRegObjVtGuid(ocb), disp)
                        If Len(vb) > 0 Then NVPushTop = 0: NativeProcessInst = ind & vb & vbCrLf: Exit Function
                    End If
                End If
                'Method call on a tracked USER-CLASS instance (e.g. ds7.Load): the
                'vtable register carries the class name (typed at __vbaNew) and the
                'receiver expression.  User methods sit at vtable offset >= 0x1C.
                If ocb >= 0 And ocb <= 7 Then
                    If Len(NVRegObjVt(ocb)) > 0 And disp >= &H1C Then
                        Dim umAddr As Long
                        umAddr = NativeClassMethodAddr(NVRegObjVt(ocb), disp)
                        If umAddr <> 0 Then
                            Dim umName As String, umDot As Long, umRecv As String, umArgs As String, umSig As String
                            umName = NativeCallTargetName(umAddr)
                            umDot = InStr(umName, "."): If umDot > 0 Then umName = Mid$(umName, umDot + 1)
                            umRecv = NVRegObjInst(ocb)
                            If Len(umRecv) = 0 Then umRecv = NVRegObjVt(ocb)
                            'Args (source order) lead with the `this` pointer (pushed
                            'last) and may end with a hidden return buffer (pushed
                            'first).  Drop the leading `this` when it is the receiver,
                            'then keep the method's real parameter count from the front.
                            umArgs = NativeArgList()
                            Dim umP() As String, umStart As Long, umKeep As Long, umI As Long, umOut As String
                            Dim umTotal As Long, umRetbuf As String
                            umTotal = 0: umRetbuf = ""
                            If Len(umArgs) > 0 Then
                                umP = Split(umArgs, ", ")
                                umStart = 0
                                If umP(0) = umRecv Then umStart = 1
                                umTotal = UBound(umP) - umStart + 1
                                umKeep = umTotal
                                If NativeTryMethodSig(umAddr, umSig) Then umKeep = NativeArgCount(umSig)
                                If umKeep > umTotal Then umKeep = umTotal
                                For umI = umStart To umStart + umKeep - 1
                                    If umI > UBound(umP) Then Exit For
                                    If Len(umOut) > 0 Then umOut = umOut & ", "
                                    umOut = umOut & umP(umI)
                                Next
                                umArgs = umOut
                                'An argument beyond the real parameter count is the
                                'hidden [out,retval] retbuf - the method returns a value.
                                If umTotal > umKeep Then umRetbuf = umP(umStart + umKeep)
                            End If
                            'A resolved method call on a module global proves its
                            'class - type its declaration `Public global_X As <Class>`
                            '(reliable, unlike guessing from the __vbaNew args).
                            If Left$(umRecv, 7) = "global_" Then
                                Dim grcVa As Long
                                grcVa = CLng("&H" & Mid$(umRecv, 8, 8))
                                If gNativeGlobalClass Is Nothing Then Set gNativeGlobalClass = New Collection
                                On Error Resume Next
                                gNativeGlobalClass.Add NVRegObjVt(ocb), "g" & grcVa
                                On Error GoTo 0
                            End If
                            'A value-returning method (Function or Property Get)
                            'delivers its result through the hidden retbuf out-param.
                            'Fold recv.method(realArgs) into that retbuf local + eax so
                            'the value flows to its consumer (e.g.
                            'global_108 = clsBitmap.InvertImageDC, or
                            'If picBmp.SetBitmap(x) Then) instead of a bogus Call.
                            Dim umKind As String, umVal As String, umIsVal As Boolean
                            umIsVal = (Len(umRetbuf) > 0)
                            If Not umIsVal Then
                                If NativeTryMethodKind(umAddr, umKind) Then umIsVal = (InStr(umKind, "Get") > 0)
                            End If
                            If umIsVal Then
                                umVal = umRecv & "." & umName
                                If Len(umArgs) > 0 Then umVal = umVal & "(" & umArgs & ")"
                                NVPushTop = 0
                                NVReg(0) = umVal
                                NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjInst(0) = ""
                                If Left$(umRetbuf, 4) = "var_" Then
                                    On Error Resume Next
                                    NativeSetLocalExpr -CLng("&H" & Mid$(umRetbuf, 5)), umVal
                                    On Error GoTo 0
                                ElseIf NVLastLeaSet Then
                                    NativeSetLocalExpr NVLastLea, umVal: NVLastLeaSet = False
                                End If
                                Exit Function           'value flows to the consumer
                            End If
                            NVPushTop = 0
                            NativeProcessInst = ind & "Call " & umRecv & "." & umName & "(" & umArgs & ")" & vbCrLf
                            Exit Function
                        End If
                    End If
                End If
                'call [reg+disp] -> the object's own method (via its vtable), a
                'control accessor, or a property vtable call.
                'VB6 intrinsic global objects (App/Screen/Clipboard) are getters
                'at fixed low vtable offsets on the runtime "Global" object (held
                'in a module global) or the form.  Resolve `call [objVt + 0x14]`
                'etc. when objVt is a tracked object vtable.  These offsets are
                'IDispatch slots on arbitrary objects, which VB never raw-calls, so
                'the object-vtable guard keeps it safe.
                Dim gobj As String, gcb As Long
                gobj = NativeGlobalObjByOffset(disp)
                If Len(gobj) > 0 Then
                    gcb = NativeMemBase(inst.dump)
                    If gcb >= 0 And gcb <= 7 Then
                        If NVRegIsFormVt(gcb) Then
                            NVPushTop = 0
                            NVReg(0) = gobj
                            NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                            NVRegObjType(0) = gobj: NVRegObjVt(0) = ""    'eax now holds the App/Screen/Clipboard pointer
                            'The object is written to the out-param local addressed
                            'by the LEA just before the call - surface var_X = App.
                            If NVLastLeaSet Then
                                Dim gln As String
                                gln = "var_" & Hex$(Abs(NVLastLea))
                                NativeSetLocalExpr NVLastLea, gobj
                                NVLastLeaSet = False
                                NativeProcessInst = ind & gln & " = " & gobj & vbCrLf
                            End If
                            Exit Function           'value flows to the consumer
                        End If
                    End If
                End If
                'A built-in Form method on Me's OWN vtable: call [Me_vt + off] where
                'off is a fixed _Form-interface slot (e.g. Hide = 0x2B4).  Gated to a
                'tracked Me vtable and checked BEFORE the control heuristic because
                'these offsets fall in the control-accessor range (< 0x2F8) and would
                'otherwise be mis-read as a control property call.
                Dim fmeth As String
                If ocb >= 0 And ocb <= 7 Then
                    If NVRegIsFormVt(ocb) Then
                        fmeth = NativeFormMethodByOffset(disp)
                        If Len(fmeth) > 0 Then
                            NVPushTop = 0
                            NVReg(0) = "": NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                            NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjGuid(0) = "": NVRegObjVtGuid(0) = ""
                            NativeProcessInst = ind & "Me." & fmeth & vbCrLf
                            Exit Function
                        End If
                    End If
                End If
                'A form calling its own method: call [vtable + 0x6F8 + slot*4].
                'Checked first: requiring a real gFormVtable slot is a stronger
                'signal than the NVBase control heuristic (which the same form-method
                'calls can otherwise mis-solve into a bogus control base).
                Dim ftgt As Long
                ftgt = NativeFormVtableTarget(disp)
                If ftgt <> 0 Then
                    'A COM-style form-method call leads with the implicit Me/this (the
                    'form pointer, tracked as arg_8 when it reaches the arg stack); drop
                    'it and the trailing retbuf, keeping the real parameters (count from
                    'the method's typeinfo signature) - so `Update_Status()` no longer
                    'renders as frmMain.Update_Status(arg_8).
                    Dim fpname As String, fargs As String, fRetbuf As String
                    fargs = NativeDropThisArgs(NativeArgList(), ftgt, fRetbuf)
                    fpname = NativeCallTargetName(ftgt)
                    NVReg(0) = NVForm & "." & fpname & "(" & fargs & ")"
                    NVPendingCall = "Call " & NVForm & "." & fpname & "(" & fargs & ")"
                    Exit Function
                End If
                'A class calling its OWN method: call [Me_vtable + off] where off is a
                'COM-vtable user-method slot (>=0x1C).  ocb holds Me's vtable
                '(NVRegIsFormVt, the deref of the Me pointer), which guards against
                'mis-resolving a call on some other object's vtable.
                Dim ctgt As Long
                ctgt = NativeClassVtableTarget(disp)
                If ctgt <> 0 And ocb >= 0 And ocb <= 7 Then
                    If NVRegIsFormVt(ocb) Then
                        'Emitted directly (not deferred): a COM method returns its
                        'result through a hidden by-reference buffer, leaving an HRESULT
                        'in eax that VB immediately error-checks (cmp eax,0) - which
                        'would otherwise fold/consume a deferred call and lose it.
                        Dim cpname As String, cargs As String, csig As String, ckind As String
                        Dim cP() As String, cStart As Long, cKeep As Long, cTotal As Long, cRetbuf As String, cI As Long, cOut As String
                        cargs = NativeArgList()
                        cpname = NativeCallTargetName(ctgt)
                        cRetbuf = ""
                        If Len(cargs) > 0 Then
                            cP = Split(cargs, ", ")
                            'A COM vtable call always leads with the implicit Me/this
                            '(the Me pointer at ebp+8, tracked as arg_8) - drop it.
                            cStart = 0
                            If cP(0) = "arg_8" Then cStart = 1
                            cTotal = UBound(cP) - cStart + 1
                            cKeep = cTotal
                            If NativeTryMethodSig(ctgt, csig) Then cKeep = NativeArgCount(csig)
                            If cKeep > cTotal Then cKeep = cTotal
                            For cI = cStart To cStart + cKeep - 1
                                If cI > UBound(cP) Then Exit For
                                If Len(cOut) > 0 Then cOut = cOut & ", "
                                cOut = cOut & cP(cI)
                            Next
                            cargs = cOut
                            'An argument beyond the real parameter count is the hidden
                            '[out,retval] retbuf - the method returns a value through it.
                            If cTotal > cKeep Then cRetbuf = cP(cStart + cKeep)
                        End If
                        'A value-returning method (retbuf present, or a Function/Get)
                        'folds recv.method(args) into the retbuf local + eax so the
                        'result flows to its consumer (var_X = clsBitmap.SetBitmap(...)).
                        Dim cIsVal As Boolean, cVal As String
                        cIsVal = (Len(cRetbuf) > 0)
                        If Not cIsVal Then
                            If NativeTryMethodKind(ctgt, ckind) Then cIsVal = (InStr(ckind, "Get") > 0)
                        End If
                        NVPushTop = 0
                        If cIsVal Then
                            cVal = NVForm & "." & cpname
                            If Len(cargs) > 0 Then cVal = cVal & "(" & cargs & ")"
                            NativeResetValue
                            NVReg(0) = cVal
                            NVRegObjType(0) = "": NVRegObjVt(0) = ""
                            If Left$(cRetbuf, 4) = "var_" Then
                                On Error Resume Next
                                NativeSetLocalExpr -CLng("&H" & Mid$(cRetbuf, 5)), cVal
                                On Error GoTo 0
                            End If
                            Exit Function
                        End If
                        NativeResetValue
                        NativeProcessInst = ind & "Call " & NVForm & "." & cpname & "(" & cargs & ")" & vbCrLf
                        Exit Function
                    End If
                End If
                If NVBase >= 0 And disp >= NVBase And NativeControlByOffset(disp) <> "" Then
                    NVLastControl = NVForm & "." & NativeControlByOffset(disp)
                    NVLastGuid = NativeGuidByOffset(disp)
                    NVPushTop = 0
                    'Tag the returned control object (in eax) with its identity so a
                    'following __vbaObjSet store binds the local's GUID and a direct
                    'property call on it resolves.  The value flows to the consumer.
                    NVReg(0) = NVLastControl
                    NVRegObjType(0) = NVLastControl: NVRegObjGuid(0) = NVLastGuid
                    NVRegObjVt(0) = "": NVRegObjVtGuid(0) = ""
                    NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                    Exit Function
                Else
                    'A property vtable call that the per-register control tracking
                    'above did NOT resolve.  The old single-slot NVLastControl is
                    'stale-prone (it names whatever control was fetched last, not the
                    'one this call is on), so leave the call raw rather than emit a
                    'wrong Form.Control.Prop.  Correct resolutions come only from the
                    'tracked-vtable path (NativeControlProp).
                    NVPushTop = 0
                    'Drop the COM lifetime/identity plumbing VB6 emits around every
                    'object use - QueryInterface / AddRef / Release at the IUnknown
                    'vtable slots 0 / 4 / 8 - which is pure noise, never user code.
                    If disp = 0 Or disp = 4 Or disp = 8 Then Exit Function
                    ann = "call ." & Hex$(disp)
                End If
            Else
                'A direct (E8) call to a user Sub/Function within the image
                Dim tgt As Long
                tgt = NativeCallTarget(inst)
                If tgt >= OptHeader.ImageBase And NativeInImage(tgt) Then
                    Dim pname As String, uargs As String
                    pname = NativeCallTargetName(tgt)
                    uargs = NativeArgList()
                    NVPushTop = 0
                    'Keep the call as the value in eax and defer the "Call X()"
                    'statement: if the next instruction consumes the result it
                    'folds into an assignment / argument / condition; otherwise
                    'the deferred call is emitted as a statement (see the decode
                    'loop's fold/flush of NVPendingCall).
                    NVReg(0) = pname & "(" & uargs & ")"
                    NVPendingCall = "Call " & pname & "(" & uargs & ")"
                    Exit Function
                End If
                'call <reg> -> a runtime helper previously cached into the register
                rn = NativeCallRegName(inst)
                vb = NativeRuntimeCall(inst, rn)
                If Not NVKeepPushStack Then NVPushTop = 0
                If NVCallHandled Then
                    If Len(vb) > 0 Then NativeProcessInst = ind & vb & vbCrLf
                    Exit Function
                End If
                ann = "call " & IIf(Len(rn) > 0, rn, "<reg>")
            End If

        Case C_JMC
            'A jo/jno integer-overflow check: VB6 emits one after arithmetic on
            'Integer/Long operands, jumping to an overflow-error stub appended
            'after the proc epilogue.  No VB source construct produces jo/jno, so
            'this is never user code - drop it (target -> skip label) rather than
            'let it become a bogus If block.
            If mn = "JO" Or mn = "JNO" Then
                NativeAddUnique NVSkipLabels, inst.jmpConst
                NVCmpSet = False: NVLastCmp = ""
                Exit Function
            End If
            'A Jcc guarding VB6's automatic FPU overflow check (preceded by
            'fnstsw ax; test al,imm) jumps to the error handler - drop it (it can be
            'a FAR jump, so it is not covered by the short error-check rule below).
            If NVFpuChk Then
                NativeAddUnique NVSkipLabels, inst.jmpConst
                NVFpuChk = False: NVCmpSet = False: NVLastCmp = ""
                Exit Function
            End If
            'A short forward jge/jns/jae after a call is VB's automatic
            'HRESULT error-check guard around __vbaHresultCheckObj - drop it so
            'it does not turn into a bogus If block.
            If inst.jmpConst > inst.va And NativeIsErrorCheckJcc(mn) And (inst.jmpConst - inst.va) <= 48 Then
                NativeAddUnique NVSkipLabels, inst.jmpConst
                Exit Function
            End If
            'Top-tested Do While header (detected pre-pass): the exit Jcc becomes the
            'loop condition - the loop runs while the jump is NOT taken.
            If Len(NativeColGet(NVWhileCond, "W" & inst.va)) > 0 Then
                NativeProcessInst = ind & "Do While " & NativeCondExpr(mn, True) & vbCrLf
                NVIndent = NVIndent + 1
                Exit Function
            End If
            If inst.jmpConst <= inst.va And NativeHas(NVLoopHdr, inst.jmpConst) Then
                'Back-edge to a loop header -> the bottom of a Do...Loop
                If NVIndent > 0 Then NVIndent = NVIndent - 1
                'Loop continues while the back-jump is taken.
                NativeProcessInst = NativeIndentStr() & "Loop While " & NativeCondExpr(mn, False) & vbCrLf
            ElseIf inst.jmpConst > inst.va Then
                'Forward conditional branch -> a structured If guarding the block
                'up to the branch target.  The block runs when the jump is NOT
                'taken, so negate the jump's relational.
                NativeProcessInst = ind & "If " & NativeCondExpr(mn, True) & " Then" & vbCrLf
                NativePushIf inst.jmpConst
            Else
                'Backward branch (not a recognised loop header) -> conditional GoTo
                '(the GoTo fires when the jump is taken).
                NativeProcessInst = ind & "If " & NativeCondExpr(mn, False) & " Then GoTo loc_" & Right$("00000000" & Hex$(inst.jmpConst), 8) & vbCrLf
            End If
            Exit Function

        Case C_JMP
            If Len(NativeColGet(NVWhileLoop, "W" & inst.va)) > 0 Then
                'Back-edge of a reconstructed Do While loop -> Loop.
                If NVIndent > 0 Then NVIndent = NVIndent - 1
                NativeProcessInst = NativeIndentStr() & "Loop" & vbCrLf
                Exit Function
            End If
            If inst.jmpConst <= inst.va And NativeHas(NVLoopHdr, inst.jmpConst) Then
                'Unconditional back-edge -> bottom of an infinite/top-tested Do...Loop
                If NVIndent > 0 Then NVIndent = NVIndent - 1
                NativeProcessInst = NativeIndentStr() & "Loop" & vbCrLf
            Else
                NativeProcessInst = ind & "GoTo loc_" & Right$("00000000" & Hex$(inst.jmpConst), 8) & vbCrLf
            End If
            Exit Function

        Case C_RET
            NativeProcessInst = ind & "Exit Sub" & vbCrLf
            Exit Function

        Case C_FLT
            'fnstsw ax (DF E0) / fstsw ax (9B DF E0) stores the FPU status word for
            'VB6's automatic overflow check (fnstsw; test al,imm; jcc <handler>).
            'Flag it so the following Jcc is dropped instead of becoming an If.
            Dim fdump As String
            fdump = UCase$(Replace(inst.dump, " ", ""))
            If Left$(fdump, 4) = "DFE0" Or Left$(fdump, 6) = "9BDFE0" Then NVFpuChk = True
            NativeFpuOp inst, mn
            Exit Function       'FPU ops build expressions silently

        Case C_PSH
            'Record the pushed value (immediate / string literal / local / reg)
            'as the next call argument, and as a single-value candidate.
            Dim pv As String
            NVLastPushDisp = 0
            pv = NativePushOperand(inst)
            If Len(pv) > 0 Then NVLastImm = pv: NativePushImm pv: NativeRecordPushDisp NVLastPushDisp
            'Keep a ring of the raw `push imm32/imm8` values so the __vbaNew object
            'creation can recover its Object Info pointer + destination-global args.
            Dim prawn As Long, praw As Long
            prawn = NativePushImmRaw(inst, praw)
            If prawn <> 0 Then
                NVRecentPush(NVRecentTop And 7) = praw
                NVRecentTop = NVRecentTop + 1
            End If
            Exit Function

        Case C_CMD
            'LEA records the local slot that a following property GET writes into
            If mn = "LEA" Then
                Dim ld As Long, lAbs As Boolean
                If NativeDecodeDisp(inst.dump, ld, lAbs) Then NVLastLea = ld: NVLastLeaSet = True
                Call NativeTrackReg(inst)
                Exit Function
            End If
            'mov [local], imm32 (opcode C7): record the value into the Variant-slot
            'map so a Variant built in this slot (a string-literal data field, or a
            'VT tag) can be read back when its address is passed by reference to
            'MsgBox / InputBox.
            If mn = "MOV" Then NativeTrackVariantStore inst
            'mov reg,[IAT] (opcode 8B) caches a runtime helper into a register.
            If mn = "MOV" And Left$(UCase$(Replace(inst.dump, " ", "")), 2) = "8B" Then
                Dim destReg As String
                destReg = NativeFirstReg(inst.command)
                If NativeRegIndex(destReg) >= 0 Then NativeSetRegImport destReg, NativeApiName(inst)
            End If
            'General GP-register value tracking; emits an assignment when a
            'recovered call result is stored to a local (else stays quiet).
            Dim asn As String
            asn = NativeTrackReg(inst)
            If Len(asn) > 0 Then NativeProcessInst = ind & asn & vbCrLf
            Exit Function

        Case Else
            Exit Function
    End Select

    NativeProcessInst = ind & "' " & Hex$(inst.va) & "  " & ann & vbCrLf
End Function

'---------------------------------------------------------------------------
' FPU / data-flow expression model
'---------------------------------------------------------------------------

Private Sub NativeFpuOp(inst As CInstruction, ByVal mn As String)
    Dim operand As String, a As String, disp As Long, isAbs As Boolean, hasMem As Boolean
    hasMem = NativeDecodeDisp(inst.dump, disp, isAbs)
    operand = NativeFpuOperand(hasMem, disp, isAbs)

    Select Case True
        Case mn = "FLD", mn = "FILD"
            NativeFpuPush operand
        Case mn = "FADD", mn = "FADDP", mn = "FIADD"
            a = NativeFpuPop(): NativeFpuPush "(" & a & " + " & operand & ")"
        Case mn = "FSUB", mn = "FSUBP", mn = "FISUB"
            a = NativeFpuPop(): NativeFpuPush "(" & a & " - " & operand & ")"
        Case mn = "FSUBR", mn = "FSUBRP"
            a = NativeFpuPop(): NativeFpuPush "(" & operand & " - " & a & ")"
        Case mn = "FMUL", mn = "FMULP", mn = "FIMUL"
            a = NativeFpuPop(): NativeFpuPush "(" & a & " * " & operand & ")"
        Case mn = "FDIV", mn = "FDIVP", mn = "FIDIV"
            a = NativeFpuPop(): NativeFpuPush "(" & a & " / " & operand & ")"
        Case mn = "FSTP", mn = "FST"
            a = NativeFpuPop()
            If hasMem And Not isAbs And disp < 0 Then
                NativeSetLocalExpr disp, a             'store into a local slot
            Else
                NVPendingArg = a                       'store into the outgoing arg area
            End If
        Case mn = "FCHS"
            a = NativeFpuPop(): NativeFpuPush "-" & a
        Case mn = "FCOM", mn = "FCOMP", mn = "FCOMPP", mn = "FICOM", mn = "FICOMP"
            'Comparison consumes the top; remember the operands as a hint for the
            'condition of the branch that follows.
            Dim lhs As String
            If NVFpuTop > 0 Then lhs = NativeFpuPop() Else lhs = "st0"
            NVLastCmp = "(" & lhs & " ? " & operand & ")"
    End Select
End Sub

Private Function NativeFpuOperand(ByVal hasMem As Boolean, ByVal disp As Long, ByVal isAbs As Boolean) As String
    'The memory operand of an FPU instruction: a global float constant
    '(absolute address), a known local slot, or a synthetic name.
    If Not hasMem Then
        NativeFpuOperand = "st0"
    ElseIf isAbs Then
        NativeFpuOperand = NativeFloatAtAddr(disp)
    ElseIf disp < 0 Then
        NativeFpuOperand = NativeGetLocalExpr(disp)
    Else
        NativeFpuOperand = "var_" & Hex$(disp)
    End If
End Function

Private Sub NativeFpuPush(ByVal s As String)
    If NVFpuTop > UBound(NVFpu) Then ReDim Preserve NVFpu(NVFpuTop + 16)
    NVFpu(NVFpuTop) = s
    NVFpuTop = NVFpuTop + 1
End Sub

Private Function NativeFpuPop() As String
    If NVFpuTop > 0 Then
        NVFpuTop = NVFpuTop - 1
        NativeFpuPop = NVFpu(NVFpuTop)
    Else
        NativeFpuPop = "st0"
    End If
End Function

Private Sub NativeSetLocalExpr(ByVal disp As Long, ByVal expr As String)
    Dim k As String
    k = "L" & disp
    On Error Resume Next
    NVLocal.Remove k
    On Error GoTo 0
    NVLocal.Add expr, k
End Sub

Private Function NativeGetLocalExpr(ByVal disp As Long) As String
    On Error Resume Next
    NativeGetLocalExpr = NVLocal("L" & disp)
    If Len(NativeGetLocalExpr) = 0 Then NativeGetLocalExpr = "var_" & Hex$(Abs(disp))
End Function

Private Function NativeFpuPopOrEmpty() As String
    'Pop the FPU expression top when present, else "" (the operand is lost).
    If NVFpuTop > 0 Then NativeFpuPopOrEmpty = NativeFpuPop()
End Function

Private Function NativeNumConvWrap(ByVal v As String, ByVal conv As String) As String
    'Wrap a value in a numeric conversion (CInt/CLng), collapsing a redundant inner
    'String->float conversion so `CInt(CDbl(x))` reads as `CInt(x)`.
    If Len(v) = 0 Then v = "<arg>"
    If (Left$(v, 5) = "CDbl(" Or Left$(v, 5) = "CSng(") And Right$(v, 1) = ")" Then
        v = Mid$(v, 6, Len(v) - 6)
    End If
    NativeNumConvWrap = conv & "(" & v & ")"
End Function

Private Function NativeIsCleanNamedVal(ByVal s As String) As Boolean
    'A clean tracked variable reference - a local, parameter, or module global
    '(optionally a deref chain like global_X(12)).  Deliberately conservative: bare
    'registers, numbers, control/property expressions and folded arithmetic are
    'excluded, so only an unambiguous Integer/Boolean variable test resolves.
    If Len(s) = 0 Then Exit Function
    NativeIsCleanNamedVal = (Left$(s, 4) = "var_") Or (Left$(s, 4) = "arg_") Or (Left$(s, 7) = "global_")
End Function

Private Function NativeIsStrOperand(ByVal s As String) As Boolean
    'True when s is a usable string-comparison operand: not empty, not the <arg>
    'placeholder, and not a bare number (which is an unresolved string pointer or a
    'lost value - a genuine string literal is quoted, a variable/expr is not numeric).
    If Len(s) = 0 Then Exit Function
    If s = "<arg>" Then Exit Function
    If IsNumeric(s) Then Exit Function
    NativeIsStrOperand = True
End Function

Private Sub NativeInvalidateLocalArg(ByVal argExpr As String)
    'A statement (Line Input #, Get #, Input #) wrote a local by reference.  The
    'slot's tracked value (often a stale zero-init constant the prologue stored)
    'is now wrong, so clear it - later reads then render the variable name
    'instead of the dead `0`.  Only plain `var_XXXX` locals are addressed here.
    If Left$(argExpr, 4) <> "var_" Then Exit Sub
    Dim hx As String, disp As Long
    hx = Mid$(argExpr, 5)
    On Error Resume Next
    disp = -CLng("&H" & hx)
    If Err.Number = 0 And disp < 0 Then NativeSetLocalExpr disp, ""
End Sub

Private Sub NativeSetVSlot(ByVal disp As Long, ByVal v As String)
    'Record the last value written to a 4-byte stack slot, so a Variant built in
    'that slot (a VT tag field plus a data field 8 bytes higher) can be read back
    'when its address is passed by reference to MsgBox / InputBox.
    Dim k As String
    k = "S" & disp
    On Error Resume Next
    NVVSlot.Remove k
    On Error GoTo 0
    NVVSlot.Add v, k
End Sub

Private Function NativeGetVSlot(ByVal disp As Long) As String
    On Error Resume Next
    NativeGetVSlot = NVVSlot("S" & disp)
End Function

Private Function NativeVariantVal(ByVal baseDisp As Long) As String
    'Resolve a Variant whose VT field is at baseDisp (its data field is 8 bytes
    'higher in memory, i.e. at baseDisp + 8).  Returns the held string / value, the
    'NV_MISSING sentinel for an omitted optional argument (VT_ERROR), or the bare
    'local name when the slot was never recognised as a Variant.
    Dim vt As String, dataV As String
    vt = NativeGetVSlot(baseDisp)
    If vt = "10" Then NativeVariantVal = NV_MISSING: Exit Function   'VT_ERROR -> missing optional
    dataV = NativeGetVSlot(baseDisp + 8)
    If Len(dataV) > 0 Then
        NativeVariantVal = dataV
    Else
        NativeVariantVal = "var_" & Hex$(Abs(baseDisp))
    End If
End Function

Private Sub NativeSetLocalGuid(ByVal disp As Long, ByVal guid As String)
    'Remember that a local stack slot holds a control object of a given GUID, so a
    'later `mov reg,[slot]; mov vt,[reg]; call [vt+off]` can resolve its property.
    Dim k As String
    k = "L" & disp
    On Error Resume Next
    NVLocalGuid.Remove k
    On Error GoTo 0
    If Len(guid) > 0 Then NVLocalGuid.Add guid, k
End Sub

Private Function NativeGetLocalGuid(ByVal disp As Long) As String
    On Error Resume Next
    NativeGetLocalGuid = NVLocalGuid("L" & disp)
End Function

'---------------------------------------------------------------------------
' Idiom helpers
'---------------------------------------------------------------------------

Private Function NativeProperty(ByVal vtOffset As Long) As String
'Resolve a property vtable call (call [obj + vtOffset]) on the last control.
    Dim p As String, propName As String, kind As String, valExpr As String
    If Len(NVLastGuid) = 0 Then Exit Function
    p = modPCode.GetProperty(NVLastGuid, vtOffset)
    If InStr(p, "Unknown GUID") > 0 Then Exit Function   'GUID/offset not in a loaded TypeLib - leave raw
    NativeSplitProp p, propName, kind
    If Len(propName) = 0 Then Exit Function

    Select Case kind
        Case "Get"
            'Result is written to the local addressed just before the call
            valExpr = NVLastControl & "." & propName
            If NVLastLeaSet Then NativeSetLocalExpr NVLastLea, valExpr
            NVLastLeaSet = False
            NativeProperty = "' get " & valExpr
        Case "Let", "Set"
            valExpr = NativePopValue()
            NativeProperty = NVLastControl & "." & propName & " = " & valExpr
            NativeResetValue
        Case Else
            NativeProperty = "' " & NVLastControl & "." & propName & "()"
    End Select
End Function

Private Function NativeControlProp(ByVal ctlName As String, ByVal guid As String, ByVal vtOffset As Long) As String
'Resolve a property call on a tracked control object (its vtable carries the
'control GUID).  Like NativeProperty but takes the control/guid explicitly (from
'the per-register tracking) and reads a Let's value from the pushed argument
'rather than NVReg(0) (which holds the control object itself, not the value).
    Dim p As String, propName As String, kind As String, valExpr As String
    If Len(guid) = 0 Then Exit Function
    p = modPCode.GetProperty(guid, vtOffset)
    If InStr(p, "Unknown GUID") > 0 Then Exit Function   'GUID/offset not in a loaded TypeLib - leave raw
    NativeSplitProp p, propName, kind
    If Len(propName) = 0 Then Exit Function

    Select Case kind
        Case "Get"
            valExpr = ctlName & "." & propName
            If NVLastLeaSet Then NativeSetLocalExpr NVLastLea, valExpr
            NVLastLeaSet = False
            NativeControlProp = "' get " & valExpr
        Case "Let", "Set"
            'Property Let compiles to `push value; push this; call [vt+off]`.  The TOP
            'push is the control `this` itself - never the value - so take the push
            'just below it (or a pending RGB()/FPU value).  Do NOT fall back to
            'NVReg(0)/NVLastImm: those hold the control object and would leak it as
            'the right-hand side.  Leave a placeholder when no value was tracked.
            If Len(NVPendingArg) > 0 Then
                valExpr = NVPendingArg
            ElseIf NVFpuTop > 0 Then
                valExpr = NativeFpuPop()
            ElseIf NVPushTop >= 2 Then
                valExpr = NVPushImm(NVPushTop - 2)
            End If
            If Len(valExpr) = 0 Then valExpr = "<value>"
            NativeControlProp = ctlName & "." & propName & " = " & valExpr
            NativeResetValue
        Case Else
            'A control METHOD (Refresh, Cls, Line, Move, SetFocus, ...).  Emit a
            'real statement with its arguments instead of a dropped comment.  The
            'vtable call pushes the control `this` last (so it is the first
            'source-order arg) - drop it when present; the rest are the method args.
            Dim mArgs As String, mParts() As String, mk As Long, mOut As String, mLo As Long, mN As Long
            mArgs = NativeArgList()
            mLo = 0
            If Len(mArgs) > 0 Then
                mParts = Split(mArgs, ", ")
                If mParts(0) = ctlName Then mLo = 1           'drop the leading `this`
                mN = UBound(mParts) - mLo + 1                 'real argument count
            End If
            'The graphics methods use a special coordinate syntax that the flat
            'vtable argument order does not.  Line's params are
            '[flags, x1, y1, x2, y2, color] -> `Line (x1, y1)-(x2, y2), color`.
            If propName = "Line" And mN >= 6 Then
                NativeControlProp = ctlName & ".Line (" & mParts(mLo + 1) & ", " & mParts(mLo + 2) & ")-(" _
                                  & mParts(mLo + 3) & ", " & mParts(mLo + 4) & "), " & mParts(mLo + 5)
            ElseIf propName = "PSet" And mN >= 4 Then
                NativeControlProp = ctlName & ".PSet (" & mParts(mLo + 1) & ", " & mParts(mLo + 2) & "), " & mParts(mLo + 3)
            Else
                For mk = mLo To mLo + mN - 1
                    If Len(mOut) > 0 Then mOut = mOut & ", "
                    mOut = mOut & mParts(mk)
                Next
                If Len(mOut) > 0 Then
                    NativeControlProp = ctlName & "." & propName & " " & mOut
                Else
                    NativeControlProp = ctlName & "." & propName
                End If
            End If
    End Select
End Function

Private Function NativePopValue() As String
    'The right-hand side of an assignment, in priority order.
    If Len(NVPendingArg) > 0 Then
        NativePopValue = NVPendingArg
    ElseIf NVFpuTop > 0 Then
        NativePopValue = NativeFpuPop()
    ElseIf Len(NVReg(0)) > 0 Then
        NativePopValue = NVReg(0)
    ElseIf Len(NVLastImm) > 0 Then
        NativePopValue = NVLastImm
    Else
        NativePopValue = "<value>"
    End If
End Function

Private Sub NativeResetValue()
    NVPendingArg = "": NVLastImm = "": NVFpuTop = 0: NVReg(0) = ""
End Sub

Private Function NativeRuntimeCall(inst As CInstruction, ByVal apiName As String) As String
'Map an msvbvm60 / VB runtime helper to a VB statement, a folded value (left in
'eax / NVReg(0) for the consumer), or a dropped no-op.  Sets NVCallHandled.
    Dim nm As String, vbName As String, arity As Long, isStmt As Boolean, args As String, aa As String, bb As String
    NVCallHandled = False
    NVKeepPushStack = False
    nm = apiName
    If Len(nm) = 0 Then Exit Function
    NVCallHandled = True

    'Internal __vba* helpers handled specially
    Select Case True
        Case InStr(nm, "__vbaHresultCheckObj") > 0
            NativeRuntimeCall = "": Exit Function           'automatic error check - drop
        Case InStr(nm, "__vbaObjSet") > 0, InStr(nm, "__vbaObjSetAddref") > 0
            'Object store: __vbaObjSet(&dest, srcObj).  When the source is a tracked
            'control object (eax still carries the accessor result's identity),
            'remember the destination local's control GUID - so a later property
            'access through that local resolves - and surface the Set statement.
            '__vbaObjSet returns its source object in eax.  When the source is the
            'Err object (rtcErrObj just ran, even if a `lea eax,[dest]` clobbered the
            'tag in between), re-tag eax so the following `mov reg,[eax]; call
            '[reg+0x1C]` chain resolves to Err.Number / Err.Description.
            If NVErrObjPending Then
                NVErrObjPending = False
                NVPushTop = 0
                NVReg(0) = "Err": NVRegObjType(0) = "Err": NVRegObjGuid(0) = ""
                NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                NVRegObjVt(0) = "": NVRegObjVtGuid(0) = ""
                NativeRuntimeCall = "": Exit Function
            End If
            Dim osGuid As String, osName As String
            osGuid = NVRegObjGuid(0): osName = NVRegObjType(0)
            'The source control's identity is often cleared off eax by a `lea eax,
            '[dest]` placed just before the call; fall back to the last control
            'accessor (the __vbaObjSet source is always the just-fetched control).
            If Len(osGuid) = 0 And Len(NVLastGuid) > 0 Then osGuid = NVLastGuid: osName = NVLastControl
            NVPushTop = 0
            'Re-tag eax with the control identity: __vbaObjSet returns the same
            'object, so a following `mov [tempLocal], eax` (the property LET target)
            'keeps the GUID and the LET through that local resolves.
            If Len(osGuid) > 0 Then
                NVReg(0) = osName: NVRegObjType(0) = osName: NVRegObjGuid(0) = osGuid
                NVRegObjVt(0) = "": NVRegObjVtGuid(0) = ""
            End If
            If NVLastLeaSet And Len(osGuid) > 0 Then
                NativeSetLocalGuid NVLastLea, osGuid
                NativeSetLocalExpr NVLastLea, osName
                NativeRuntimeCall = "Set var_" & Hex$(Abs(NVLastLea)) & " = " & osName
                NVLastLeaSet = False
                Exit Function
            End If
            NativeRuntimeCall = "": Exit Function    'object-store plumbing - drop
        Case InStr(nm, "__vbaOnError") > 0
            'On Error setup. The handler label is recovered structurally (the
            'block after the VB error epilogue) rather than from this call's args.
            NVPushTop = 0: NativeResetValue
            If NVErrHandler <> 0 Then
                NativeRuntimeCall = "On Error GoTo loc_" & Right$("00000000" & Hex$(NVErrHandler), 8)
            Else
                NativeRuntimeCall = "On Error GoTo <handler>"
            End If
            Exit Function
        Case nm = "Err", InStr(nm, "rtcErrObj") > 0
            'rtcErrObj returns the VB6 Err object in eax.  Tag the register so the
            'following vtable property GETs (call [ErrVt + 0x1C/0x2C]) resolve to
            'Err.Number / Err.Description; emit nothing.
            NVPushTop = 0
            NVReg(0) = "Err": NVRegObjType(0) = "Err": NVRegObjGuid(0) = ""
            NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
            NVRegObjVt(0) = "": NVRegObjVtGuid(0) = ""
            NVErrObjPending = True   'the following __vbaObjSet stores this Err object
            NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaStrI4") > 0, InStr(nm, "__vbaStrI2") > 0, _
             InStr(nm, "__vbaStrR4") > 0, InStr(nm, "__vbaStrR8") > 0, _
             InStr(nm, "__vbaStrBool") > 0, InStr(nm, "__vbaStrDate") > 0
            'Numeric/typed value -> its string form (one stack argument).  Used to
            'render `& number &` operands in string concatenations.  Folds to eax.
            aa = NativeArgPop()
            If Len(aa) = 0 Then aa = "<arg>"
            NVReg(0) = "CStr(" & aa & ")": NVKeepPushStack = True
            NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaR8Str") > 0, InStr(nm, "__vbaR4Str") > 0
            'String -> floating point (the reverse of __vbaStrR8).  The result is
            'returned on the FPU stack, so push it into the FPU model; a following
            'narrowing (__vbaR?IntI2/I4) or an fstp store consumes it.  Folds the
            'whole `<field> = Int(Trim$(...))` config-parse assignment.
            aa = NativeArgPop()
            If Len(aa) = 0 Then aa = "<arg>"
            NativeFpuPush IIf(InStr(nm, "R4Str") > 0, "CSng(", "CDbl(") & aa & ")"
            NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaR8IntI4") > 0, InStr(nm, "__vbaR4IntI4") > 0
            'Floating value (FPU top) -> Long.  Folds to eax for the store.
            NVReg(0) = NativeNumConvWrap(NativeFpuPopOrEmpty(), "CLng")
            NVRegIsAddr(0) = False: NVRegObjType(0) = "": NVRegObjVt(0) = ""
            NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaR8IntI2") > 0, InStr(nm, "__vbaR4IntI2") > 0
            'Floating value (FPU top) -> Integer.  Folds to eax for the store.
            NVReg(0) = NativeNumConvWrap(NativeFpuPopOrEmpty(), "CInt")
            NVRegIsAddr(0) = False: NVRegObjType(0) = "": NVRegObjVt(0) = ""
            NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaI4Str") > 0
            'String -> Long (register-based, result in eax - no FPU).
            aa = NativeArgPop()
            If Len(aa) = 0 Then aa = "<arg>"
            NVReg(0) = "CLng(" & aa & ")"
            NVRegIsAddr(0) = False: NVRegObjType(0) = "": NVRegObjVt(0) = ""
            NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaI2Str") > 0
            'String -> Integer (register-based, result in eax - no FPU).
            aa = NativeArgPop()
            If Len(aa) = 0 Then aa = "<arg>"
            NVReg(0) = "CInt(" & aa & ")"
            NVRegIsAddr(0) = False: NVRegObjType(0) = "": NVRegObjVt(0) = ""
            NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaStrCat") > 0
            '__vbaStrCat(p1, p2) returns p1 & p2.  p1 is pushed first (deeper on the
            'argument stack) and p2 last (on top), so the deeper operand is the LEFT
            'side of the concatenation: pop top (p2) then deeper (p1) and join p1 & p2.
            aa = NativeArgPop(): bb = NativeArgPop()
            NVReg(0) = NativeConcat(bb, aa): NVKeepPushStack = True
            NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaStrToAnsi") > 0, InStr(nm, "__vbaStrToUnicode") > 0
            'Charset conversion: StrToAnsi(dst, src) returns the converted string
            'in eax.  For decompilation the value is just the source string, so
            'fold it through to the consumer (e.g. an API argument).
            Dim adst As String, asrc As String
            adst = NativeArgPop(): asrc = NativeArgPop()
            If Len(asrc) = 0 Then asrc = adst
            NVReg(0) = asrc: NVPushTop = 0: NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaStrCmp") > 0, InStr(nm, "__vbaStrComp") > 0
            '__vbaStrCmp(p1, p2) returns strcmp(p1, p2); p1 is pushed deeper, p2 on
            'top.  When the pre-pass found the boolean materialisation that turns the
            'tri-state into (strcmp = 0), bind the equality relational "(p1 = p2)"
            'into the materialisation's target register so the following test/jcc
            'renders `If p1 = p2`.  Otherwise leave it as a visible Call.
            Dim scReg As String, scA() As String, scN As Long
            scReg = NativeColGet(NVStrCmpReg, "P" & inst.va)
            NativeArgsSnapshot scA, scN          'a(0) = top of arg stack (= p2)
            If Len(scReg) > 0 And scN >= 2 Then
                Dim scTop As String, scDeep As String
                scTop = scA(0)                   'p2 (top of the arg stack)
                scDeep = scA(1)                  'p1 (deeper)
                'Both operands must be resolved string values.  A bare number is an
                'unresolved string pointer (a real literal would be quoted) or a lost
                'value - emitting `(0 = 4218532)` is meaningless, so fall back to a
                'visible Call for those, matching the pre-change output.
                If NativeIsStrOperand(scTop) And NativeIsStrOperand(scDeep) Then
                    'Place the relational in eax; the materialisation's `mov REG,eax`
                    '(left in place) copies it on to the tested register, and the
                    'neg/sbb/inc/neg are suppressed so they don't clobber it.
                    NVReg(0) = "(" & scDeep & " = " & scTop & ")"
                    NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                    NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjGuid(0) = "": NVRegObjVtGuid(0) = ""
                    NativeRuntimeCall = "": Exit Function
                End If
            End If
            'No materialisation, or an operand we could not resolve - render as the
            'prior visible Call (reproducing the pre-change output exactly).
            Dim scList As String, scK As Long
            For scK = 0 To scN - 1
                If Len(scList) > 0 Then scList = scList & ", "
                scList = scList & scA(scK)
            Next
            NativeRuntimeCall = "Call " & NativeFriendlyName(nm) & "(" & scList & ")": Exit Function
        Case InStr(nm, "__vbaObjIs") > 0
            'Object identity: __vbaObjIs(p1, p2) returns p1 Is p2 (a Boolean in ax,
            'no materialisation - VB tests it directly).  Bind the relational into eax
            'so the following `test ax,ax`/jcc renders `If a Is Nothing`.  A 0 operand
            'is Nothing; order it as `<object> Is Nothing` for readability.
            Dim oiA() As String, oiN As Long, oiObj As String, oiOther As String, oiTmp As String
            NativeArgsSnapshot oiA, oiN
            If oiN >= 2 Then
                oiObj = oiA(0)                   'top of the arg stack
                oiOther = oiA(1)                 'deeper
                If oiObj = "0" And oiOther <> "0" Then oiTmp = oiObj: oiObj = oiOther: oiOther = oiTmp
                If oiOther = "0" Then oiOther = "Nothing"
                If oiObj = "0" Then oiObj = "Nothing"
                If Len(oiObj) > 0 And oiObj <> "<arg>" And Len(oiOther) > 0 And oiOther <> "<arg>" Then
                    NVReg(0) = "(" & oiObj & " Is " & oiOther & ")"
                    NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                    NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjGuid(0) = "": NVRegObjVtGuid(0) = ""
                    NativeRuntimeCall = "": Exit Function
                End If
            End If
            'Unresolved operand - render as the prior visible Call.
            Dim oiList As String, oiK As Long
            For oiK = 0 To oiN - 1
                If Len(oiList) > 0 Then oiList = oiList & ", "
                oiList = oiList & oiA(oiK)
            Next
            NativeRuntimeCall = "Call " & NativeFriendlyName(nm) & "(" & oiList & ")": Exit Function
        Case InStr(nm, "__vbaI2I4") > 0, InStr(nm, "__vbaUI1I2") > 0, _
             InStr(nm, "__vbaUI1I4") > 0, InStr(nm, "__vbaI4UI1") > 0, _
             InStr(nm, "__vbaI2UI1") > 0
            'Implicit integer widening/narrowing.  Register-based (arg and result in
            'eax), so the value already sits in NVReg(0) - just suppress the Call.
            'Do NOT touch the push stack: it belongs to the FOLLOWING consumer call
            '(clearing it dropped EOF/Close/UBound arguments).
            NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaFpR4") > 0, InStr(nm, "__vbaFpR8") > 0, _
             InStr(nm, "__vbaFpI4") > 0, InStr(nm, "__vbaFpI2") > 0, _
             InStr(nm, "__vbaFpUI1") > 0, InStr(nm, "__vbaFpCY") > 0
            'Convert the FPU top to an integer/real type (result in eax).  Fold the
            'FPU expression value through to eax; emit no Call.  The push stack is
            'left untouched (the FPU stack, not the arg stack, holds the input).
            If NVFpuTop > 0 Then NVReg(0) = NativeFpuPop()
            NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaCastObj") > 0
            'Object cast / interface coercion - the result is the same object, which
            'is already tracked in eax (NVReg(0)).  Suppress the Call; leave the push
            'stack alone.
            NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaVarDup") > 0
            '__vbaVarDup(dest /*ecx*/, src /*edx*/) duplicates a Variant.  VB6 uses
            'it to copy a freshly built BSTR/literal Variant into the by-reference
            'argument slot of an MsgBox / InputBox call, so propagate the tracked
            'slot values dest <- src (both the VT field and the data field 8 bytes
            'higher) for later argument reconstruction.
            If NVRegIsAddr(1) And NVRegIsAddr(2) Then
                Dim vdDest As Long, vdSrc As Long
                vdDest = NVRegAddrDisp(1): vdSrc = NVRegAddrDisp(2)
                NativeSetVSlot vdDest, NativeGetVSlot(vdSrc)
                NativeSetVSlot vdDest + 8, NativeGetVSlot(vdSrc + 8)
            End If
            NVPushTop = 0: NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaStrMove") > 0, InStr(nm, "__vbaStrCopy") > 0, _
             InStr(nm, "__vbaVarMove") > 0, InStr(nm, "__vbaVarCopy") > 0, _
             InStr(nm, "__vbaStrVarMove") > 0
            'Move/copy of a computed value into a local -> a VB assignment.
            NativeRuntimeCall = NativeMoveAssign(): Exit Function
        Case InStr(nm, "__vbaNew2") > 0, InStr(nm, "__vbaNew") > 0
            'Object creation.  VB pushes a pointer to the object's Object Info; for
            'an `As New` auto-instantiation it ALSO pushes the address of the global
            'that holds the instance.  Recover the class name from the Object Info
            'chain (ObjectInfo+0x18 -> Public Object Descriptor +0x18 -> name) and,
            'for the auto-instantiation, remember that the global is of that class so
            'its later method calls resolve.
            Dim ncls As String, ngObj As Long, rp As Long, rv As Long
            For rp = 0 To 7
                rv = NVRecentPush(rp)
                If rv <> 0 And Len(ncls) = 0 Then ncls = NativeClassFromObjInfo(rv)
                If rv <> 0 And NativeIsGlobalAddr(rv) Then ngObj = rv
            Next
            NVPushTop = 0
            If Len(ncls) > 0 Then
                If ngObj <> 0 Then NativeColPut NVObjClass, "G" & ngObj, ncls
                NVReg(0) = "New " & ncls: NVRegObjType(0) = ncls
            Else
                NVReg(0) = "New (object)"
            End If
            NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaExitProc") > 0, InStr(nm, "__vbaErrorOverflow") > 0, _
             InStr(nm, "__vbaError") > 0, InStr(nm, "__vbaErrVar") > 0, _
             InStr(nm, "__vbaFree") > 0, InStr(nm, "__vbaVarDup") > 0, _
             InStr(nm, "__vbaAryLock") > 0, InStr(nm, "__vbaAryUnlock") > 0, _
             InStr(nm, "__vbaAryDestruct") > 0, InStr(nm, "__vbaAryConstruct") > 0, _
             InStr(nm, "__vbaGenerateBoundsError") > 0, _
             InStr(nm, "__vbaExceptHandler") > 0, InStr(nm, "__vbaSetSystemError") > 0
            NVPushTop = 0: NativeRuntimeCall = "": Exit Function       'silent

        '--- Array intrinsics ---
        Case InStr(nm, "__vbaUbound") > 0, InStr(nm, "__vbaLbound") > 0
            'UBound/LBound(array[, dim]).  Value-returning -> folds into eax.  Args
            'in source order are [dim, arrayPtr]; the array is the deepest push.
            Dim ubA() As String, ubN As Long, ubArr As String, ubDim As String
            NativeArgsSnapshot ubA, ubN
            If ubN >= 1 Then
                ubArr = ubA(ubN - 1)
                If ubN >= 2 Then If ubA(0) <> "1" And Len(ubA(0)) > 0 Then ubDim = ", " & ubA(0)
                NVReg(0) = IIf(InStr(nm, "Ubound") > 0, "UBound", "LBound") & "(" & ubArr & ubDim & ")"
            End If
            NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaRedimPreserve") > 0, InStr(nm, "__vbaRedim") > 0
            'ReDim [Preserve] array(lb To ub, ...).  Args: [elemFlags, elemSize,
            'arrayPtr, saFlags, nDims, (ub,lb) per dimension].
            Dim rdA() As String, rdN As Long
            NativeArgsSnapshot rdA, rdN
            NativeRuntimeCall = NativeRedimStmt(rdA, rdN, (InStr(nm, "Preserve") > 0))
            Exit Function
        Case InStr(nm, "__vbaEraseKeepData") > 0
            'Internal bookkeeping emitted just before a ReDim Preserve - drop.
            NVPushTop = 0: NativeRuntimeCall = "": Exit Function

        '--- File I/O statements ---
        Case InStr(nm, "__vbaFileOpen") > 0
            'Open <path> For <mode> As #<n> [Len = <rl>].  Args (source order):
            '[mode, reclen, filenum, pathname].
            Dim foA() As String, foN As Long
            NativeArgsSnapshot foA, foN
            NativeRuntimeCall = NativeFileOpenStmt(foA, foN)
            Exit Function
        Case InStr(nm, "__vbaFileClose") > 0
            Dim fcA() As String, fcN As Long
            NativeArgsSnapshot fcA, fcN
            If fcN >= 1 And Len(fcA(0)) > 0 Then NativeRuntimeCall = "Close #" & fcA(0) Else NativeRuntimeCall = "Close"
            Exit Function
        Case InStr(nm, "__vbaPutOwner") > 0, InStr(nm, "__vbaPut3") > 0, InStr(nm, "__vbaPut4") > 0
            'Put #<filenum>, , <var>.  Args: [descriptor/size, var, filenum].
            Dim ptA() As String, ptN As Long
            NativeArgsSnapshot ptA, ptN
            If ptN >= 3 Then NativeRuntimeCall = "Put #" & ptA(2) & ", , " & ptA(1)
            Exit Function
        Case InStr(nm, "__vbaGetOwner") > 0, InStr(nm, "__vbaGet3") > 0, InStr(nm, "__vbaGet4") > 0
            'Get #<filenum>, , <var>.  Args: [descriptor/size, var, filenum].
            Dim gtA() As String, gtN As Long
            NativeArgsSnapshot gtA, gtN
            If gtN >= 3 Then
                NativeRuntimeCall = "Get #" & gtA(2) & ", , " & gtA(1)
                NativeInvalidateLocalArg gtA(1)     'the read-into local is now live, not its zero-init
            End If
            Exit Function
        Case InStr(nm, "__vbaLineInputStr") > 0
            'Line Input #<filenum>, <var>.  Args: [var, filenum].
            Dim liA() As String, liN As Long
            NativeArgsSnapshot liA, liN
            If liN >= 2 Then
                NativeRuntimeCall = "Line Input #" & liA(1) & ", " & liA(0)
                NativeInvalidateLocalArg liA(0)     'the read-into local is now live, not its zero-init
            End If
            Exit Function
    End Select

    'MsgBox / InputBox: arguments are passed as by-reference Variant pointers
    '(lea reg,[ebp-X]; push reg).  Resolve each pointer to the Variant it built
    'earlier (a string literal, an expression, or a missing optional) rather than
    'showing the raw temporary local.
    If InStr(nm, "rtcMsgBox") > 0 Or InStr(UCase$(nm), "MSGBOX") > 0 Then
        'Emitted directly as a statement: MsgBox is normally used in statement form,
        'and its result is followed by VB's Variant cleanup (lea eax,[..]; push eax),
        'which the deferred-call fold would mistake for the result being consumed.
        Dim mbArgs As String
        mbArgs = NativeVariantArgList(True)
        If Len(mbArgs) > 0 Then NativeRuntimeCall = "MsgBox " & mbArgs Else NativeRuntimeCall = "MsgBox"
        NVPushTop = 0: Exit Function
    End If
    If InStr(nm, "rtcInputBox") > 0 Or InStr(UCase$(nm), "INPUTBOX") > 0 Then
        'Value-returning: the result is normally assigned (var = InputBox(...)).
        NVReg(0) = "InputBox(" & NativeVariantArgList(False) & ")"
        NVPendingCall = "Call " & NVReg(0)
        NativeRuntimeCall = "": Exit Function
    End If
    If InStr(UCase$(nm), "RGB") > 0 Then
        'Result feeds a property Let; keep it in the pending-value channel so it
        'survives eax being overwritten before the Let call.
        NVPendingArg = "RGB(" & NativeArgsN(3) & ")": NativeRuntimeCall = "": Exit Function
    End If

    'Friendly VB-intrinsic descriptions from the API database already read as the
    'VB name with trailing "()" (e.g. "Environ$()", "QBColor()", "InputBox()").
    'Strip the parens and render with the real argument list.
    If Right$(nm, 2) = "()" Then
        vbName = Left$(nm, Len(nm) - 2)
        Select Case UCase$(vbName)
            Case "DOEVENTS", "BEEP", "RANDOMIZE", "STOP", "END", "KILL", "MKDIR", _
                 "RMDIR", "CHDIR", "CHDRIVE", "FILECOPY", "SAVESETTING", "DELETESETTING", _
                 "SAVEPICTURE", "SENDKEYS", "APPACTIVATE", "RESET", "SETATTR", _
                 "OPEN", "PRINT", "CLOSE", "WRITE", "PUT", "GET", "SEEK", "LOCK", _
                 "UNLOCK", "NAME", "WIDTH", "LINE INPUT", "INPUT"
                NativeRuntimeCall = Trim$(vbName & " " & NativeArgList()): Exit Function
            Case "LEFT", "LEFT$", "RIGHT", "RIGHT$", "MID", "MID$", "TRIM", "TRIM$", _
                 "LTRIM", "LTRIM$", "RTRIM", "RTRIM$", "UCASE", "UCASE$", "LCASE", "LCASE$"
                'Pure value-returning string transforms.  These almost always feed a
                'following call / concat / comparison (string-parsing idioms like
                'UCase$(Left$(s, n)) = "TAG="), so fold the result into eax AND defer
                'a "Call X()" statement: if the next instruction consumes eax the
                'value threads into the consumer; otherwise the deferred call is
                'flushed as a statement (NVPendingCall fold/flush in the decode loop),
                'so an unconsumed result is never silently dropped.
                Dim stArgs As String
                stArgs = NativeArgList()
                NVReg(0) = vbName & "(" & stArgs & ")"
                NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjGuid(0) = "": NVRegObjVtGuid(0) = ""
                NVPendingCall = "Call " & NVReg(0)
                NativeRuntimeCall = "": Exit Function
            Case Else
                'Emit each intrinsic as its own visible statement.  (Threading the
                'result into the next op silently dropped calls such as DateAdd /
                'DateValue / CDate; the __vba* conversions and StrCat still flow,
                'so the concat/assignment reconstruction is unaffected.)
                NativeRuntimeCall = "Call " & vbName & "(" & NativeArgList() & ")": Exit Function
        End Select
    End If

    'Table-driven symbolization (mostly __vba conversions)
    NativeRuntimeSyntax nm, vbName, arity, isStmt
    If Len(vbName) > 0 Then
        If arity >= 0 Then args = NativeArgsN(arity) Else args = NativeArgList()
        If isStmt Then
            If Len(args) > 0 Then NativeRuntimeCall = vbName & " " & args Else NativeRuntimeCall = vbName
        Else
            NVReg(0) = vbName & "(" & args & ")"     'value-returning -> flows to consumer
            NativeRuntimeCall = ""
        End If
        Exit Function
    End If

    'Unknown helper: emit a visible Call with whatever arguments were collected
    NativeRuntimeCall = "Call " & NativeFriendlyName(nm) & "(" & NativeArgList() & ")"
End Function

Private Function NativeMoveAssign() As String
    'A move/copy runtime helper stores a value into the local addressed by the
    'most recent LEA.  Surface it as "var_X = <value>" when the value is worth
    'showing (a call/concat, or a string literal), and leave the local's name in
    'eax (the helpers return the moved value) so a following use references it.
    'Source priority: eax (computed expressions), the FPU pending arg, then edx
    '(the register-argument form used by StrCopy for `var = "literal"`).
    Dim dn As String, src As String
    src = NVReg(0)
    'A quoted string literal in edx (the fastcall source register) is an explicit
    '`var = "literal"` StrCopy source; it takes priority over a stale eax left by a
    'preceding call (e.g. a __vbaStrCmp whose result was never cleared off eax).
    If Left$(NVReg(2), 1) = Chr$(34) Then src = NVReg(2)
    If Not NativeIsExprValue(src) And Len(NVPendingArg) > 0 Then src = NVPendingArg
    If Not NativeIsExprValue(src) And NativeIsExprValue(NVReg(2)) Then src = NVReg(2)
    If NVLastLeaSet And NativeIsExprValue(src) Then
        dn = "var_" & Hex$(Abs(NVLastLea))
        NativeMoveAssign = dn & " = " & src
        NativeSetLocalExpr NVLastLea, dn
        NVReg(0) = dn                       'the helper returns the moved value in eax
    End If
    'These move/copy helpers are fastcall (ecx/edx) and take no stack arguments, so
    'pending pushes belong to an enclosing expression (e.g. a string-literal operand
    'awaiting a later __vbaStrCat) and must be preserved, not discarded.
    NVLastLeaSet = False: NVPendingArg = "": NVKeepPushStack = True
End Function

Private Function NativeIsExprValue(ByVal s As String) As Boolean
    'A value worth surfacing in an assignment: a call/concat (parenthesised), a
    'string literal, or a local variable (incl. a plain copy).  Bare numbers /
    'register names are treated as not worth it.
    If Len(s) = 0 Then Exit Function
    If InStr(s, "(") > 0 Then NativeIsExprValue = True: Exit Function
    If Left$(s, 1) = Chr$(34) Then NativeIsExprValue = True: Exit Function
    If Left$(s, 4) = "var_" Then NativeIsExprValue = True
End Function

Private Function NativeConcat(ByVal aa As String, ByVal bb As String) As String
    'Build "a & b", dropping a null operand (VB pushes vbNullString as 0, which
    'would otherwise render as the spurious `0 & "x"` / `"x" & 0`).  An untracked
    '"<arg>" operand is kept - it stands for a real value we could not recover.
    Dim a As Boolean, b As Boolean
    a = (Len(aa) > 0 And aa <> "0")
    b = (Len(bb) > 0 And bb <> "0")
    If a And b Then
        NativeConcat = "(" & aa & " & " & bb & ")"
    ElseIf a Then
        NativeConcat = aa
    ElseIf b Then
        NativeConcat = bb
    Else
        NativeConcat = Chr$(34) & Chr$(34)      'empty string ""
    End If
End Function

Private Sub NativeRuntimeSyntax(ByVal nm As String, ByRef vbName As String, ByRef arity As Long, ByRef isStmt As Boolean)
'Map a runtime-helper name to a VB function/statement, its argument arity
'(-1 = variadic: take all pending args), and whether it is a statement.
    vbName = "": arity = -1: isStmt = False
    Select Case nm
        '--- type conversions (value, 1 arg) ---
        Case "__vbaI2": vbName = "CInt": arity = 1
        Case "__vbaI4": vbName = "CLng": arity = 1
        Case "__vbaR4": vbName = "CSng": arity = 1
        Case "__vbaR8": vbName = "CDbl": arity = 1
        Case "__vbaCy": vbName = "CCur": arity = 1
        Case "__vbaUI1": vbName = "CByte": arity = 1
        Case "__vbaBool", "__vbaBoolVar", "__vbaBoolVarNull": vbName = "CBool": arity = 1
        Case "__vbaStrI2", "__vbaStrI4", "__vbaStrR4", "__vbaStrR8", "__vbaStrCy", _
             "__vbaStrBool", "__vbaStrVarVal", "__vbaStrVarMove", "__vbaStrVarCopy": vbName = "CStr": arity = 1
        '--- string functions (value) ---
        Case "__vbaLenBstr", "__vbaLenVar": vbName = "Len": arity = 1
        Case "rtcLeftCharBstr", "rtcLeftCharVar", "rtcLeftBstr": vbName = "Left$": arity = 2
        Case "rtcRightCharBstr", "rtcRightCharVar", "rtcRightBstr": vbName = "Right$": arity = 2
        Case "rtcMidCharBstr", "rtcMidCharVar", "rtcMidBstr": vbName = "Mid$": arity = -1
        Case "rtcUpperCaseBstr", "rtcUpperCaseVar": vbName = "UCase$": arity = 1
        Case "rtcLowerCaseBstr", "rtcLowerCaseVar": vbName = "LCase$": arity = 1
        Case "rtcTrimBstr", "rtcTrimVar": vbName = "Trim$": arity = 1
        Case "rtcLeftTrimBstr", "rtcLeftTrimVar": vbName = "LTrim$": arity = 1
        Case "rtcRightTrimBstr", "rtcRightTrimVar": vbName = "RTrim$": arity = 1
        Case "rtcSpaceBstr", "rtcSpaceVar": vbName = "Space$": arity = 1
        Case "__vbaStrComp", "__vbaStrCompVar": vbName = "StrComp": arity = -1
        Case "rtcInStr", "rtcInStrVar": vbName = "InStr": arity = -1
        Case "rtcInStrRev": vbName = "InStrRev": arity = -1
        Case "rtcReplace": vbName = "Replace": arity = -1
        Case "rtcStrReverse": vbName = "StrReverse": arity = 1
        Case "rtcStringBstr", "rtcStringVar": vbName = "String$": arity = 2
        Case "rtcAscChar", "rtcAnsiValueBstr": vbName = "Asc": arity = 1
        Case "rtcChrBstr", "rtcChrVar": vbName = "Chr$": arity = 1
        Case "rtcVarStrFromVar", "rtcFormatVar", "__vbaStrFormatVar": vbName = "Format$": arity = -1
        '--- math (value) ---
        Case "rtcAbsVar": vbName = "Abs": arity = 1
        Case "rtcSgn": vbName = "Sgn": arity = 1
        Case "rtcSqr": vbName = "Sqr": arity = 1
        Case "rtcInt", "rtcFix": vbName = "Int": arity = 1
        '--- date/time + type info (value) ---
        Case "rtcGetTimer": vbName = "Timer": arity = 0
        Case "rtcGetDateVar", "rtcGetPresentDate": vbName = "Now": arity = 0
        Case "rtcTypeName": vbName = "TypeName": arity = 1
        Case "rtcVarType": vbName = "VarType": arity = 1
        Case "rtcEnvironBstr", "rtcEnvironVar": vbName = "Environ$": arity = 1
        Case "rtcFreeFile": vbName = "FreeFile": arity = -1
        '--- statements ---
        Case "rtcDoEvents": vbName = "DoEvents": arity = 0: isStmt = True
        Case "rtcBeep": vbName = "Beep": arity = 0: isStmt = True
        Case "rtcShell": vbName = "Shell": arity = -1
        Case "rtcFileOpen", "__vbaFileOpen": vbName = "Open": arity = -1: isStmt = True
        Case "__vbaFileClose", "rtcFileClose": vbName = "Close": arity = -1: isStmt = True
        Case "rtcPrintFile", "__vbaPrintFile": vbName = "Print": arity = -1: isStmt = True
        Case "rtcWriteFile": vbName = "Write": arity = -1: isStmt = True
        Case "rtcKillFiles", "rtcKill": vbName = "Kill": arity = 1: isStmt = True
    End Select
End Sub

'---------------------------------------------------------------------------
' Control resolution (offset -> control name), shared with the P-Code map
'---------------------------------------------------------------------------

Private Function NativeIsControlCall(inst As CInstruction, ByRef disp As Long) As Boolean
    'A call [reg+disp] (not absolute) whose displacement is in the control-accessor
    'range of the form vtable.  Control accessors sit BELOW the user-method block
    '(which starts at 0x6F8); a call at 0x6F8+ is a form method, not a control, and
    'must be excluded or it poisons the base solve (no control maps to its index).
    Dim isAbs As Boolean
    If (inst.cmdType And C_TYPEMASK) <> C_CAL Then Exit Function
    If Not NativeDecodeDisp(inst.dump, disp, isAbs) Then Exit Function
    If isAbs Then Exit Function
    NativeIsControlCall = (disp >= &H250 And disp < &H6F8)
End Function

Private Function NativeSolveControlBase(col As Collection) As Long
    'VB6 lays a form's control accessors in its vtable starting at a FIXED offset:
    'control with index i (controls are 1-based) is at 0x2F8 + i*4.  This base is a
    'standard of the VB6 form interface - verified identical across binaries (Step2
    'and pMasterMaker both 0x2F8) - and sits a constant 0x400 below the user-method
    'block at 0x6F8.  Solving the base by aligning offsets to control indices is
    'ambiguous (with contiguous indices many bases "fit", so it mis-names controls),
    'so use the constant whenever it maps at least one of this proc's control calls
    'to a real control; fall back to the old per-proc solve only for a non-standard
    'layout (e.g. a UserControl) where the constant maps nothing.
    Const FORM_CONTROL_BASE As Long = &H2F8
    If NativeBaseHits(col, FORM_CONTROL_BASE) > 0 Then
        NativeSolveControlBase = FORM_CONTROL_BASE
    Else
        NativeSolveControlBase = NativeSolveControlBasePerProc(col)
    End If
End Function

Private Function NativeBaseHits(col As Collection, ByVal base As Long) As Long
    'How many of the proc's control-accessor calls map, at this base, to a REAL
    'control of the current form.  A positive count means the base is the right one.
    Dim inst As CInstruction, off As Long, idx As Long, c As Long
    On Error Resume Next
    For Each inst In col
        If NativeIsControlCall(inst, off) Then
            idx = off - base
            If idx >= 0 And (idx Mod 4) = 0 Then
                If NativeControlIndexName(idx \ 4) <> "" Then c = c + 1
            End If
        End If
    Next
    NativeBaseHits = c
End Function

Private Function NativeSolveControlBasePerProc(col As Collection) As Long
    Dim inst As CInstruction, k As Long, idx As Long, cand As Long, firstOff As Long, d As Long
    On Error Resume Next
    NativeSolveControlBasePerProc = -1
    firstOff = -1
    For Each inst In col
        If NativeIsControlCall(inst, d) Then firstOff = d: Exit For
    Next
    If firstOff < 0 Then Exit Function
    For k = 0 To UBound(gControlNameArray)
        If gControlNameArray(k).strParentForm = NVForm Then
            idx = gControlNameArray(k).lControlIndex
            cand = firstOff - idx * 4
            If cand >= 0 Then
                If NativeBaseFits(col, cand) Then NativeSolveControlBasePerProc = cand: Exit Function
            End If
        End If
    Next
End Function

Private Function NativeBaseFits(col As Collection, ByVal base As Long) As Boolean
    Dim inst As CInstruction, d As Long, off As Long
    On Error Resume Next
    For Each inst In col
        If NativeIsControlCall(inst, off) Then
            d = off - base
            If d < 0 Then Exit Function
            If (d Mod 4) <> 0 Then Exit Function
            If NativeControlIndexName(d \ 4) = "" Then Exit Function
        End If
    Next
    NativeBaseFits = True
End Function

Private Function NativeFormVtableTarget(ByVal disp As Long) As Long
    'Resolve "call [vtable + disp]" on the current object's own methods.  VB6
    'lays a form's user methods in its vtable starting at offset 0x6F8 (one 4-byte
    'slot per method, in the object's method order), so slot = (disp-0x6F8)/4.
    'gFormVtable maps "ObjectName:slot" -> method address (filled from the
    'event-link table).  Only forms reach an offset this large, so class/usercontrol
    'method calls (much smaller vtables) yield a negative slot and are left alone.
    Const FORM_VTABLE_BASE As Long = &H6F8
    Dim slot As Long, v As Variant
    If disp < FORM_VTABLE_BASE Then Exit Function
    If ((disp - FORM_VTABLE_BASE) Mod 4) <> 0 Then Exit Function
    slot = (disp - FORM_VTABLE_BASE) \ 4
    On Error Resume Next
    v = gFormVtable(NVForm & ":" & slot)
    If Err.Number = 0 Then NativeFormVtableTarget = CLng(v)
End Function

Private Function NativeTryMethodSig(ByVal addr As Long, ByRef sig As String) As Boolean
    'Reconstructed parameter-name list for a public class method, keyed by address
    '(filled by modNative.LinkNativePublicParams).  Returns True (with sig, possibly
    'empty for a no-parameter method) when one is recorded.
    On Error Resume Next
    Dim v As Variant
    v = gMethodSig("A" & addr)
    If Err.Number = 0 Then sig = CStr(v): NativeTryMethodSig = True
End Function

Private Sub NativeBuildArgNameMap(ByVal addr As Long, ByVal psig As String)
    'Build the per-proc substitution list mapping the generic arg_<offset> token of
    'each parameter to its recovered name (applied at proc finalisation, so all the
    'internal arg_<offset> tracking / detection stays intact).  Params sit at base +
    'i*4 (base 0xC for a method with a hidden Me at ebp+8, else 0x8); a hidden return
    'buffer, when present, is a TRAILING slot and does not shift the leading params.
    NVArgN = 0
    If Len(psig) = 0 Then Exit Sub
    Dim parts() As String, i As Long, base As Long, nm As String
    parts = Split(psig, ", ")
    base = IIf(NativeProcHasMe(addr), &HC, &H8)
    ReDim NVArgTok(UBound(parts)): ReDim NVArgNm(UBound(parts))
    For i = 0 To UBound(parts)
        nm = Trim$(parts(i))
        If InStr(nm, " ") > 0 Then nm = Mid$(nm, InStrRev(nm, " ") + 1)   'keep the bare identifier
        If Len(nm) > 0 And NativeIsIdent(nm) Then
            NVArgTok(NVArgN) = "arg_" & Hex$(base + i * 4)
            NVArgNm(NVArgN) = nm
            NVArgN = NVArgN + 1
        End If
    Next
End Sub

Private Sub NativeDetectReturnSlot(col As Collection, ByVal addr As Long)
    'A class Function / Property Get returns through a hidden [out,retval] pointer -
    'the LAST stack argument (ebp + NVRetN + 4) - and the epilogue copies the return
    'value from a local into it:  mov REGa,[ebp+retOff] ; mov REGb,[ebp-X] ; mov
    '[REGa],REGb.  Rename that local (var_X) to the procedure name, so its assignments
    'read `FuncName = value` (VB's implicit return variable) instead of var_X = value.
    On Error GoTo done
    If NVRetN < 4 Then Exit Sub
    'Only Functions / Property Get have a return value.
    Dim kind As String, isFunc As Boolean, mi As Long
    If NativeTryMethodKind(addr, kind) Then
        isFunc = (InStr(kind, "Function") > 0 Or InStr(kind, "Get") > 0)
    Else
        mi = NativeProcMatchIdx(addr)
        If mi >= 0 Then isFunc = (InStr(SubNamelist(mi).kind, "Function") > 0 Or InStr(SubNamelist(mi).kind, "Get") > 0)
    End If
    If Not isFunc Then Exit Sub
    Dim funcName As String
    funcName = NativeProcName(addr)
    Dim pp As Long
    pp = InStr(funcName, "(")
    If pp > 0 Then funcName = Left$(funcName, pp - 1)
    funcName = Trim$(funcName)
    If Not NativeIsIdent(funcName) Then Exit Sub

    Dim retOff As Long
    retOff = NVRetN + 4                                  'the hidden retbuf is the last stack slot
    Dim n As Long, k As Long, inst As CInstruction
    n = col.Count
    Dim arr() As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next
    Dim i As Long, ra As Long, rb As Long, dret As Long, dloc As Long, bs As Long, sr As Long
    For i = 0 To n - 3
        'mov REGa, [ebp + retOff]
        If NativeMovRegEbp(arr(i), ra, dret) Then
            If dret = retOff Then
                'mov REGb, [ebp - X]  (the return local; X negative)
                If NativeMovRegEbp(arr(i + 1), rb, dloc) Then
                    If dloc < 0 Then
                        'mov [REGa], REGb  (store the local through the retbuf pointer)
                        If NativeMovToBase(arr(i + 2), bs, sr) Then
                            If bs = ra And sr = rb Then
                                ReDim Preserve NVArgTok(NVArgN): ReDim Preserve NVArgNm(NVArgN)
                                NVArgTok(NVArgN) = "var_" & Hex$(Abs(dloc))
                                NVArgNm(NVArgN) = funcName
                                NVArgN = NVArgN + 1
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
done:
End Sub

Private Function NativeMovRegEbp(inst As CInstruction, ByRef reg As Long, ByRef disp As Long) As Boolean
    'Match `mov r(16/32), [ebp + disp]` (8B /r, base ebp): dest reg + signed disp.
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long, isAbs As Boolean
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op <> &H8B Then Exit Function
    If NativeMemBase(dump) <> 5 Then Exit Function          'base must be ebp
    modrm = NativeDumpByte(dump, i + 1)
    If (modrm \ &H40) = 3 Then Exit Function                'register-direct, not memory
    reg = (modrm \ 8) And 7
    If NativeDecodeDisp(dump, disp, isAbs) Then If Not isAbs Then NativeMovRegEbp = True
End Function

Private Function NativeMovToBase(inst As CInstruction, ByRef baseReg As Long, ByRef srcReg As Long) As Boolean
    'Match `mov [REG], r(16/32)` (89 /r, mod=00, rm=base register, no disp).
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long, md As Long, rm As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op <> &H89 Then Exit Function
    modrm = NativeDumpByte(dump, i + 1)
    md = (modrm \ &H40) And 3: rm = modrm And 7
    If md <> 0 Or rm = 4 Or rm = 5 Then Exit Function       'need [REG] with no disp, no SIB/ebp
    baseReg = rm
    srcReg = (modrm \ 8) And 7
    NativeMovToBase = True
End Function

Private Function NativeIsIdent(ByVal s As String) As Boolean
    'A safe VB identifier (letters / digits / underscore, not starting with a digit).
    Dim i As Long, c As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = Asc(Mid$(s, i, 1))
        If Not ((c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c = 95 Or (c >= 48 And c <= 57 And i > 1)) Then Exit Function
    Next
    NativeIsIdent = True
End Function

Private Function NativeSubstituteArgNames(ByVal src As String) As String
    'Replace each generic arg_<offset> token in the finished proc body with its
    'recovered parameter name (whole-identifier match, so arg_1 never clobbers part
    'of arg_10 and a substring inside another identifier / literal is left alone).
    NativeSubstituteArgNames = src
    If NVArgN <= 0 Then Exit Function
    Dim i As Long
    For i = 0 To NVArgN - 1
        NativeSubstituteArgNames = NativeReplaceToken(NativeSubstituteArgNames, NVArgTok(i), NVArgNm(i))
    Next
End Function

Private Function NativeReplaceToken(ByVal src As String, ByVal tok As String, ByVal repl As String) As String
    'Replace whole-identifier occurrences of tok with repl (a token boundary is any
    'char that is not a letter, digit or underscore).
    Dim res As String, p As Long, q As Long, before As Long, after As Long, tl As Long
    tl = Len(tok)
    p = 1
    Do
        q = InStr(p, src, tok)
        If q = 0 Then res = res & Mid$(src, p): Exit Do
        before = 0: after = 0
        If q > 1 Then before = Asc(Mid$(src, q - 1, 1))
        If q + tl <= Len(src) Then after = Asc(Mid$(src, q + tl, 1))
        If NativeIsIdentChar(before) Or NativeIsIdentChar(after) Then
            res = res & Mid$(src, p, q - p + tl)            'part of a larger token - keep as-is
        Else
            res = res & Mid$(src, p, q - p) & repl
        End If
        p = q + tl
    Loop
    NativeReplaceToken = res
End Function

Private Function NativeIsIdentChar(ByVal c As Long) As Boolean
    NativeIsIdentChar = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c = 95 Or (c >= 48 And c <= 57)
End Function

Private Function NativeTryMethodKind(ByVal addr As Long, ByRef kind As String) As Boolean
    'Method kind (Sub/Function/Property Get) recovered from the class typeinfo, keyed
    'by address (filled by modNative.LinkNativePublicParams).
    On Error Resume Next
    Dim v As Variant
    v = gMethodKind("A" & addr)
    If Err.Number = 0 Then kind = CStr(v): NativeTryMethodKind = True
End Function

Private Function NativeArgCount(ByVal csv As String) As Long
    If Len(csv) = 0 Then Exit Function
    NativeArgCount = UBound(Split(csv, ", ")) + 1
End Function

Private Function NativeTakeArgs(ByVal csv As String, ByVal n As Long) As String
    'The first n comma-separated arguments of csv (used to drop a method call's
    'trailing hidden return-value buffer once the real parameter count is known).
    If n <= 0 Or Len(csv) = 0 Then Exit Function
    Dim parts() As String, i As Long, s As String
    parts = Split(csv, ", ")
    If n > UBound(parts) + 1 Then n = UBound(parts) + 1
    For i = 0 To n - 1
        If Len(s) > 0 Then s = s & ", "
        s = s & parts(i)
    Next
    NativeTakeArgs = s
End Function

Private Function NativeDropThisArgs(ByVal rawArgs As String, ByVal tgt As Long, ByRef retbuf As String) As String
    'Strip the implicit Me/this (the leading arg_8) and the trailing [out,retval]
    'buffer from a COM/form method call's raw argument list, keeping exactly the
    'method's real parameters (count from its typeinfo signature).  retbuf returns the
    'dropped buffer arg ("" if none), so the caller can surface a value-returning call.
    retbuf = ""
    If Len(rawArgs) = 0 Then Exit Function
    Dim p() As String, st As Long, keep As Long, total As Long, i As Long, out As String, sig As String
    p = Split(rawArgs, ", ")
    st = 0
    If p(0) = "arg_8" Then st = 1                        'drop the implicit Me/this
    total = UBound(p) - st + 1
    keep = total
    If NativeTryMethodSig(tgt, sig) Then keep = NativeArgCount(sig)
    If keep > total Then keep = total
    If keep < 0 Then keep = 0
    For i = st To st + keep - 1
        If i > UBound(p) Then Exit For
        If Len(out) > 0 Then out = out & ", "
        out = out & p(i)
    Next
    If total > keep Then retbuf = p(st + keep)
    NativeDropThisArgs = out
End Function

Private Function NativeClassVtableTarget(ByVal disp As Long) As Long
    'Resolve "call [Me_vtable + disp]" on a CLASS's own methods.  A COM class vtable
    'is IUnknown(3) + IDispatch(4) = 7 slots, so user methods begin at 0x1C; the map
    '(filled in LinkNativeProcNames for class/usercontrol objects) is keyed by the
    'absolute offset: "Owner:off<disp>" -> method address.
    Dim v As Variant
    On Error Resume Next
    v = gFormVtable(NVForm & ":off" & disp)
    If Err.Number = 0 Then NativeClassVtableTarget = CLng(v)
End Function

Private Function NativeControlByOffset(ByVal offset As Long) As String
    If NVBase < 0 Or offset < NVBase Or ((offset - NVBase) Mod 4) <> 0 Then Exit Function
    NativeControlByOffset = NativeControlIndexName((offset - NVBase) \ 4)
End Function

Private Function NativeGuidByOffset(ByVal offset As Long) As String
    Dim k As Long, idx As Long
    On Error Resume Next
    If NVBase < 0 Then Exit Function
    idx = (offset - NVBase) \ 4
    For k = 0 To UBound(gControlNameArray)
        If gControlNameArray(k).strParentForm = NVForm And gControlNameArray(k).lControlIndex = idx Then
            NativeGuidByOffset = gControlNameArray(k).strGuid: Exit Function
        End If
    Next
End Function

Private Function NativeControlIndexName(ByVal idx As Long) As String
    Dim k As Long
    On Error Resume Next
    For k = 0 To UBound(gControlNameArray)
        If gControlNameArray(k).strParentForm = NVForm And gControlNameArray(k).lControlIndex = idx Then
            NativeControlIndexName = gControlNameArray(k).strControlName: Exit Function
        End If
    Next
End Function

'---------------------------------------------------------------------------
' Small utilities
'---------------------------------------------------------------------------

Private Function NativeMnem(inst As CInstruction) As String
    Dim s As String, p As Long
    s = Trim$(inst.command)
    p = InStr(s, " ")
    If p > 0 Then s = Left$(s, p - 1)
    NativeMnem = UCase$(s)
End Function

Private Function NativeApiName(inst As CInstruction) As String
    'Resolve from the absolute address decoded out of the instruction bytes,
    'so we never re-read the wrong operand or hit a bad seek.
    Dim disp As Long, isAbs As Boolean
    On Error Resume Next
    If NativeDecodeDisp(inst.dump, disp, isAbs) Then
        If isAbs And disp >= OptHeader.ImageBase Then NativeApiName = dsmNative.GetApiByIatVa(disp)
    End If
End Function

'---------------------------------------------------------------------------
' Instruction-byte displacement decoder
' olly's adrConst is unreliable for [reg+disp32] operands, so the control,
' property and FPU offsets are decoded straight from the machine bytes.
'---------------------------------------------------------------------------

Private Function NativeDecodeDisp(ByVal dump As String, ByRef disp As Long, ByRef isAbs As Boolean) As Boolean
    'Parse the memory displacement of a single-byte-opcode ModR/M instruction
    '(call FF, FPU D8-DF, mov 8B/89, lea 8D...).  Returns True when the
    'instruction has a memory operand; sets isAbs for an absolute [disp32].
    Dim n As Long, i As Long, op As Long, modrm As Long, md As Long, rm As Long, base As Long
    On Error GoTo no
    dump = Replace(dump, " ", "")
    disp = 0: isAbs = False
    n = Len(dump) \ 2
    i = 0
    Do While i < n              'skip legacy prefixes
        op = NativeDumpByte(dump, i)
        Select Case op
            Case &H66, &H67, &HF0, &HF2, &HF3, &H26, &H2E, &H36, &H3E, &H64, &H65
                i = i + 1
            Case Else
                Exit Do
        End Select
    Loop
    If i >= n Then GoTo no
    op = NativeDumpByte(dump, i): i = i + 1
    If op = &HF Then GoTo no     '2-byte opcode (0F xx) not handled
    'E8/E9 (call/jmp rel32) and EB (jmp rel8) carry NO ModR/M byte - the bytes
    'that follow are a relative displacement, not a memory operand.  Treating
    'the first rel byte as ModR/M wrongly reports a [reg+disp] memory call
    '(e.g. "call .0"), so reject these opcodes outright.
    If op = &HE8 Or op = &HE9 Or op = &HEB Then GoTo no
    If i >= n Then GoTo no
    modrm = NativeDumpByte(dump, i): i = i + 1
    md = (modrm \ &H40) And 3
    rm = modrm And 7
    If md = 3 Then GoTo no       'register operand, no memory
    If rm = 4 Then               'SIB byte present
        If i >= n Then GoTo no
        base = NativeDumpByte(dump, i) And 7: i = i + 1
    Else
        base = rm
    End If
    If md = 0 Then
        If rm = 5 Then
            disp = NativeDumpInt32(dump, i): isAbs = True
        ElseIf rm = 4 And base = 5 Then
            disp = NativeDumpInt32(dump, i)
        Else
            disp = 0
        End If
    ElseIf md = 1 Then
        disp = NativeDumpInt8(dump, i)
    Else                         'md = 2
        disp = NativeDumpInt32(dump, i)
    End If
    NativeDecodeDisp = True
    Exit Function
no:
    NativeDecodeDisp = False
End Function

Private Function NativeDumpByte(ByVal dump As String, ByVal idx As Long) As Long
    NativeDumpByte = CLng("&H" & Mid$(dump, idx * 2 + 1, 2))
End Function

Private Function NativeDumpInt8(ByVal dump As String, ByVal idx As Long) As Long
    Dim v As Long
    v = NativeDumpByte(dump, idx)
    If v >= 128 Then v = v - 256
    NativeDumpInt8 = v
End Function

Private Function NativeDumpInt32(ByVal dump As String, ByVal idx As Long) As Long
    Dim bb(3) As Byte, k As Long, v As Long
    For k = 0 To 3
        bb(k) = CByte(NativeDumpByte(dump, idx + k))
    Next
    CopyMemory v, bb(0), 4
    NativeDumpInt32 = v
End Function

Private Function NativeDumpInt16(ByVal dump As String, ByVal idx As Long) As Long
    'Signed little-endian 16-bit value (so a Boolean True word 0xFFFF reads as -1).
    Dim v As Long
    v = NativeDumpByte(dump, idx) + NativeDumpByte(dump, idx + 1) * &H100&
    If v >= &H8000& Then v = v - &H10000
    NativeDumpInt16 = v
End Function

Private Function NativeHas66(ByVal dump As String) As Boolean
    'True when the instruction carries the 0x66 operand-size prefix (a 16-bit
    'store), so an immediate is read as a word, not a dword.
    Dim i As Long, n As Long
    n = Len(dump) \ 2
    For i = 0 To n - 1
        Select Case NativeDumpByte(dump, i)
            Case &H66: NativeHas66 = True: Exit Function
            Case &H67, &HF0, &HF2, &HF3, &H26, &H2E, &H36, &H3E, &H64, &H65   'skip other prefixes
            Case Else: Exit Function
        End Select
    Next
End Function

Private Function NativeFieldName(ByVal disp As Long) As String
    'Name of an instance FIELD at byte offset `disp` of the current form/class
    'instance (the Me pointer).  Prefers the public-variable name recovered from
    'the object's typeinfo (gFieldName, keyed Owner:offset); falls back to a
    'synthetic field_<off> until those names are linked.
    On Error Resume Next
    Dim s As String
    s = gFieldName(NVForm & ":" & disp)
    If Len(s) > 0 Then NativeFieldName = s Else NativeFieldName = "field_" & Hex$(disp)
End Function

Private Function NativeFieldStoreLHS(ByVal base As Long, ByVal off As Long) As String
    'Left-hand side for a store `mov [base + off], value` (off > 0):
    '  - an instance field of Me (this/self) in a class/form method -> the public
    '    variable name (or synthetic field_<off>);
    '  - else a field of a UDT the base register holds BY REFERENCE (a ByRef param
    '    or tracked pointer) -> base(off), the same deref-with-offset form a read
    '    renders as.
    'Empty when the base is neither, so an unmodelled store is left dropped.  The
    'Me case is gated to NVHasMe so a .bas module's ByRef param (also at ebp+8) is
    'never mistaken for an instance field.
    If base < 0 Or base > 7 Then Exit Function
    If NVHasMe And NVRegIsMe(base) Then
        NativeFieldStoreLHS = NativeFieldName(off)
    ElseIf NativeIsDerefBase(NVReg(base)) Then
        NativeFieldStoreLHS = NVReg(base) & "(" & CStr(off) & ")"
    End If
End Function

Private Function NativeIsResumePush(inst As CInstruction, ByVal procStart As Long) As Boolean
    'A "push imm32" of an address inside this procedure - the resume address VB
    'pushes immediately before jumping over the On Error handler block.
    Dim dump As String, n As Long, i As Long, imm As Long
    On Error Resume Next
    If (inst.cmdType And C_TYPEMASK) <> C_PSH Then Exit Function
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    If n < 5 Then Exit Function
    i = NativeOpStart(dump, n)
    If NativeDumpByte(dump, i) <> &H68 Then Exit Function       'push imm32
    imm = NativeDumpInt32(dump, i + 1)
    If imm >= OptHeader.ImageBase And imm > inst.va And imm < inst.va + &H10000 Then NativeIsResumePush = True
End Function

Private Function NativeOpStart(ByVal dump As String, ByVal n As Long) As Long
    'Index (in bytes) of the primary opcode, past any legacy prefixes.
    Dim i As Long, op As Long
    Do While i < n
        op = NativeDumpByte(dump, i)
        Select Case op
            Case &H66, &H67, &HF0, &HF2, &HF3, &H26, &H2E, &H36, &H3E, &H64, &H65: i = i + 1
            Case Else: Exit Do
        End Select
    Loop
    NativeOpStart = i
End Function

Private Function NativePushOperand(inst As CInstruction) As String
    'The symbolic value being pushed: immediate / string-constant / local slot /
    'tracked register.  Empty for an untracked push (e.g. the "this" pointer).
    Dim dump As String, n As Long, i As Long, op As Long, disp As Long, isAbs As Boolean, imm As Long, s As String
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    If n < 1 Then Exit Function
    i = NativeOpStart(dump, n)
    op = NativeDumpByte(dump, i)
    Select Case op
        Case &H68                       'push imm32 (string constant / global address / number)
            imm = NativeDumpInt32(dump, i + 1)
            If imm >= OptHeader.ImageBase Then s = NativeStringAt(imm)
            If Len(s) > 0 Then
                NativePushOperand = s
            ElseIf NativeIsGlobalAddr(imm) Then
                NativePushOperand = NativeGlobalName(imm)
            Else
                NativePushOperand = NativeNumFromBits(imm)
            End If
        Case &H6A                       'push imm8
            NativePushOperand = CStr(NativeDumpInt8(dump, i + 1))
        Case &H50 To &H57               'push reg
            'A register holding &local (from LEA) is a by-reference argument:
            'show the local itself rather than its (often 0/stale) value.
            If NVRegIsAddr(op - &H50) Then
                NativePushOperand = NVRegAddr(op - &H50)
                NVLastPushDisp = NVRegAddrDisp(op - &H50)   'by-ref local -> Variant resolution
            Else
                NativePushOperand = NVReg(op - &H50)
            End If
        Case &HFF                       'push r/m  (FF /6)
            If NativeDecodeDisp(dump, disp, isAbs) Then
                If Not isAbs And disp < 0 Then NativePushOperand = NativeGetLocalExpr(disp)
            End If
    End Select
End Function

Private Function NativeTrackReg(inst As CInstruction) As String
    'Lightweight GP-register value tracking for mov/lea/xor so pushes and
    'assignment right-hand-sides can be reconstructed.  Returns a "var_X = expr"
    'assignment when a recovered call result is stored to a local, else "".
    Dim dump As String, n As Long, i As Long, op As Long, modrm As Long, md As Long, reg As Long, rm As Long
    Dim disp As Long, isAbs As Boolean, lname As String
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    If n < 2 Then Exit Function
    i = NativeOpStart(dump, n)
    op = NativeDumpByte(dump, i)
    Select Case op
        Case &HB8 To &HBF               'mov reg, imm32
            Dim immv As Long, sv As String
            immv = NativeDumpInt32(dump, i + 1)
            If immv >= OptHeader.ImageBase Then sv = NativeStringAt(immv)
            If Len(sv) > 0 Then NVReg(op - &HB8) = sv Else NVReg(op - &HB8) = NativeNumFromBits(immv)
            NVRegIsAddr(op - &HB8) = False: NVRegIsMe(op - &HB8) = False: NVRegIsFormVt(op - &HB8) = False
            NVRegObjType(op - &HB8) = "": NVRegObjVt(op - &HB8) = "": NVRegObjGuid(op - &HB8) = "": NVRegObjVtGuid(op - &HB8) = ""
        Case &H8B                       'mov r32, r/m32
            modrm = NativeDumpByte(dump, i + 1)
            md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
            If NativeHas66(dump) Then
                'A 16-bit move (mov si,ax) writes only the LOW WORD of the dest; we
                'model whole 32-bit values, so the dest is now unknown.  Copying the
                'source clobbered the high half with a stale value and collapsed
                'conditions like `SendMessage(...) <> 0` into `0 <> 0`.
                NVReg(reg) = "": NVRegIsAddr(reg) = False: NVRegIsMe(reg) = False: NVRegIsFormVt(reg) = False
                NVRegObjType(reg) = "": NVRegObjVt(reg) = "": NVRegObjGuid(reg) = "": NVRegObjVtGuid(reg) = ""
            ElseIf md = 3 Then
                NVReg(reg) = NVReg(rm)
                NVRegIsAddr(reg) = NVRegIsAddr(rm): NVRegAddr(reg) = NVRegAddr(rm): NVRegAddrDisp(reg) = NVRegAddrDisp(rm)   'address propagates on reg->reg
                NVRegIsMe(reg) = NVRegIsMe(rm): NVRegIsFormVt(reg) = NVRegIsFormVt(rm)
                NVRegObjType(reg) = NVRegObjType(rm): NVRegObjVt(reg) = NVRegObjVt(rm)
                NVRegObjGuid(reg) = NVRegObjGuid(rm): NVRegObjVtGuid(reg) = NVRegObjVtGuid(rm)
                NVRegObjInst(reg) = NVRegObjInst(rm)
            ElseIf NativeDecodeDisp(dump, disp, isAbs) Then
                Dim bse As Long, baseObj As Boolean
                bse = NativeMemBase(dump)
                If bse >= 0 And bse <= 7 Then baseObj = NVRegIsMe(bse)
                If Not isAbs And disp < 0 Then
                    NVReg(reg) = NativeGetLocalExpr(disp)
                ElseIf isAbs And NativeIsGlobalAddr(disp) Then
                    NVReg(reg) = NativeGlobalName(disp)      'load of a module-level global
                ElseIf Not isAbs And bse = 5 And disp >= 8 And disp <= &H200 Then
                    NVReg(reg) = "arg_" & Hex$(disp)         'a procedure parameter (ebp+positive)
                ElseIf Not isAbs And disp = 0 And bse >= 0 And bse <= 7 And Left$(NVReg(bse), 4) = "arg_" Then
                    NVReg(reg) = NVReg(bse)                  'deref of a ByRef parameter pointer -> its value
                ElseIf Not isAbs And NativeMemIndex(dump) < 0 And bse >= 0 And bse <= 7 And NativeIsDerefBase(NVReg(bse)) Then
                    'Deref-with-offset of a tracked DATA pointer (NativeIsDerefBase
                    'excludes bare-constant / control / property / folded bases):
                    'reg = <baseVal>(disp).  Builds the chained struct / SAFEARRAY
                    'field expressions the operand renderer emits as global_X(12)(20).
                    If disp = 0 Then NVReg(reg) = NVReg(bse) Else NVReg(reg) = NVReg(bse) & "(" & CStr(disp) & ")"
                Else
                    NVReg(reg) = ""
                End If
                NVRegIsAddr(reg) = False
                'Track object pointers and their vtables so an intrinsic-global
                'getter `call [objVt + 0x14]` can be resolved.  An object pointer
                'comes from `[ebp+8]` (this/Me) or from a module global `[abs]`;
                'its vtable is the deref `[objPtr]`.
                NVRegIsMe(reg) = (Not isAbs And disp = 8 And bse = 5) Or (isAbs And disp >= OptHeader.ImageBase)
                NVRegIsFormVt(reg) = (Not isAbs And disp = 0 And baseObj)
                'Propagate intrinsic-object identity for property chains (App.Path):
                'loading an App/Screen/Clipboard-typed local tags the register as
                'that object pointer; dereferencing such a pointer ([objPtr], disp 0)
                'tags the register as that object's vtable.
                NVRegObjType(reg) = "": NVRegObjVt(reg) = "": NVRegObjGuid(reg) = "": NVRegObjVtGuid(reg) = "": NVRegObjInst(reg) = ""
                If Not isAbs And disp < 0 Then
                    If NativeIsIntrinsicObj(NVReg(reg)) Then NVRegObjType(reg) = NVReg(reg)
                    'Loading a local that holds a control object tags this register
                    'with the control's identity + GUID (for a later property call).
                    Dim lguid As String
                    lguid = NativeGetLocalGuid(disp)
                    If Len(lguid) > 0 Then NVRegObjGuid(reg) = lguid: NVRegObjType(reg) = NVReg(reg)
                ElseIf isAbs And NativeIsGlobalAddr(disp) Then
                    'Loading a module-global that holds a user-class instance (typed
                    'at its __vbaNew auto-instantiation) tags the register as that
                    'object pointer, so a following `mov vt,[obj]; call [vt+off]`
                    'resolves to obj.Method.
                    Dim gcls As String
                    gcls = NativeColGet(NVObjClass, "G" & disp)
                    If Len(gcls) > 0 Then NVRegObjType(reg) = gcls: NVRegObjInst(reg) = NVReg(reg)
                ElseIf Not isAbs And disp = 0 And bse >= 0 And bse <= 7 Then
                    NVRegObjVt(reg) = NVRegObjType(bse)
                    NVRegObjVtGuid(reg) = NVRegObjGuid(bse)   'deref of a control pointer -> its vtable carries the GUID
                    NVRegObjInst(reg) = NVRegObjInst(bse)     'deref of a user-class pointer -> its vtable carries the receiver
                ElseIf Not isAbs And disp > 0 And baseObj Then
                    'Reading an `As New` private object field Me.<field> (e.g. a
                    'clsBitmap member): type it from the auto-instantiation map so a
                    'following `mov vt,[field]; call [vt+off]` resolves to <Class>.Method.
                    Dim ffcls As String
                    ffcls = NativeColGet(gFormFieldClass, NVForm & ":" & disp)
                    If Len(ffcls) > 0 Then NVRegObjType(reg) = ffcls: NVRegObjInst(reg) = ffcls
                End If
            End If
        Case &H8D                       'lea r32, [mem]  (address-of / ptr arithmetic)
            modrm = NativeDumpByte(dump, i + 1)
            reg = (modrm \ 8) And 7
            If NativeDecodeDisp(dump, disp, isAbs) Then
                Dim lbse As Long
                lbse = NativeMemBase(dump)
                'LEA on [ebp-X] takes the ADDRESS of a local.  Keep its value in
                'NVReg (a read of the register wants the value), but ALSO remember
                'the local's name so a later PUSH of the register - passing the
                'local by reference - shows the local, not its (often 0/stale)
                'value.  Gated to base = ebp(5): otherwise a `lea esi,[edi-1]`
                '(index arithmetic i-1) was mislabelled as the local var_1.
                If Not isAbs And disp < 0 And lbse = 5 And NativeMemIndex(dump) < 0 Then
                    NVReg(reg) = NativeGetLocalExpr(disp)
                    NVRegIsAddr(reg) = True: NVRegAddr(reg) = "var_" & Hex$(Abs(disp)): NVRegAddrDisp(reg) = disp
                ElseIf isAbs And NativeIsGlobalAddr(disp) Then
                    'address-of a module-level global (e.g. an array passed by ref)
                    NVReg(reg) = NativeGlobalName(disp): NVRegIsAddr(reg) = False
                ElseIf Not isAbs And lbse >= 0 And lbse <= 7 And lbse <> 5 And NativeMemIndex(dump) < 0 _
                       And Len(NVReg(lbse)) > 0 And Not NativeIsNumLit(NVReg(lbse)) Then
                    'lea reg,[base+disp] on a tracked non-constant base = address /
                    'index arithmetic, e.g. lea esi,[edi-1] -> the zero-based index
                    'i-1 (used to index a 0-based SAFEARRAY from a 1-based counter).
                    If disp = 0 Then
                        NVReg(reg) = NVReg(lbse)
                    ElseIf disp < 0 Then
                        NVReg(reg) = "(" & NVReg(lbse) & " - " & CStr(-disp) & ")"
                    Else
                        NVReg(reg) = "(" & NVReg(lbse) & " + " & CStr(disp) & ")"
                    End If
                    NVRegIsAddr(reg) = False
                Else
                    NVReg(reg) = "": NVRegIsAddr(reg) = False
                End If
                NVRegIsMe(reg) = False: NVRegIsFormVt(reg) = False
                NVRegObjType(reg) = "": NVRegObjVt(reg) = "": NVRegObjGuid(reg) = "": NVRegObjVtGuid(reg) = ""
            End If
        Case &H89                       'mov r/m32, r32 (store)
            modrm = NativeDumpByte(dump, i + 1)
            md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
            If NativeHas66(dump) And md = 3 Then
                '16-bit reg->reg write (mov bx,ax): partial, so the dest register's
                'tracked 32-bit value is now unknown.  (A 16-bit write to MEMORY -
                'the word field stores - falls through to the store path below.)
                NVReg(rm) = "": NVRegIsAddr(rm) = False: NVRegIsMe(rm) = False: NVRegIsFormVt(rm) = False
                NVRegObjType(rm) = "": NVRegObjVt(rm) = "": NVRegObjGuid(rm) = "": NVRegObjVtGuid(rm) = ""
            ElseIf md = 3 Then
                NVReg(rm) = NVReg(reg)
                NVRegIsAddr(rm) = NVRegIsAddr(reg): NVRegAddr(rm) = NVRegAddr(reg): NVRegAddrDisp(rm) = NVRegAddrDisp(reg)
                NVRegIsMe(rm) = NVRegIsMe(reg): NVRegIsFormVt(rm) = NVRegIsFormVt(reg)
                NVRegObjType(rm) = NVRegObjType(reg): NVRegObjVt(rm) = NVRegObjVt(reg)
                NVRegObjGuid(rm) = NVRegObjGuid(reg): NVRegObjVtGuid(rm) = NVRegObjVtGuid(reg)
            ElseIf NativeDecodeDisp(dump, disp, isAbs) Then
                If Not isAbs And disp < 0 Then
                    'Record the register's value against the slot (a Variant data or
                    'VT field filled from a register, e.g. `mov [ebp-X], esi` with
                    'esi = 0xA for a missing optional argument).
                    NativeSetVSlot disp, NVReg(reg)
                    'A stored call result (expression containing a call) is worth
                    'surfacing as a real assignment; bind the local to its name so
                    'later uses reference the variable rather than re-expanding.  A
                    'loop induction slot (NVCounterSlot) is treated the same way even
                    'for a plain value, so the counter shows as var_X (not a stale
                    'per-iteration constant) and its init / increment surface.
                    Dim isCtrSlot As Boolean
                    isCtrSlot = Len(NativeColGet(NVCounterSlot, "C" & disp)) > 0
                    lname = "var_" & Hex$(Abs(disp))
                    If InStr(NVReg(reg), "(") > 0 And Left$(NVReg(reg), 1) <> Chr$(34) Then
                        NativeTrackReg = lname & " = " & NVReg(reg)
                        NativeSetLocalExpr disp, lname
                    ElseIf isCtrSlot And Len(NVReg(reg)) > 0 And NVReg(reg) <> lname Then
                        NativeTrackReg = lname & " = " & NVReg(reg)
                        NativeSetLocalExpr disp, lname
                    Else
                        NativeSetLocalExpr disp, NVReg(reg)
                    End If
                    'Storing a tracked control object to a local (plain mov, not
                    '__vbaObjSet) - remember its GUID so a later property access
                    'through that local resolves (e.g. the LET target temp).
                    If Len(NVRegObjGuid(reg)) > 0 Then NativeSetLocalGuid disp, NVRegObjGuid(reg)
                    'If the register held NO symbolic value, it now MIRRORS this local
                    '(after `mov [var_X], reg` the register equals var_X).  Name it so
                    'a following read/compare shows var_X instead of a raw register -
                    'e.g. a register-allocated loop counter `cmp eax,edi` right after
                    '`mov [var_20],eax`.  Gated to empty NVReg so a meaningful tracked
                    'expression is never replaced by the bare local name - EXCEPT for a
                    'loop induction slot, whose register must read as var_X (the value
                    'it holds is this iteration's, stale for the rest of the loop).
                    If Len(NVReg(reg)) = 0 Or isCtrSlot Then NVReg(reg) = lname
                ElseIf isAbs And NativeIsGlobalAddr(disp) Then
                    'Store to a module-level global: mov [abs], reg.  Surface any
                    'tracked value (number / local / global / string / call / concat)
                    'as `global_X = ...`; a bare/untracked register is left dropped.
                    If Len(NVReg(reg)) > 0 Then
                        NativeTrackReg = NativeGlobalName(disp) & " = " & NVReg(reg)
                    End If
                ElseIf Not isAbs And disp > 0 And disp < &H2000 Then
                    'Store to a struct FIELD: mov [base + off], reg - an instance
                    'field of Me, or a ByRef-passed UDT field (see NativeFieldStoreLHS).
                    Dim lhs89 As String, fv89 As String
                    lhs89 = NativeFieldStoreLHS(NativeMemBase(dump), disp)
                    If Len(lhs89) > 0 Then
                        fv89 = NVReg(reg)
                        If Len(fv89) = 0 Then fv89 = NativeRegName(reg)
                        NativeTrackReg = lhs89 & " = " & fv89
                    End If
                End If
            End If
        Case &HC7                       'mov r/m32, imm32 (store immediate)
            If NativeDecodeDisp(dump, disp, isAbs) Then
                If isAbs And NativeIsGlobalAddr(disp) Then
                    'mov [abs], imm32 to a module global -> global_X = <imm/string>.
                    Dim c7imm As Long, c7s As String
                    c7imm = NativeDumpInt32(dump, n - 4)         'imm32 is the trailing dword
                    If c7imm >= OptHeader.ImageBase Then c7s = NativeStringAt(c7imm)
                    If Len(c7s) = 0 Then c7s = NativeNumFromBits(c7imm)
                    NativeTrackReg = NativeGlobalName(disp) & " = " & c7s
                ElseIf Not isAbs And disp > 0 And disp < &H2000 Then
                    'Store an immediate to a struct FIELD: mov [base + off], imm.
                    'A 0x66-prefixed store writes a word (Boolean True = 0xFFFF -> -1).
                    Dim lhsC7 As String, fimm As Long, fis As String
                    lhsC7 = NativeFieldStoreLHS(NativeMemBase(dump), disp)
                    If Len(lhsC7) > 0 Then
                        If NativeHas66(dump) Then fimm = NativeDumpInt16(dump, n - 2) Else fimm = NativeDumpInt32(dump, n - 4)
                        If fimm >= OptHeader.ImageBase Then fis = NativeStringAt(fimm)
                        If Len(fis) = 0 Then fis = NativeNumFromBits(fimm)
                        NativeTrackReg = lhsC7 & " = " & fis
                    End If
                End If
            End If
        Case &H33                       'xor r32, r/m32 (xor reg,reg -> 0)
            modrm = NativeDumpByte(dump, i + 1)
            md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
            If md = 3 And reg = rm Then NVReg(reg) = "0": NVRegIsAddr(reg) = False: NVRegIsMe(reg) = False: NVRegIsFormVt(reg) = False: NVRegObjType(reg) = "": NVRegObjVt(reg) = "": NVRegObjGuid(reg) = "": NVRegObjVtGuid(reg) = ""
        '--- Arithmetic folding (Option 1c) ---------------------------------------
        'Fold add/sub of a tracked operand so a compare LHS reads as the real
        'expression, e.g. (arg_C - global_X(20)).  When either side is unknown the
        'register is cleared, never left stale.  TEST/CMP are decoded and Exit'd
        'earlier, so they never reach here.  (Kept deliberately narrow: broad
        'clears on every arithmetic op regressed resolved control-property /
        'call-result registers, so only add/sub - the struct-offset idiom - folds.)
        Case &H03, &H2B                 'add / sub  r32, r/m32  -> fold into reg dest
            modrm = NativeDumpByte(dump, i + 1)
            md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
            Dim rhsv As String, oprt As String, lhsv As String
            lhsv = NVReg(reg)
            rhsv = NativeRmVal(dump, md, rm)
            If op = &H2B Then oprt = " - " Else oprt = " + "
            'Fold ONLY a meaningful, bounded expression: both operands must be a
            'data-pointer expression or a numeric constant (never a bare register
            'or control/property name), at least one must be a real pointer
            'reference, and the result is length-capped so repeated `add reg,reg`
            'in float/constant code cannot blow up into nested noise.  Otherwise
            'the register's value is dropped (the operand falls back to its raw name).
            If (NativeIsPtrExpr(lhsv) Or NativeIsNumLit(lhsv)) _
               And (NativeIsPtrExpr(rhsv) Or NativeIsNumLit(rhsv)) _
               And (NativeIsPtrExpr(lhsv) Or NativeIsPtrExpr(rhsv)) _
               And Len(lhsv) + Len(rhsv) <= 60 Then
                NVReg(reg) = "(" & lhsv & oprt & rhsv & ")"
                NVRegIsAddr(reg) = False: NVRegIsMe(reg) = False: NVRegIsFormVt(reg) = False
                NVRegObjType(reg) = "": NVRegObjVt(reg) = "": NVRegObjGuid(reg) = "": NVRegObjVtGuid(reg) = ""
            ElseIf NativeIsDerefBase(lhsv) Then
                'A data pointer was modified but the result is unfoldable: clear it
                'so a later deref can never render a STALE field (e.g. arg_C(20)).
                NVReg(reg) = ""
                NVRegIsAddr(reg) = False: NVRegIsMe(reg) = False: NVRegIsFormVt(reg) = False
                NVRegObjType(reg) = "": NVRegObjVt(reg) = "": NVRegObjGuid(reg) = "": NVRegObjVtGuid(reg) = ""
            End If
            'Otherwise leave the register untouched - a control / property / constant
            'value matches prior behaviour and Change A will never deref it anyway.
    End Select
End Function

Private Sub NativePushImm(ByVal s As String)
    If NVPushTop > UBound(NVPushImm) Then ReDim Preserve NVPushImm(NVPushTop + 16)
    NVPushImm(NVPushTop) = s
    NVPushTop = NVPushTop + 1
End Sub

Private Sub NativeRecordPushDisp(ByVal disp As Long)
    'Tag the most-recently pushed argument (index NVPushTop-1) with the by-reference
    'local displacement it addresses, so MsgBox/InputBox can resolve the Variant the
    'pointer refers to.  Kept parallel to NVPushImm.
    Dim k As Long
    k = NVPushTop - 1
    If k < 0 Then Exit Sub
    If k > UBound(NVPushDisp) Then ReDim Preserve NVPushDisp(k + 16)
    NVPushDisp(k) = disp
End Sub

Private Sub NativeArgsSnapshot(ByRef a() As String, ByRef cnt As Long)
    'Snapshot the pending pushed arguments in SOURCE order (last-pushed = arg 0)
    'into a() and drain the stack.  Lets a runtime-helper handler index its args
    'and reorder them into VB statement syntax (Open/Put/Get/ReDim/...).
    Dim k As Long, idx As Long
    cnt = NVPushTop
    If cnt < 0 Then cnt = 0
    ReDim a(IIf(cnt > 0, cnt - 1, 0))
    idx = 0
    For k = NVPushTop - 1 To 0 Step -1
        a(idx) = NVPushImm(k): idx = idx + 1
    Next
    NVPushTop = 0
End Sub

Private Function NativeRedimStmt(ByRef a() As String, ByVal cnt As Long, ByVal bPreserve As Boolean) As String
    'Build `ReDim [Preserve] arr(lb To ub, ...)` from a __vbaRedim arg snapshot:
    'a() = [elemFlags, elemSize, arrayPtr, saFlags, nDims, (ub,lb) per dimension].
    Dim arr As String, nd As Long, k As Long, bounds As String, ub As String, lb As String
    If cnt < 7 Then
        'Not enough args to parse the bound list - fall back to a visible call.
        NativeRedimStmt = "ReDim " & IIf(bPreserve, "Preserve ", "") & "(" & NativeJoinArr(a, cnt) & ")"
        Exit Function
    End If
    arr = a(2)
    nd = Val(a(4))
    If nd < 1 Then nd = 1
    For k = 0 To nd - 1
        If (6 + 2 * k) > (cnt - 1) Then Exit For       'ran out of bound pairs
        ub = a(5 + 2 * k)
        lb = a(6 + 2 * k)
        If Len(bounds) > 0 Then bounds = bounds & ", "
        bounds = bounds & lb & " To " & ub
    Next
    NativeFlagArrayGlobal arr                'a ReDim'd global is an array -> declare with "()"
    NativeRedimStmt = "ReDim " & IIf(bPreserve, "Preserve ", "") & arr & "(" & bounds & ")"
End Function

Private Function NativeFileOpenStmt(ByRef a() As String, ByVal cnt As Long) As String
    'Build `Open <path> For <mode> As #<n> [Len = <rl>]` from a __vbaFileOpen arg
    'snapshot: a() = [mode, reclen, filenum, pathname].
    Dim rl As String, fn As String, pth As String, s As String
    'Args (source order): [mode, reclen, filenum, pathname].  The pathname is the
    'deepest push and is dropped when it is an unresolved value (e.g. a parameter),
    'leaving 3 args - still render it with an <arg> placeholder for the path.
    If cnt < 3 Then NativeFileOpenStmt = "Open " & NativeJoinArr(a, cnt): Exit Function
    rl = a(1): fn = a(2)
    If cnt >= 4 Then pth = a(3) Else pth = "<arg>"
    s = "Open " & pth & " For " & NativeOpenMode(Val(a(0))) & " As #" & fn
    If rl <> "-1" And Len(rl) > 0 Then s = s & " Len = " & rl
    NativeFileOpenStmt = s
End Function

Private Function NativeOpenMode(ByVal m As Long) As String
    'Decode the VB6 __vbaFileOpen mode bitmask to the Open ... For <mode> keyword.
    Select Case True
        Case (m And &H20) <> 0: NativeOpenMode = "Binary"
        Case (m And &H4) <> 0: NativeOpenMode = "Random"
        Case (m And &H8) <> 0: NativeOpenMode = "Append"
        Case (m And &H2) <> 0: NativeOpenMode = "Output"
        Case (m And &H1) <> 0: NativeOpenMode = "Input"
        Case Else: NativeOpenMode = "Random"
    End Select
End Function

Private Function NativeJoinArr(ByRef a() As String, ByVal cnt As Long) As String
    Dim k As Long, s As String
    For k = 0 To cnt - 1
        If Len(s) > 0 Then s = s & ", "
        s = s & a(k)
    Next
    NativeJoinArr = s
End Function

Private Function NativeArgsN(ByVal nArgs As Long) As String
    'The last nArgs pushed values, in source order (last pushed = first arg).
    Dim k As Long, s As String, base As Long
    If nArgs > NVPushTop Then nArgs = NVPushTop
    If nArgs < 0 Then nArgs = 0
    base = NVPushTop - nArgs
    For k = NVPushTop - 1 To base Step -1
        If Len(s) > 0 Then s = s & ", "
        s = s & NVPushImm(k)
    Next
    NVPushTop = base
    NativeArgsN = s
End Function

Private Function NativeArgList() As String
    'All pending pushed values, in source order (drains the stack).
    NativeArgList = NativeArgsN(NVPushTop)
End Function

Private Function NativeVariantArgList(ByVal trimZero As Boolean) As String
    'Reconstruct a MsgBox / InputBox argument list.  Each pending pushed argument is
    'either a by-reference Variant pointer (resolved to the value the Variant holds)
    'or a direct immediate (e.g. the Buttons value).  Trailing missing optionals -
    'and, when trimZero is set (MsgBox), a trailing Buttons value of 0 - are dropped;
    'an omitted interior optional renders as an empty slot, e.g. InputBox(a, , c).
    Dim k As Long, cnt As Long, vals() As String, idx As Long, last As Long, s As String
    cnt = NVPushTop
    If cnt < 1 Then NVPushTop = 0: Exit Function
    ReDim vals(cnt - 1)
    idx = 0
    For k = NVPushTop - 1 To 0 Step -1              'last pushed = first argument
        If NVPushDisp(k) < 0 Then
            vals(idx) = NativeVariantVal(NVPushDisp(k))
        Else
            vals(idx) = NVPushImm(k)
        End If
        idx = idx + 1
    Next
    NVPushTop = 0
    last = idx - 1
    Do While last >= 0
        If vals(last) = NV_MISSING Then
            last = last - 1
        ElseIf trimZero And vals(last) = "0" Then
            last = last - 1
        Else
            Exit Do
        End If
    Loop
    For k = 0 To last
        If k > 0 Then s = s & ", "
        If vals(k) <> NV_MISSING Then s = s & vals(k)
    Next
    NativeVariantArgList = s
End Function

Private Function NativeArgPop() As String
    'Pop the most-recently pushed value (top of the argument stack).
    If NVPushTop > 0 Then
        NVPushTop = NVPushTop - 1
        NativeArgPop = NVPushImm(NVPushTop)
    Else
        NativeArgPop = "<arg>"
    End If
End Function

Private Function NativeFriendlyName(ByVal nm As String) As String
    NativeFriendlyName = nm
    If Left$(nm, 5) = "__vba" Then NativeFriendlyName = Mid$(nm, 6)
    If Left$(nm, 3) = "rtc" Then NativeFriendlyName = Mid$(nm, 4)
End Function

Private Sub NativeSplitProp(ByVal gp As String, ByRef propName As String, ByRef kind As String)
    Dim p As Long, namePart As String
    p = InStr(gp, " (")
    If p > 0 Then namePart = Left$(gp, p - 1) Else namePart = gp
    p = InStrRev(namePart, "_")
    If p > 0 Then propName = Mid$(namePart, p + 1) Else propName = namePart
    If InStr(gp, "Get") > 0 Then
        kind = "Get"
    ElseIf InStr(gp, "Let") > 0 Then
        kind = "Let"
    ElseIf InStr(gp, "Set") > 0 Then
        kind = "Set"
    Else
        kind = "Method"
    End If
End Sub

Private Function NativeNumFromBits(ByVal bits As Long) As String
    'Render a 32-bit constant: as a Single when its bits form a real (finite,
    'non-tiny) float, else as the integer value.  The integer is set FIRST as the
    'default so a NaN/Inf bit pattern (e.g. 0xFFFFFFFF = -1, a Boolean True) - whose
    'Int(s) would raise under On Error Resume Next and blank the result - still
    'yields a value.  isFloat is computed in one guarded step (no Int() on NaN).
    Dim s As Single, isFloat As Boolean
    On Error Resume Next
    NativeNumFromBits = CStr(bits)
    CopyMemory s, bits, 4
    isFloat = (Abs(s) >= 0.0001 And Abs(s) < 1E+18)
    If isFloat Then
        If s = Int(s) Then
            NativeNumFromBits = CStr(CLng(s))
        Else
            NativeNumFromBits = Format$(s)
        End If
    End If
End Function

Private Function NativeFloatAtAddr(ByVal va As Long) As String
    Dim s As Single, fp As Integer
    On Error Resume Next
    If va = 0 Then NativeFloatAtAddr = "?": Exit Function
    fp = FreeFile
    Open SFilePath For Binary Access Read As #fp
        Get #fp, va + 1 - OptHeader.ImageBase, s
    Close #fp
    If s = Int(s) Then NativeFloatAtAddr = CStr(CLng(s)) Else NativeFloatAtAddr = Format$(s)
End Function

Private Function NativeFirstReg(ByVal cmd As String) As String
    Dim t As String, p As Long
    t = Trim$(Mid$(cmd, 4))
    p = InStr(t, ",")
    If p > 0 Then t = Left$(t, p - 1)
    NativeFirstReg = UCase$(Trim$(t))
End Function

Private Function NativePushImmRaw(inst As CInstruction, ByRef raw As Long) As Long
    'Raw operand of a `push imm32` (0x68) or `push imm8` (0x6A) - used to recover
    'the __vbaNew Object Info / destination-global pointers.  Returns 1 (with raw)
    'when the instruction is such a push, else 0.
    Dim dump As String, n As Long, i As Long, op As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    i = NativeOpStart(dump, n)
    op = NativeDumpByte(dump, i)
    If op = &H68 Then
        raw = NativeDumpInt32(dump, i + 1): NativePushImmRaw = 1
    ElseIf op = &H6A Then
        raw = NativeDumpInt8(dump, i + 1): NativePushImmRaw = 1
    End If
End Function

Private Function NativeFileDword(ByVal va As Long) As Long
    'Read a little-endian dword at virtual address va from the image file.
    Dim fp As Integer, v As Long
    On Error Resume Next
    If va < OptHeader.ImageBase Or va >= OptHeader.ImageBase + &H1000000 Then Exit Function
    fp = FreeFile
    Open SFilePath For Binary Access Read As #fp
        Get #fp, va + 1 - OptHeader.ImageBase, v
    Close #fp
    NativeFileDword = v
End Function

Private Function NativeClassFromObjInfo(ByVal objInfoVA As Long) As String
    'Resolve a VB Object Info pointer (pushed by __vbaNew) to its class name via
    'the documented chain: ObjectInfo +0x18 = lpObject (Public Object Descriptor);
    'Public Object Descriptor +0x18 = lpszObjectName.  Validated against the known
    'project objects so a stray pointer yields "" rather than garbage.
    Dim pubDesc As Long, namePtr As Long, nm As String
    On Error Resume Next
    If objInfoVA < OptHeader.ImageBase Or objInfoVA >= OptHeader.ImageBase + &H1000000 Then Exit Function
    pubDesc = NativeFileDword(objInfoVA + &H18)
    If pubDesc = 0 Then Exit Function
    namePtr = NativeFileDword(pubDesc + &H18)
    If namePtr = 0 Then Exit Function
    nm = NativeAsciiAt(namePtr)
    If Len(nm) > 0 And NativeIsKnownObject(nm) Then NativeClassFromObjInfo = nm
End Function

Private Sub NativeScanFormFieldClasses()
    'Build (Owner:fieldOffset -> class) for `As New <class>` PRIVATE form/class
    'member fields, recovered from their auto-instantiation
    '(lea reg,[Me+off]; push reg; push <ObjInfo>; call __vbaNew).  Their type is
    'stripped from the public typeinfo, so this is the only way to resolve method
    'calls on them.  Run once (program-wide) before any proc renders.
    Set gFormFieldClass = New Collection
    On Error Resume Next
    Dim pp As Long, addr As Long, formNm As String, b() As Byte, fp As Integer
    Dim col As Collection, inst As CInstruction, arr() As CInstruction, n As Long, k As Long, j As Long
    If dsmNative Is Nothing Then Set dsmNative = New CDisassembler
    For pp = 0 To UBound(gNativeProcArray) - 1
        addr = gNativeProcArray(pp).offset
        If addr = 0 Then GoTo nextpp
        formNm = NativeFormOf(addr)
        If Len(formNm) = 0 Then GoTo nextpp
        addr = NativeSnapEntry(addr)
        ReDim b(8191)
        fp = FreeFile
        Open SFilePath For Binary Access Read As #fp
            Get #fp, addr + 1 - OptHeader.ImageBase, b
        Close #fp
        Set col = dsmNative.DisasmProc(b, addr, 8192)
        n = col.Count
        If n < 3 Then GoTo nextpp
        ReDim arr(n - 1): k = 0
        For Each inst In col: Set arr(k) = inst: k = k + 1: Next
        For k = 2 To n - 1
            If NativeIsVbaNewCall(arr(k)) Then
                Dim oi As Long, foff As Long, leaOk As Boolean, lo As Long, pim As Long, lf As Long
                oi = 0: foff = -1: leaOk = False
                lo = k - 6: If lo < 0 Then lo = 0
                For j = k - 1 To lo Step -1
                    If oi = 0 Then
                        If NativePushImmRaw(arr(j), pim) <> 0 Then If pim >= OptHeader.ImageBase Then oi = pim
                    End If
                    If Not leaOk Then If NativeLeaFieldOff(arr(j), lf) Then foff = lf: leaOk = True
                Next j
                If oi <> 0 And leaOk Then
                    Dim cls As String
                    cls = NativeClassFromObjInfo(oi)
                    If Len(cls) > 0 Then NativeColPut gFormFieldClass, formNm & ":" & foff, cls
                End If
            End If
        Next k
nextpp:
    Next pp
End Sub

Private Function NativeIsVbaNewCall(inst As CInstruction) As Boolean
    'A `call dword ptr [abs]` that resolves (via the IAT) to __vbaNew / __vbaNew2.
    Dim disp As Long, isAbs As Boolean
    On Error Resume Next
    If (inst.cmdType And C_TYPEMASK) <> C_CAL Then Exit Function
    If Not NativeDecodeDisp(inst.dump, disp, isAbs) Then Exit Function
    If Not isAbs Then Exit Function
    NativeIsVbaNewCall = (InStr(dsmNative.GetApiByIatVa(disp), "__vbaNew") > 0)
End Function

Private Function NativeLeaFieldOff(inst As CInstruction, ByRef off As Long) As Boolean
    'A `lea reg, [base + disp]` whose base is a register (not ebp, no SIB) and whose
    'disp is in the instance-field range - the &Me.<field> of an auto-instantiation.
    Dim dump As String, n As Long, i As Long, disp As Long, isAbs As Boolean, bse As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    i = NativeOpStart(dump, n)
    If NativeDumpByte(dump, i) <> &H8D Then Exit Function
    If Not NativeDecodeDisp(dump, disp, isAbs) Then Exit Function
    If isAbs Then Exit Function
    bse = NativeMemBase(dump)
    If bse < 0 Or bse = 5 Then Exit Function                'need a base register, not ebp
    If NativeMemIndex(dump) >= 0 Then Exit Function          'no SIB index
    If disp >= &H10 And disp < &H800 Then off = disp: NativeLeaFieldOff = True
End Function

Private Function NativeAsciiAt(ByVal va As Long) As String
    'Read a null-terminated ASCII string at virtual address va from the image file.
    Dim fp As Integer, b As Byte, s As String, k As Long
    On Error Resume Next
    If va < OptHeader.ImageBase Or va >= OptHeader.ImageBase + &H1000000 Then Exit Function
    fp = FreeFile
    Open SFilePath For Binary Access Read As #fp
        Do While k < 256
            Get #fp, va + 1 - OptHeader.ImageBase + k, b
            If b = 0 Then Exit Do
            If b < 32 Or b > 126 Then s = "": Exit Do      'not a clean name
            s = s & Chr$(b)
            k = k + 1
        Loop
    Close #fp
    NativeAsciiAt = s
End Function

Private Function NativeIsKnownObject(ByVal nm As String) As Boolean
    'True when nm is one of the project's objects (form / class / module).
    Dim i As Long
    On Error Resume Next
    For i = 0 To UBound(gObjectNameArray)
        If gObjectNameArray(i) = nm Then NativeIsKnownObject = True: Exit Function
    Next
End Function

Private Function NativeClassMethodAddr(ByVal className As String, ByVal off As Long) As Long
    'Address of a user class's method at vtable offset off, via the class map
    'gFormVtable("Class:off<off>") -> address.  0 when unknown (property / runtime
    'slot, or an external COM type not in this project).
    Dim v As Variant
    On Error Resume Next
    v = gFormVtable(className & ":off" & off)
    If Err.Number = 0 Then NativeClassMethodAddr = CLng(v)
End Function

Private Function NativeCallRegName(inst As CInstruction) As String
    Dim r As String
    r = UCase$(Trim$(Mid$(Trim$(inst.command), 5)))
    NativeCallRegName = NativeGetRegImport(r)
    If Len(NativeCallRegName) = 0 Then NativeCallRegName = r
End Function

Private Sub NativeSetRegImport(ByVal reg As String, ByVal nm As String)
    Dim i As Long
    i = NativeRegIndex(reg)
    If i >= 0 Then NVRegImport(i) = nm
End Sub
Private Function NativeGetRegImport(ByVal reg As String) As String
    Dim i As Long
    i = NativeRegIndex(reg)
    If i >= 0 Then NativeGetRegImport = NVRegImport(i)
End Function
Private Function NativeRegIndex(ByVal reg As String) As Long
    Select Case reg
        Case "EAX": NativeRegIndex = 0
        Case "ECX": NativeRegIndex = 1
        Case "EDX": NativeRegIndex = 2
        Case "EBX": NativeRegIndex = 3
        Case "ESP": NativeRegIndex = 4
        Case "EBP": NativeRegIndex = 5
        Case "ESI": NativeRegIndex = 6
        Case "EDI": NativeRegIndex = 7
        Case Else:  NativeRegIndex = -1
    End Select
End Function

'---------------------------------------------------------------------------
' Structured-If, entry snapping, condition hints, string literals
'---------------------------------------------------------------------------

Private Function NativeSnapEntry(ByVal addr As Long) As Long
    'The native procedure list points a few bytes into the prologue, so scan
    'backward for the standard VB6 entry "55 8B EC" (push ebp; mov ebp,esp).
    Dim b() As Byte, fp As Integer, k As Long
    On Error GoTo done
    NativeSnapEntry = addr
    ReDim b(34)                       'covers addr-32 .. addr+2
    fp = FreeFile
    Open SFilePath For Binary Access Read As #fp
        Get #fp, (addr - 32) + 1 - OptHeader.ImageBase, b
    Close #fp
    If b(32) = &H55 And b(33) = &H8B And b(34) = &HEC Then Exit Function
    For k = 31 To 0 Step -1
        If b(k) = &H55 And b(k + 1) = &H8B And b(k + 2) = &HEC Then
            NativeSnapEntry = (addr - 32) + k
            Exit Function
        End If
    Next
done:
End Function

Private Function NativeIndentStr() As String
    NativeIndentStr = Space$(4 * (1 + NVIndent))
End Function

Private Sub NativePushIf(ByVal tgt As Long)
    If NVIfTop > UBound(NVIfTarget) Then ReDim Preserve NVIfTarget(NVIfTop + 16)
    NVIfTarget(NVIfTop) = tgt
    NVIfTop = NVIfTop + 1
    NVIndent = NVIndent + 1
End Sub

Private Sub NativeCloseIfs(ByRef output As String, ByVal addr As Long)
    Do While NVIfTop > 0
        If NVIfTarget(NVIfTop - 1) > addr Then Exit Do
        NVIfTop = NVIfTop - 1
        If NVIndent > 0 Then NVIndent = NVIndent - 1
        output = output & NativeIndentStr() & "End If" & vbCrLf
    Loop
End Sub

Private Function NativeIsErrorCheckJcc(ByVal mn As String) As Boolean
    'Condition codes VB uses to skip a successful-HRESULT error check.
    Select Case mn
        Case "JGE", "JNS", "JAE", "JNB", "JNL", "JNC": NativeIsErrorCheckJcc = True
    End Select
End Function

Private Function NativeIsIfTarget(ByVal addr As Long) As Boolean
    Dim k As Long
    For k = 0 To NVIfTop - 1
        If NVIfTarget(k) = addr Then NativeIsIfTarget = True: Exit Function
    Next
End Function

Private Function NativeCondExpr(ByVal jmpMnem As String, ByVal blockGuard As Boolean) As String
    'Build the source condition for a conditional jump.  jmpMnem is the Jcc
    'mnemonic; blockGuard is True for a forward If (the block runs when the jump
    'is NOT taken, so the condition is the negation of "jump taken") and False
    'for a loop-continue / conditional-GoTo (condition = "jump taken").
    Dim op As String, L As String, R As String
    'NVCmpL is already a relational Boolean (recovered fp compare): render it
    'directly, negated with "Not" only when the block runs on its FALSE side.
    If NVCmpSet And NVCmpIsBool Then
        Dim jumpWhenTrue As Boolean
        Select Case UCase$(jmpMnem)
            Case "JNE", "JNZ": jumpWhenTrue = True       'jump taken when Boolean <> 0 (True)
            Case "JE", "JZ": jumpWhenTrue = False        'jump taken when Boolean = 0 (False)
            Case Else: GoTo notBool
        End Select
        Dim wantTrue As Boolean
        If blockGuard Then wantTrue = Not jumpWhenTrue Else wantTrue = jumpWhenTrue
        NVCmpSet = False: NVCmpIsBool = False: NVLastCmp = ""
        If wantTrue Then NativeCondExpr = NVCmpL Else NativeCondExpr = "Not " & NVCmpL
        Exit Function
    End If
notBool:
    If NVCmpSet Then
        op = NativeJccOp(jmpMnem)
        If blockGuard Then op = NativeNegOp(op)
        L = NVCmpL
        If NVCmpIsTest Then R = "0" Else R = NVCmpR
        NVCmpSet = False: NVCmpIsBool = False: NVLastCmp = ""
        If Len(op) > 0 And Len(L) > 0 And Len(R) > 0 Then
            NativeCondExpr = L & " " & op & " " & R
            Exit Function
        End If
    End If
    'FPU compare hint (FCOM) or nothing.
    If Len(NVLastCmp) > 0 Then
        NativeCondExpr = NVLastCmp: NVLastCmp = "": Exit Function
    End If
    NativeCondExpr = "<cond>"
End Function

Private Function NativeJccOp(ByVal mn As String) As String
    'Relational operator for "the jump is taken" (left <op> right).
    Select Case UCase$(mn)
        Case "JE", "JZ": NativeJccOp = "="
        Case "JNE", "JNZ": NativeJccOp = "<>"
        Case "JL", "JNGE", "JB", "JC", "JNAE": NativeJccOp = "<"
        Case "JLE", "JNG", "JBE", "JNA": NativeJccOp = "<="
        Case "JG", "JNLE", "JA", "JNBE": NativeJccOp = ">"
        Case "JGE", "JNL", "JAE", "JNB", "JNC": NativeJccOp = ">="
        Case "JS": NativeJccOp = "<"          'sign set ~ negative
        Case "JNS": NativeJccOp = ">="
        Case Else: NativeJccOp = ""
    End Select
End Function

Private Function NativeNegOp(ByVal op As String) As String
    Select Case op
        Case "=": NativeNegOp = "<>"
        Case "<>": NativeNegOp = "="
        Case "<": NativeNegOp = ">="
        Case "<=": NativeNegOp = ">"
        Case ">": NativeNegOp = "<="
        Case ">=": NativeNegOp = "<"
        Case Else: NativeNegOp = op
    End Select
End Function

Private Function NativeRegName(ByVal idx As Long) As String
    Select Case idx
        Case 0: NativeRegName = "eax"
        Case 1: NativeRegName = "ecx"
        Case 2: NativeRegName = "edx"
        Case 3: NativeRegName = "ebx"
        Case 4: NativeRegName = "esp"
        Case 5: NativeRegName = "ebp"
        Case 6: NativeRegName = "esi"
        Case 7: NativeRegName = "edi"
    End Select
End Function

Private Function NativeRegVal(ByVal idx As Long) As String
    'Tracked symbolic value of a register, else its raw name.
    If idx >= 0 And idx <= 7 Then
        If Len(NVReg(idx)) > 0 Then NativeRegVal = NVReg(idx) Else NativeRegVal = NativeRegName(idx)
    End If
End Function

Private Function NativeMemBase(ByVal dump As String) As Long
    'Base register index of a single-opcode ModR/M memory operand, or -1 for an
    'absolute / register-direct / SIB-without-base operand.
    Dim n As Long, i As Long, op As Long, modrm As Long, md As Long, rm As Long, sib As Long
    On Error GoTo none
    NativeMemBase = -1
    dump = Replace(dump, " ", "")
    n = Len(dump) \ 2
    i = NativeOpStart(dump, n)
    op = NativeDumpByte(dump, i): i = i + 1
    If op = &HF Or op = &HE8 Or op = &HE9 Or op = &HEB Then Exit Function
    If i >= n Then Exit Function
    modrm = NativeDumpByte(dump, i)
    md = (modrm \ &H40) And 3: rm = modrm And 7
    If md = 3 Then Exit Function
    If rm = 4 Then
        sib = NativeDumpByte(dump, i + 1)
        If md = 0 And (sib And 7) = 5 Then Exit Function    'no base
        NativeMemBase = sib And 7
        Exit Function
    End If
    If md = 0 And rm = 5 Then Exit Function                 'abs [disp32]
    NativeMemBase = rm
    Exit Function
none:
    NativeMemBase = -1
End Function

Private Function NativeMemIndex(ByVal dump As String) As Long
    'Index register (0..7) of a SIB memory operand, or -1 if none / no SIB byte.
    'Used to keep the deref-with-offset rendering (base(offset)) off genuinely
    'indexed array-element operands [base + index*scale + disp], which we leave
    'unmodelled (the commercial decompiler also punts on those).
    Dim n As Long, i As Long, op As Long, modrm As Long, md As Long, rm As Long, sib As Long, idx As Long
    On Error GoTo none
    NativeMemIndex = -1
    dump = Replace(dump, " ", "")
    n = Len(dump) \ 2
    i = NativeOpStart(dump, n)
    op = NativeDumpByte(dump, i): i = i + 1
    If op = &HF Or op = &HE8 Or op = &HE9 Or op = &HEB Then Exit Function
    If i >= n Then Exit Function
    modrm = NativeDumpByte(dump, i)
    md = (modrm \ &H40) And 3: rm = modrm And 7
    If md = 3 Or rm <> 4 Then Exit Function              'register-direct or no SIB
    sib = NativeDumpByte(dump, i + 1)
    idx = (sib \ 8) And 7
    If idx <> 4 Then NativeMemIndex = idx                'idx = 4 encodes "no index"
    Exit Function
none:
    NativeMemIndex = -1
End Function

Private Function NativeIsPtrExpr(ByVal s As String) As Boolean
    'True when s is a DATA-pointer expression we are willing to dereference at a
    'byte offset: a module global / parameter / local / code label, optionally
    'already a chained deref (global_X(12)).  Deliberately EXCLUDES bare registers
    '(edx), pure numbers, and property / control expressions (frmMain.Text1) - for
    'those a byte-offset deref would be misleading, so they stay as their raw form.
    Dim p As String
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = "(" Then p = Mid$(s, 2) Else p = s      'peel a leading paren of a folded expr
    NativeIsPtrExpr = (Left$(p, 7) = "global_") Or (Left$(p, 4) = "arg_") _
                   Or (Left$(p, 4) = "var_") Or (Left$(p, 4) = "loc_")
End Function

Private Function NativeIsNumLit(ByVal s As String) As Boolean
    'True when s is a plain numeric constant (leading sign or digit).
    Dim c As String
    If Len(s) = 0 Then Exit Function
    c = Left$(s, 1)
    NativeIsNumLit = (c = "-") Or (c >= "0" And c <= "9")
End Function

Private Function NativeIsDerefBase(ByVal s As String) As Boolean
    'True when s may be DEREFERENCED at a byte offset - a genuine data pointer or a
    'pointer-deref chain (global_X, arg_C, var_18, global_X(12)).  Unlike
    'NativeIsPtrExpr it does NOT peel a leading paren, so a FOLDED arithmetic result
    '(arg_C - global_X(20)) - a value, not a pointer - is rejected and never
    'rendered as the misleading (arg_C - global_X(20))(16).
    NativeIsDerefBase = (Left$(s, 7) = "global_") Or (Left$(s, 4) = "arg_") _
                     Or (Left$(s, 4) = "var_") Or (Left$(s, 4) = "loc_")
End Function

Private Function NativeGlobalObjByOffset(ByVal disp As Long) As String
    'VB6 form-interface vtable slots for the intrinsic global objects.
    Select Case disp
        Case &H14: NativeGlobalObjByOffset = "App"
        Case &H18: NativeGlobalObjByOffset = "Screen"
        Case &H1C: NativeGlobalObjByOffset = "Clipboard"
    End Select
End Function

Private Function NativeFormMethodByOffset(ByVal disp As Long) As String
    'Built-in VB6 _Form-interface methods at fixed vtable offsets, below the
    'control-accessor block (0x2F8) and the user-method block (0x6F8).  These are
    'part of the form runtime interface and stable across programs.  Verified by
    'tracing real forms (Dungeon frmMainMenu cmdExit/cmdLoad/cmdNew all
    '`call [Me_vt + 0x2B4]`) and cross-checking the commercial decompiler's Me.Hide.
    'Only argument-less, value-less methods belong here (rendered as a bare
    'statement Me.<method>); extend as more offsets are confirmed by tracing.
    Select Case disp
        Case &H2B4: NativeFormMethodByOffset = "Hide"
    End Select
End Function

Private Function NativeErrPropByOffset(ByVal disp As Long) As String
    'Read-only property GETTER vtable offsets on the VB6 Err object (the _ErrObject
    'interface).  Err.Number is call [ErrVt + 0x1C], Err.Description is [ErrVt + 0x2C]
    '- stable across VB6 programs, verified against VB6LangTest.
    Select Case disp
        Case &H1C: NativeErrPropByOffset = "Number"
        Case &H2C: NativeErrPropByOffset = "Description"
        Case &H20: NativeErrPropByOffset = "Source"
        Case &H30: NativeErrPropByOffset = "HelpFile"
        Case &H34: NativeErrPropByOffset = "HelpContext"
        Case &H28: NativeErrPropByOffset = "LastDllError"
    End Select
End Function

Private Function NativeIsIntrinsicObj(ByVal s As String) As Boolean
    'True for the VB6 intrinsic global object names whose property vtables we know.
    Select Case s
        Case "App", "Screen", "Clipboard": NativeIsIntrinsicObj = True
    End Select
End Function

Private Function NativeIntrinsicPropByOffset(ByVal obj As String, ByVal disp As Long) As String
    'Read-only property GETTER vtable offsets on a VB6 intrinsic object (e.g.
    'App.Path is call [App_vtable + 0x50]).  These offsets are part of the runtime
    '_App / _Screen interfaces and are stable across every VB6 program; each is
    'verified by tracing real binaries (pMasterMaker, VB6LangTest).  Offsets are
    'per-interface, so Screen.Height (0x50) and App.Path (0x50) legitimately
    'coincide - keying by object name keeps them distinct.  Only read-only GETs
    'belong here: the resolver surfaces them as `var_X = obj.Prop`, so a writable
    'property (e.g. Screen.MousePointer, a Let) or an arg-taking method (Clipboard
    'SetText/GetText) must NOT be listed - it would drop the assignment/args.
    Select Case obj
        Case "App"
            Select Case disp
                Case &H50: NativeIntrinsicPropByOffset = "Path"
                Case &H58: NativeIntrinsicPropByOffset = "EXEName"
                Case &H60: NativeIntrinsicPropByOffset = "Title"
                Case &HB8: NativeIntrinsicPropByOffset = "Major"
                Case &HC0: NativeIntrinsicPropByOffset = "Minor"
                Case &HC8: NativeIntrinsicPropByOffset = "Revision"
            End Select
        Case "Screen"
            Select Case disp
                Case &H98: NativeIntrinsicPropByOffset = "Width"
                Case &H50: NativeIntrinsicPropByOffset = "Height"
                Case &H80: NativeIntrinsicPropByOffset = "TwipsPerPixelX"
            End Select
    End Select
End Function

Private Function NativeEaxUse(inst As CInstruction) As Long
    'How this instruction uses eax (the previous call's result):
    '  1 = reads eax as a source  -> a deferred call folds into this instruction
    '  2 = clobbers eax / control-flow boundary -> emit the deferred call now
    '  0 = does not touch eax      -> keep the call deferred
    'Unknown opcodes fall through to 0 so a deferred call is never lost early - it
    'is emitted at the next boundary or the end of the procedure instead.
    Dim cls As Long, dump As String, n As Long, i As Long, op As Long
    Dim modrm As Long, md As Long, reg As Long, rm As Long
    On Error GoTo none0
    cls = inst.cmdType And C_TYPEMASK
    If cls = C_CAL Or cls = C_JMP Or cls = C_JMC Or cls = C_RET Then NativeEaxUse = 2: Exit Function
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    If n < 1 Then Exit Function
    i = NativeOpStart(dump, n)
    op = NativeDumpByte(dump, i)
    Select Case op
        Case &H50: NativeEaxUse = 1: Exit Function          'push eax (reads)
        Case &H58: NativeEaxUse = 2: Exit Function          'pop eax (clobbers)
        Case &HB8: NativeEaxUse = 2: Exit Function          'mov eax, imm32
        Case &HA1: NativeEaxUse = 2: Exit Function          'mov eax, [abs]
        Case &H3D, &HA9: NativeEaxUse = 1: Exit Function     'cmp/test eax, imm (reads)
    End Select
    If n >= i + 2 Then
        modrm = NativeDumpByte(dump, i + 1)
        md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
        Select Case op
            Case &H89                          'mov r/m, r (store reg)
                If reg = 0 Then NativeEaxUse = 1: Exit Function          'eax is source
                If md = 3 And rm = 0 Then NativeEaxUse = 2: Exit Function 'eax is dest
            Case &H8B                          'mov r, r/m
                If reg = 0 Then NativeEaxUse = 2: Exit Function          'eax is dest
                If md = 3 And rm = 0 Then NativeEaxUse = 1: Exit Function 'eax is source
            Case &H85, &H3B, &H39              'test/cmp involving a register
                If reg = 0 Or (md = 3 And rm = 0) Then NativeEaxUse = 1: Exit Function
            Case &H83, &H81                    'grp1 r/m, imm (cmp eax, imm reads)
                If md = 3 And rm = 0 Then NativeEaxUse = 1: Exit Function
            Case &H01, &H03, &H09, &H0B, &H21, &H23, &H29, &H2B, &H31, &H33
                'alu op with eax as an operand: reads then clobbers eax, and we do
                'not model the arithmetic - emit the call rather than fold it.
                If reg = 0 Or (md = 3 And rm = 0) Then NativeEaxUse = 2: Exit Function
        End Select
    End If
none0:
    NativeEaxUse = 0
End Function

Private Function NativeRmVal(ByVal dump As String, ByVal md As Long, ByVal rm As Long) As String
    'Symbolic value of a ModR/M r/m operand: a register's tracked value, a local
    'stack slot, or a deref-with-offset of a tracked pointer register rendered in
    'the commercial decompiler's base(offset) style.  "" when we cannot model it.
    Dim disp As Long, isAbs As Boolean, bse As Long
    If md = 3 Then
        NativeRmVal = NativeRegVal(rm)
    ElseIf NativeDecodeDisp(dump, disp, isAbs) Then
        If Not isAbs And disp < 0 Then
            NativeRmVal = NativeGetLocalExpr(disp)
        ElseIf isAbs And NativeIsGlobalAddr(disp) Then
            NativeRmVal = NativeGlobalName(disp)
        ElseIf Not isAbs And disp >= 8 And disp <= &H200 And NativeMemBase(dump) = 5 Then
            NativeRmVal = "arg_" & Hex$(disp)        'a procedure parameter (ebp+positive)
        ElseIf Not isAbs And NativeMemIndex(dump) < 0 Then
            'Deref-with-offset of a tracked pointer: [base + disp] -> <baseVal>(disp),
            'or just <baseVal> at disp 0.  Gated to a single base register (no SIB
            'index) currently holding a known value, so it never fires on an operand
            'we have not modelled.  Subsumes the old ByRef-param deref ([arg_reg])
            'and adds struct / SAFEARRAY field reads ([global_ptr + fieldOff]); the
            'chain global_X -> global_X(12) -> global_X(12)(20) forms via NVReg.
            'The base must be a DATA-pointer expression (global / arg / var / loc,
            'incl. chained derefs), enforced by NativeIsPtrExpr: a constant "0"
            'base would render the nonsense 0(20), and a control/property base a
            'misleading frmMain.Text1(20) - both stay as their raw register instead.
            bse = NativeMemBase(dump)
            If bse >= 0 And bse <= 7 And NativeIsDerefBase(NVReg(bse)) Then
                If disp = 0 Then
                    NativeRmVal = NVReg(bse)
                Else
                    NativeRmVal = NVReg(bse) & "(" & CStr(disp) & ")"
                End If
            End If
        ElseIf Not isAbs And NativeMemIndex(dump) >= 0 Then
            'SIB [base + index*scale + disp] -> a SAFEARRAY / array element access.
            'One register holds the array DATA pointer (a module-level global in
            'this code's arrays), the other the element index.  Render in the
            'base(offset) style as <ptr>(<index>) [ (fieldDisp) ].  Fire only when
            'exactly one register holds a global_ pointer and the other a clean
            'index value (not a bare, untracked register), so we never emit
            'nonsense like eax(ecx); those stay <cond>.
            Dim sb As Long, si As Long, pv As String, iv As String
            sb = NativeMemBase(dump): si = NativeMemIndex(dump)
            If sb >= 0 And sb <= 7 And si >= 0 And si <= 7 Then
                If Left$(NVReg(sb), 7) = "global_" And Left$(NVReg(si), 7) <> "global_" Then
                    pv = NVReg(sb): iv = NVReg(si)
                ElseIf Left$(NVReg(si), 7) = "global_" And Left$(NVReg(sb), 7) <> "global_" Then
                    pv = NVReg(si): iv = NVReg(sb)
                End If
                If Len(pv) > 0 And NativeIsCleanIndex(iv) Then
                    If disp = 0 Then
                        NativeRmVal = pv & "(" & iv & ")"
                    Else
                        NativeRmVal = pv & "(" & iv & ")(" & CStr(disp) & ")"
                    End If
                End If
            End If
        End If
    End If
End Function

Private Function NativeIsCleanIndex(ByVal s As String) As Boolean
    'A SAFEARRAY element index we are willing to render inside base(index): a
    'number, a local / parameter / global, or a folded arithmetic expression -
    'but NOT a bare untracked register (eax, ecx, ...) which would be meaningless.
    If Len(s) = 0 Then Exit Function
    Select Case s
        Case "eax", "ecx", "edx", "ebx", "esp", "ebp", "esi", "edi": Exit Function
    End Select
    NativeIsCleanIndex = True
End Function

Private Sub NativeDecodeCompare(inst As CInstruction, ByVal mn As String)
    'Decode a TEST/CMP into symbolic left/right operands for the next Jcc.
    Dim dump As String, n As Long, i As Long, op As Long
    Dim modrm As Long, md As Long, reg As Long, rm As Long
    Dim L As String, R As String, isTst As Boolean
    On Error GoTo done3
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    i = NativeOpStart(dump, n)
    op = NativeDumpByte(dump, i)
    isTst = (mn = "TEST")
    modrm = NativeDumpByte(dump, i + 1)
    md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
    'test reg,reg where the register holds a recovered EXPRESSION (a relational
    'Boolean from an fp comparison, or a folded value): test it for non-zero.
    'This is the standard "If <fp compare> Then" tail (test ax,ax / test eax,eax)
    'and resolves what the 0x66-register guard below would otherwise drop.
    If op = &H85 And md = 3 And reg = rm Then
        Dim bv As String
        bv = NVReg(rm)
        If Len(bv) > 0 And Left$(bv, 1) = "(" Then
            If NativeLooksRelational(bv) Then
                'Already a relational Boolean (recovered fp compare) - render it
                'directly, without the redundant "<> 0".
                NVCmpL = bv: NVCmpIsBool = True: NVCmpSet = True
            Else
                'A folded value (e.g. an arithmetic sum) - test it for non-zero.
                NVCmpL = bv: NVCmpR = "0": NVCmpIsTest = True: NVCmpSet = True
            End If
            Exit Sub
        ElseIf NativeHas66(dump) And NativeIsCleanNamedVal(bv) Then
            'A 16-bit `test si,si` of a register holding a clean tracked local /
            'parameter / global (var_X / arg_X / global_X).  The register was loaded
            'by a FULL 32-bit mov (a 16-bit mov clears the register's tracked value
            'above), and VB only emits a word test for an Integer/Boolean variable, so
            'the low word IS the value: resolve `<name> <> 0` instead of leaving <cond>.
            NVCmpL = bv: NVCmpR = "0": NVCmpIsTest = True: NVCmpSet = True
            Exit Sub
        End If
    End If
    'A 16-bit (0x66) compare whose operand is a REGISTER is part of VB6's Boolean
    'evaluation juggling (setcc/neg/mov si,ax/cmp si,bx) - its register operands are
    'low-word partials our 32-bit model can't follow, so resolving them yields
    'garbage like `0 <> 0`.  Resolve a word compare ONLY when its r/m is MEMORY (the
    'Integer/Boolean FIELD compares, e.g. global_X(12) <> 0); otherwise leave <cond>.
    If NativeHas66(dump) Then
        Select Case op
            Case &H3D, &HA9: GoTo done3                  'cmp/test ax, imm - 16-bit accumulator
            Case Else: If md = 3 Then GoTo done3          'r/m is a register
        End Select
    End If
    Select Case op
        Case &H85                       'test r/m32, r32
            L = NativeRmVal(dump, md, rm): R = "0": isTst = True
        Case &HA9                       'test eax, imm32
            L = NativeRegVal(0): R = "0": isTst = True
        Case &HF7                       'test r/m32, imm32 (reg field = 0)
            L = NativeRmVal(dump, md, rm): R = "0": isTst = True
        Case &H3B                       'cmp r32, r/m32
            L = NativeRegVal(reg): R = NativeRmVal(dump, md, rm)
        Case &H39                       'cmp r/m32, r32
            L = NativeRmVal(dump, md, rm): R = NativeRegVal(reg)
        Case &H3D                       'cmp eax, imm32
            L = NativeRegVal(0): R = NativeNumFromBits(NativeDumpInt32(dump, i + 1))
        Case &H83                       'cmp r/m32, imm8 (sign-extended)
            L = NativeRmVal(dump, md, rm): R = CStr(NativeDumpInt8(dump, n - 1))
        Case &H81                       'cmp r/m32, imm32
            L = NativeRmVal(dump, md, rm): R = NativeNumFromBits(NativeDumpInt32(dump, n - 4))
    End Select
    If Len(L) > 0 And Len(R) > 0 Then
        NVCmpL = L: NVCmpR = R: NVCmpIsTest = isTst: NVCmpSet = True
    End If
done3:
End Sub

Private Sub NativeTrackVariantStore(inst As CInstruction)
    'mov [local], imm32 (opcode C7 /0): record the stored value against its stack
    'slot.  A string-constant pointer becomes the quoted literal (a Variant's BSTR
    'data field); any other immediate is kept as a number (e.g. a VT tag such as 8
    'for VT_BSTR).  Only local slots (disp < 0) are tracked.
    Dim dump As String, op As Long, n As Long, i As Long, imm As Long, disp As Long, isAbs As Boolean, slit As String
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    If n < 5 Then Exit Sub
    i = 0
    Do While i < n
        op = NativeDumpByte(dump, i)
        Select Case op
            Case &H66, &H67, &HF0, &HF2, &HF3, &H26, &H2E, &H36, &H3E, &H64, &H65: i = i + 1
            Case Else: Exit Do
        End Select
    Loop
    If NativeDumpByte(dump, i) <> &HC7 Then Exit Sub
    If Not NativeDecodeDisp(dump, disp, isAbs) Then Exit Sub
    If isAbs Or disp >= 0 Then Exit Sub
    imm = NativeDumpInt32(dump, n - 4)            'imm32 is the trailing dword
    If imm >= OptHeader.ImageBase Then slit = NativeStringAt(imm)
    If Len(slit) > 0 Then
        NativeSetVSlot disp, slit
    Else
        NativeSetVSlot disp, NativeNumFromBits(imm)
    End If
End Sub

Private Function NativeMovStringLit(inst As CInstruction) As String
    'mov [mem], imm32 (opcode C7): if the immediate is a pointer to a readable
    'string constant, return that quoted literal.
    Dim dump As String, op As Long, n As Long, i As Long, imm As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    If n < 5 Then Exit Function
    i = 0
    Do While i < n
        op = NativeDumpByte(dump, i)
        Select Case op
            Case &H66, &H67, &HF0, &HF2, &HF3, &H26, &H2E, &H36, &H3E, &H64, &H65: i = i + 1
            Case Else: Exit Do
        End Select
    Loop
    If NativeDumpByte(dump, i) <> &HC7 Then Exit Function
    imm = NativeDumpInt32(dump, n - 4)            'imm32 is the trailing dword
    If imm >= OptHeader.ImageBase Then NativeMovStringLit = NativeStringAt(imm)
End Function

Private Function NativeCallTarget(inst As CInstruction) As Long
    'Absolute target of a direct (E8) relative call, else olly's computed
    'target, else 0 (register/indirect call).
    Dim dump As String, n As Long, i As Long, rel As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    If n >= 5 Then
        i = NativeOpStart(dump, n)
        If NativeDumpByte(dump, i) = &HE8 Then
            rel = NativeDumpInt32(dump, i + 1)
            NativeCallTarget = inst.va + inst.instLen + rel
            Exit Function
        End If
    End If
    If inst.jmpConst <> 0 Then NativeCallTarget = inst.jmpConst
End Function

Private Function NativeInImage(ByVal addr As Long) As Boolean
    'Within this EXE's address range (excludes msvbvm60 et al. at high addresses).
    NativeInImage = (addr >= OptHeader.ImageBase) And (addr < OptHeader.ImageBase + &H1000000)
End Function

Private Function NativeGlobalName(ByVal va As Long) As String
    'Synthetic stable name for a module-level global / static at an absolute .data
    'address, matching the commercial decompiler's scheme (8 hex digits).
    NativeGlobalName = "global_" & Right$("00000000" & Hex$(va), 8)
    'Record the reference so the owning module can declare it at its top.
    On Error Resume Next
    If gNativeUsedGlobal Is Nothing Then Set gNativeUsedGlobal = New Collection
    gNativeUsedGlobal.Add va, "g" & va
End Function

Private Sub NativeFlagArrayGlobal(ByVal expr As String)
    'Mark a global as an array (it was ReDim'd) so its declaration gets "()".
    Dim p As Long, hx As String, va As Long
    On Error Resume Next
    If Left$(expr, 7) <> "global_" Then Exit Sub
    hx = Mid$(expr, 8, 8)
    va = CLng("&H" & hx)
    If gNativeArrayGlobal Is Nothing Then Set gNativeArrayGlobal = New Collection
    gNativeArrayGlobal.Add va, "g" & va
End Sub

Private Function NativeIsGlobalAddr(ByVal va As Long) As Boolean
    'True when va is an absolute address inside a NON-executable (data) section of
    'this image - a module-level global / static, not code - so it can be rendered
    'as global_XXXXXXXX rather than a bare number.  Section fields are unsigned
    'DWORDs held in Doubles; test the IMAGE_SCN_MEM_EXECUTE bit (0x20000000) by
    'division to avoid a Long overflow on data-section characteristics (0xC0000040).
    Dim rva As Long, k As Long
    If va < OptHeader.ImageBase Then Exit Function
    rva = va - OptHeader.ImageBase
    For k = 0 To MAXSECTIONS
        If SecHeader(k).SizeRawData > 0 And SecHeader(k).Address > 0 Then
            If rva >= SecHeader(k).Address And rva < SecHeader(k).Address + SecHeader(k).SizeRawData Then
                NativeIsGlobalAddr = ((Int(SecHeader(k).Properties / &H20000000) Mod 2) = 0)
                Exit Function
            End If
        End If
    Next
End Function

Private Function NativeCallTargetName(ByVal tgt As Long) As String
    'Resolve a call target address to a procedure name, qualified by module
    'unless it is in the current form (then unqualified).
    Dim nm As String
    nm = NativeLookupName(tgt)
    If Len(nm) = 0 Then nm = NativeLookupName(NativeSnapEntry(tgt))
    If Len(nm) = 0 Then
        'Not a user procedure - it may be a declared-DLL (Win32 API) call thunk.
        nm = NativeApiStubName(tgt)
        If Len(nm) > 0 Then NativeCallTargetName = nm: Exit Function
        NativeCallTargetName = "proc_" & Hex$(tgt)
        Exit Function
    End If
    If Len(NVForm) > 0 And Left$(nm, Len(NVForm) + 1) = NVForm & "." Then nm = Mid$(nm, Len(NVForm) + 2)
    NativeCallTargetName = nm
End Function

Private Function NativeApiStubName(ByVal addr As Long) As String
    'Resolve a direct call target that is a VB6 declared-DLL (Win32 API) call
    'thunk to its API name (e.g. FindWindow, SendMessage), or "" if addr is not
    'such a stub.
    '
    'A "Declare Function ... Lib ..." compiles to a DllFunctionCall thunk:
    '    A1 <cachedPtr>          mov  eax,[cachedPtr]   (0 until first runtime call)
    '    0B C0                   or   eax,eax
    '    74 02                   jz   +2
    '    FF E0                   jmp  eax
    '    68 <descriptorVA>       push <descriptorVA>    (the import descriptor)
    '    B8 <DllFunctionCall> / FF D0
    'The descriptor is the VB external-library struct: +0x00 -> library-name ptr,
    '+0x04 -> function-name ptr (ANSI, e.g. "FindWindowA").  Self-contained: no
    'reliance on the PE import table (declared APIs are DllFunctionCall'd, not PE
    'imports) nor on the external-table parse order.
    On Error GoTo done
    If NVApiStubCache Is Nothing Then Set NVApiStubCache = New Collection

    Dim cached As Variant
    On Error Resume Next
    cached = NVApiStubCache("k" & addr)
    If Err.Number = 0 Then NativeApiStubName = cached: Exit Function
    Err.Clear
    On Error GoTo done

    Dim result As String
    result = ""
    If addr >= OptHeader.ImageBase Then
        Dim fp As Integer, b(19) As Byte, pos As Long
        fp = FreeFile
        Open SFilePath For Binary Access Read As #fp
        pos = addr + 1 - OptHeader.ImageBase
        If pos >= 1 And pos + 19 <= LOF(fp) Then
            Get #fp, pos, b
            'Match the thunk signature (mov eax,[imm]; or eax,eax; jz; jmp eax; push imm32).
            If b(0) = &HA1 And b(5) = &HB And b(6) = &HC0 And b(11) = &H68 Then
                Dim descVA As Long, nameVA As Long
                descVA = NativeBytesToLong(b(12), b(13), b(14), b(15))
                If descVA >= OptHeader.ImageBase Then
                    'Function-name pointer at descriptor+0x04.
                    Dim np As Long
                    np = descVA + 4 + 1 - OptHeader.ImageBase
                    If np >= 1 And np + 3 <= LOF(fp) Then
                        Dim nb(3) As Byte
                        Get #fp, np, nb
                        nameVA = NativeBytesToLong(nb(0), nb(1), nb(2), nb(3))
                        If nameVA >= OptHeader.ImageBase Then
                            Dim sp As Long
                            sp = nameVA + 1 - OptHeader.ImageBase
                            If sp >= 1 And sp <= LOF(fp) Then
                                Seek #fp, sp
                                result = NativeApiTrimSuffix(GetUntilNull(fp))
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Close #fp
    End If

    NVApiStubCache.Add result, "k" & addr
    NativeApiStubName = result
    Exit Function
done:
    On Error Resume Next
    Close #fp
    NativeApiStubName = ""
End Function

Private Function NativeBytesToLong(ByVal b0 As Byte, ByVal b1 As Byte, ByVal b2 As Byte, ByVal b3 As Byte) As Long
    'Assemble a little-endian DWORD without overflow on the high bit.
    Dim hi As Long
    hi = b3
    NativeBytesToLong = (CLng(b0) Or (CLng(b1) * &H100&) Or (CLng(b2) * &H10000)) Or (hi * &H1000000)
End Function

Private Function NativeApiTrimSuffix(ByVal nm As String) As String
    'Drop a single trailing A/W (ANSI/Unicode variant suffix) so the name matches
    'the VB Declare alias / commercial output (FindWindowA -> FindWindow).  Names
    'without the suffix (FatalExit, IsDebuggerPresent) are returned unchanged.
    Dim L As Long
    L = Len(nm)
    If L > 1 Then
        Dim c As String
        c = Right$(nm, 1)
        If c = "A" Or c = "W" Then NativeApiTrimSuffix = Left$(nm, L - 1): Exit Function
    End If
    NativeApiTrimSuffix = nm
End Function

Private Function NativeLookupName(ByVal addr As Long) As String
    'Nearest named procedure at or just before addr (tolerance for the few
    'prologue bytes between the call target and the named entry).
    Dim i As Long, d As Long, bestDelta As Long
    On Error Resume Next
    bestDelta = 99999
    For i = 0 To UBound(SubNamelist)
        d = addr - SubNamelist(i).offset
        If d >= 0 And d <= 24 And d < bestDelta Then bestDelta = d: NativeLookupName = SubNamelist(i).strName
    Next
End Function

Private Function NativeJoinAmp(ByVal a As String, ByVal b As String) As String
    'Join two string-expression parts with VB's concatenation operator.
    If Len(a) = 0 Then NativeJoinAmp = b Else NativeJoinAmp = a & " & " & b
End Function

Private Function NativeStringAt(ByVal va As Long) As String
    'Read a VB6 string constant (BSTR) from the image and render it as VB source.
    'When the BSTR length prefix (the 4 bytes before the data) is valid, control-
    'character constants resolve - vbCrLf / vbCr / vbLf / vbTab - and a mixed
    'string renders as "lit" & vbCrLf & "lit" (so `"Dungeon Fate" & vbCrLf` forms
    'instead of leaking the BSTR pointer as a number).  Falls back to the
    'printable-only scan when the length prefix is not a clean BSTR.  Returns ""
    'when the address does not hold clean text.
    Dim fp As Integer, pos As Long, byteLen As Long, nch As Long, k As Long
    Dim ch As Integer, ch2 As Integer, c As Long, res As String, lit As String, okBStr As Boolean, cc As String
    On Error GoTo done
    If va < OptHeader.ImageBase Then Exit Function
    fp = FreeFile
    Open SFilePath For Binary Access Read As #fp
    pos = va + 1 - OptHeader.ImageBase
    If pos >= 5 And pos <= LOF(fp) Then
        Get #fp, pos - 4, byteLen                       'BSTR byte-length prefix
        okBStr = (byteLen >= 2 And byteLen <= 1024 And (byteLen And 1) = 0)
        If okBStr Then
            nch = byteLen \ 2
            For k = 0 To nch - 1                         'must be all tab/cr/lf or printable
                Get #fp, pos + k * 2, ch
                c = ch: If c < 0 Then c = c + 65536
                If Not (c = 9 Or c = 10 Or c = 13 Or (c >= 32 And c <= 126)) Then okBStr = False: Exit For
            Next
        End If
        If okBStr Then
            k = 0
            Do While k < nch
                Get #fp, pos + k * 2, ch
                c = ch: If c < 0 Then c = c + 65536
                cc = ""
                If c = 13 And k < nch - 1 Then           'CR followed by LF -> vbCrLf
                    Get #fp, pos + (k + 1) * 2, ch2
                    If ch2 = 10 Then cc = "vbCrLf": k = k + 1
                End If
                If Len(cc) = 0 Then
                    Select Case c
                        Case 13: cc = "vbCr"
                        Case 10: cc = "vbLf"
                        Case 9: cc = "vbTab"
                    End Select
                End If
                If Len(cc) > 0 Then
                    If Len(lit) > 0 Then res = NativeJoinAmp(res, Chr$(34) & lit & Chr$(34)): lit = ""
                    res = NativeJoinAmp(res, cc)
                Else
                    lit = lit & Chr$(c)
                End If
                k = k + 1
            Loop
            If Len(lit) > 0 Then res = NativeJoinAmp(res, Chr$(34) & lit & Chr$(34))
        Else
            Do                                           'fallback: printable-only scan
                Get #fp, pos, ch
                If ch = 0 Then Exit Do
                If ch < 32 Or ch > 126 Then Exit Do
                lit = lit & Chr$(ch)
                pos = pos + 2
                If Len(lit) > 256 Then Exit Do
            Loop
            If Len(lit) >= 1 Then res = Chr$(34) & lit & Chr$(34)
        End If
    End If
    Close #fp
    NativeStringAt = res
    Exit Function
done:
    On Error Resume Next
    Close #fp
End Function

Private Function NativeStrLitArgs() As String
    Dim x As Variant, s As String
    On Error Resume Next
    For Each x In NVStrLits
        If Len(s) > 0 Then s = s & ", "
        s = s & x
    Next
    Set NVStrLits = New Collection
    NativeStrLitArgs = s
End Function

Private Function NativeProcMatchIdx(ByVal addr As Long) As Long
    'Index of the nearest named SubNamelist entry at or just before this entry
    'is (the named proc list and the clickable list can differ by a few prologue
    'bytes), or -1 if none.  Shared by NativeProcName and NativeProcHeader so the
    'name, kind and visibility all come from the same matched entry.
    Dim i As Long, d As Long, bestDelta As Long, best As Long
    On Error Resume Next
    bestDelta = 99999: best = -1
    For i = 0 To UBound(SubNamelist)
        d = addr - SubNamelist(i).offset
        If d >= 0 And d <= 24 And d < bestDelta Then bestDelta = d: best = i
    Next
    NativeProcMatchIdx = best
End Function

Private Function NativeProcHeader(ByVal addr As Long) As String
    Dim nm As String, idx As Long, vis As String, kindStr As String
    nm = NativeProcName(addr)
    vis = "Private": kindStr = "Sub": NVProcEndWord = "Sub"
    idx = NativeProcMatchIdx(addr)
    If idx >= 0 Then
        If Len(SubNamelist(idx).visibility) > 0 Then vis = SubNamelist(idx).visibility
        If Len(SubNamelist(idx).kind) > 0 Then
            kindStr = SubNamelist(idx).kind          '"Function" / "Property Get" / ...
            If InStr(kindStr, "Property") > 0 Then NVProcEndWord = "Property" Else NVProcEndWord = kindStr
        End If
    End If
    'A public class method's true kind (Function / Property Get / Sub) comes from its
    'typeinfo FuncDesc.  Apply it unless the Get/Let adjacency pass already tagged it
    'a Property (that pairing distinguishes the Let half, which the FuncDesc does not).
    Dim mkind As String
    If InStr(kindStr, "Property") = 0 And NativeTryMethodKind(addr, mkind) Then
        kindStr = mkind
        If InStr(kindStr, "Property") > 0 Then NVProcEndWord = "Property" Else NVProcEndWord = kindStr
    End If
    'Add the parameter list when the name carries no signature yet (event handlers
    'already get a typed one from getEventComplete).  Parameters are named generically
    'arg_<ebp offset> with the count from the proc's `ret N`.
    If InStr(nm, "(") = 0 Then
        Dim psig As String
        If NativeTryMethodSig(addr, psig) Then
            'Real parameter names from the class typeinfo (FuncDesc table).
            nm = nm & "(" & psig & ")"
            NativeBuildArgNameMap addr, psig          'use those names in the body too
        Else
            'Generic arg_<offset> list from the `ret N` count.
            nm = nm & "(" & NativeProcParams(kindStr, NativeProcHasMe(addr)) & ")"
        End If
    End If
    NativeProcHeader = vis & " " & kindStr & " " & nm & "   '" & Hex$(addr)
End Function

Private Function NativeProcParams(ByVal kindStr As String, ByVal hasMe As Boolean) As String
    'Reconstruct the (generic) parameter list from `ret N`.  N is the bytes the
    'callee pops = all stack arguments.  A class/form method receives a hidden
    'Me/this at ebp+8 (so user params start at ebp+0xC); a .bas module procedure
    'has no Me (user params start at ebp+8).  A Function (named as such) also
    'reserves a hidden return-value slot.  Names are arg_<ebp offset> (the offset
    'convention the commercial decompiler uses).
    Dim slots As Long, nParams As Long, i As Long, off As Long, base As Long, s As String
    If NVRetN < 4 Then Exit Function           '-1 (plain ret) or ret 0 -> no stack args
    slots = NVRetN \ 4
    If hasMe Then nParams = slots - 1 Else nParams = slots    'drop the hidden Me/this slot
    If InStr(kindStr, "Function") > 0 Then nParams = nParams - 1   'drop the return-value slot
    If nParams < 0 Then nParams = 0
    base = IIf(hasMe, &HC, &H8)                 'first user parameter's ebp offset
    For i = 0 To nParams - 1
        off = base + i * 4
        If Len(s) > 0 Then s = s & ", "
        s = s & "arg_" & Hex$(off)
    Next
    NativeProcParams = s
End Function

Private Function NativeProcHasMe(ByVal addr As Long) As Boolean
    'True when a procedure receives a hidden Me/this pointer at ebp+8 - i.e. it is a
    'class/form method, not a .bas module procedure.  Decided from the owning
    'object's type (module = ObjectType bit 0x2 clear).  Defaults to True (assume a
    'method) when the owner can't be resolved, preserving prior behaviour.
    Dim owner As String, i As Long
    owner = NativeFormOf(addr)
    If Len(owner) = 0 Then NativeProcHasMe = True: Exit Function
    On Error Resume Next
    For i = 0 To UBound(gObjectNameArray)
        If gObjectNameArray(i) = owner Then
            NativeProcHasMe = ((gObject(i).ObjectType And &H2) <> 0)   'non-module has Me
            Exit Function
        End If
    Next
    NativeProcHasMe = True
End Function

Private Function NativeProcName(ByVal addr As Long) As String
    Dim nm As String, p As Long, idx As Long
    idx = NativeProcMatchIdx(addr)
    If idx >= 0 Then nm = SubNamelist(idx).strName
    If Len(nm) = 0 Then nm = "proc_" & Hex$(addr)
    p = InStr(nm, ".")
    If p > 0 Then nm = Mid$(nm, p + 1)
    NativeProcName = nm
End Function

Private Function NativeFormOf(ByVal addr As Long) As String
    Dim i As Long, nm As String, p As Long, d As Long
    On Error Resume Next
    For i = 0 To UBound(gNativeProcArray)
        d = gNativeProcArray(i).offset - addr           'list entry is at/after the snapped entry
        If d >= 0 And d <= 24 Then nm = gNativeProcArray(i).sName: Exit For
    Next
    If Len(nm) = 0 Then
        For i = 0 To UBound(SubNamelist)
            d = addr - SubNamelist(i).offset
            If d >= 0 And d <= 24 Then nm = SubNamelist(i).strName: Exit For
        Next
    End If
    p = InStr(nm, ".")
    If p > 0 Then NativeFormOf = Left$(nm, p - 1) Else NativeFormOf = "Form1"
End Function

Private Sub NativeAddUnique(c As Collection, ByVal v As Long)
    On Error Resume Next
    c.Add v, "k" & v
End Sub

Private Function NativeHas(c As Collection, ByVal v As Long) As Boolean
    Dim x As Variant
    On Error Resume Next
    For Each x In c
        If x = v Then NativeHas = True: Exit Function
    Next
End Function
