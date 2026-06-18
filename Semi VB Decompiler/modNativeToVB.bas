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
Private NVIsClass As Boolean      'owning object is a CLASS (not a form/module) - gates private-field-type harvesting
Private NVBase As Long            'solved control-block base for this form (-1 = unknown)
Private NVLastControl As String   'most recently accessed control name
Private NVLastGuid As String      'its GUID (for property name resolution)
Private NVFpu() As String         'FPU expression stack
Private NVFpuTop As Long
Private NVLocal As Collection      'local stack slot (disp) -> expression
Private NVLocalGuid As Collection  'local stack slot (disp) -> control GUID, when the slot holds a control object
Private NVLastLea As Long          'displacement of the most recent LEA (GET out-param)
Private NVLastLeaSet As Boolean
Private NVLastLeaField As Boolean  'the most recent LEA addressed a Me-FIELD ([Me+off]),
                                   'so a move/copy into it is a field store field_<off> = src
Private NVPendingArg As String     'value fstp'd into the outgoing argument area
Private NVLastImm As String        'most recent pushed immediate (decoded)
Private NVRegImport(7) As String   'register -> runtime helper cached into it
Private NVPushImm() As String      'recent pushed immediates (call argument list)
Private NVPushDisp() As Long       'for each pushed arg, the by-ref local displacement it addresses (0 = not by-ref)
Private NVPushTop As Long
Private NVLastPushDisp As Long     'set by NativePushOperand: by-ref local disp of the value just decoded (0 = none)
Private NVVSlot As Collection      'variant stack slot (disp) -> last value stored (VT tag / string / expr)
Private NVLastVarData As String    'last value written to a Variant DATA field (offset +8) - the RHS for a following late-bound property put whose value would otherwise be <value>
Private NVLastVarBase As String    'the base local (e.g. var_C) of that Variant temp - its field-build stores are suppressed once a late put consumes the value
Private NVSuppressVarBuild As Collection 'set of Variant-temp base names whose numeric-field-store lines to strip in finalisation
Private NVVarArgList() As String   'ordered Variant data-field (offset +8) values built since the last consumer - the argument list for a following control-method call (obj.Method a, b)
Private NVVarArgBase() As String   'parallel: the temp base of each, to strip the build statements
Private NVVarArgN As Long
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
Private NVForStart As Collection   '"F"&headerVA -> start value (detect-time, for register-counter loops where the header isn't a plain cmp)
Private NVForLimit As Collection   '"F"&headerVA -> limit value (detect-time, ditto)
Private NVForCnt As Long           'per-proc loop-variable name allocation index
'--- Floating-point comparison idiom (fcom/fnstsw/test ah,mask/jcc -> boolean) ---
Private NVFpCmp As Collection      '"P"&testVA -> "<relOp>|<regIdx>" - set NVReg(reg) to the relational
Private NVFpSkip As Collection     '"P"&va -> "1" - suppress this scaffolding instruction
Private NVStrCmpReg As Collection  '"P"&strcmpCallVA -> regIdx - the StrCmp result boolean-materializes into this register
Private NVStrCmpDirect As Collection '"P"&strcmpCallVA -> "1" - the result is tested directly (`call __vbaStrCmp; test eax,eax; jcc`), no materialization/store; the handler hands its operands to the test
Private NVStrCmpP1 As String        'direct-test StrCmp idiom (call __vbaStrCmp; test eax,eax; jcc): the left operand,
Private NVStrCmpP2 As String        'and the right operand, handed from the StrCmp call to the following `test eax,eax`
Private NVStrCmpPending As Boolean  'so the Jcc renders `If p1 <op> p2` (test strcmp,strcmp has the same Jcc polarity as cmp p1,p2)
Private NVCounterSlot As Collection '"C"&disp -> "1" - a stack slot that is a loop induction variable (render by name, not a stale value)
Private NVWhileCond As Collection  '"W"&exitJccVA -> "1" - emit `Do While <cond>` here (top-tested loop header)
Private NVWhileLoop As Collection  '"W"&backedgeVA -> "1" - emit `Loop` here (the back-edge of a Do While)
'Variant For loops: a `For v = a To b` over a Variant counter compiles to
'__vbaVarForInit (header) ... __vbaVarForNext (back-edge).  Detected on top of the
'Do While structure, then rendered For/Next instead of Do While/Loop and the two
'helper calls suppressed.  NVVarForInitLink: "V"&forInitCallVA -> "jccVA|backedgeVA"
'(set in the pre-pass); the call handler reads the args and fills NVVarForFor
'("W"&jccVA -> "ctr|start|limit", render the For header) + NVVarForNext
'("W"&backedgeVA -> ctr, render Next).  NVVarForSuppress drops the ForNext call.
Private NVVarForInitLink As Collection
Private NVVarForFor As Collection
Private NVVarForNext As Collection
Private NVVarForSuppress As Collection
Private NVElemIdx As Collection    '"E"&accessVA -> recovered logical array index expr (e.g. arg_8) for a SAFEARRAY element access at that VA; injected by the SIB read/store renderers
Private NVCurVa As Long            'VA of the instruction currently being rendered (so the SIB renderers can look up NVElemIdx)
Private NVLateDispid As Collection '"L"&callVA -> comma list of candidate DISPIDs pushed to a __vbaLateIdCall (resolve to the OCX member name at render)
'Select-Case-on-Integer-parameter compares: a ByRef Integer param (e.g. KeyCode) is
'loaded ONCE into a callee-saved 16-bit register and tested per case via
'`mov ecx,const; call __vbaI2I4; cmp di,ax` or `cmp di,imm16`.  Our generic decoder
'bails on these 16-bit register compares; this map records the resolved operands
'("<paramTok>|<const>") per cmp VA so the Jcc renders `If KeyCode = 97` not `<cond>`.
Private NVKeyCmp As Collection     '"K"&cmpVA -> "<paramTok>|<const>"
Private NVAbsGlobalCmp As Collection '"T"&testVA -> global token: a `mov eax,[abs](0xA1); test eax,eax` condition (the short-form global load isn't register-tracked) -> `If global_X <op> 0`
'Branchless select-of-two-constants (`If cond Then x=c1 Else x=c2` compiled as
'xor/cmp/setcc/dec/and mask/add base) -> reconstruct IIf(cond, base, base+mask).
Private NVSelConst As Collection     '"S"&setccVA -> "<base>|<base+mask>" (true / false value)
Private NVSelConstSkip As Collection '"S"&va -> "1" - the dec/and/add tail to suppress
'--- Control-array element accessor reconstruction (lblSkillName(i).Caption) ---
'A control array's element accessor is `call [arrayVt + 0x40]` (the array object's
'Item property).  A deterministic pre-pass (NativeDetectControlArrays) recovers, per
'such call VA, the array control + index + the element retbuf local, so the call
'renders `Set var_X = Form.ctrl(idx)` and the following property put through var_X
'resolves (.Caption / .ToolTipText) via the normal control-property path.
Private NVCtlArrElem As Collection   '"K"&callVA -> the element expression (e.g. "frmMain.lblSkillName(var_20)")
Private NVCtlArrGuid As Collection   '"K"&callVA -> the array control's GUID (for element property resolution)
Private NVCtlArrRetbuf As Collection '"K"&callVA -> the element retbuf local displacement (negative), as a string
'A control object's vtable cached in a local then reloaded (`mov [ebp-X],vt; ...; mov
'reg,[ebp-X]; call [reg+off]`) - thread the vtable's control identity through the slot
'so the deferred property put still resolves (the first Caption put caches its vtable).
Private NVLocalVtName As Collection  '"L"&disp -> control NAME whose vtable that slot holds
Private NVLocalVtGuid As Collection  '"L"&disp -> control GUID whose vtable that slot holds
Private NVSuppressObjSet As Collection '"O"&objSetVA -> "1" - a __vbaObjSet whose only purpose is to pass the control into the very next late call; drop the redundant `Set temp = control`
'Byte-field increment/decrement-with-clamp idiom (VB6 `field = field +/- 1 : If field
'<op> N Then field = M`, where field is a Byte member coerced via __vbaUI1I2).  The whole
'sequence (movzx/sub/jo/call/cmp/store/jcc/mov-imm/call/store) is recognised by signature
'and reconstructed; the raw instructions are suppressed.  Beats the commercial decompiler,
'which drops the unconditional decrement store and mangles the condition (`+ 1+1 > 9`).
Private NVByteClamp As Collection      '"B"&anchorVA -> "opSign|cmpOp|cmpVal|bodyVal|fieldOff"
Private NVByteClampSkip As Collection  '"B"&va -> "1" - an idiom instruction to suppress
Private NVArgTok() As String       'per-proc: generic tokens (arg_<offset>) to replace with...
Private NVArgNm() As String        '...their recovered parameter names, at proc finalisation
Private NVArgN As Long             'count of recovered parameter-name substitutions this proc
Private NVLastCmp As String        'hint expression for the next If condition
Private NVStrLits As Collection    'pending string literals (e.g. MsgBox arguments)
Private NVSkipLabels As Collection 'branch targets that belong to dropped error-check guards
Private NVReg(7) As String         'symbolic value currently held in each GP register (eax..edi)
Private NVR16Val(7) As String      'compare-only shadow: a 16-bit memory word just loaded into a register (mov ax,word[mem]).  Consumed ONLY by `cmp ax,imm16` so an Integer-field / array-element compare resolves to its operand instead of <cond>; deliberately NOT stored in NVReg, whose 16-bit value is cleared (the low-word partial must not leak into a push/store/arithmetic).  Cleared for an instruction's dest reg at the top of NativeProcessInst and for caller-saved regs on a call.
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
Private NVRegFieldCls(7) As String 'class of the As-New object field a register's ADDRESS points to (from `lea reg,[Me+off]`); the next `mov obj,[reg]` deref tags obj as that class
Private NVRegFieldRecv(7) As String 'receiver expr (field_<off>) for that field, carried to the dereffed object so its method calls render field_<off>.Method
Private NVObjClass As Collection   'key "G"&globalVA -> user class name of the object instance stored at that global (typed at the __vbaNew auto-instantiation)
Private NVLocalObjType As Collection 'key "D"&localDisp -> user class created INTO that local by __vbaNew2(ObjInfo, &local); so var_X.Member resolves via the class vtable map
Private NVNewEmitted As Collection 'per-proc: local disps that already emitted `Set var_X = New <class>` - suppress the repeated auto-instantiation (As New) guards before each use
Private NVPropDir As Collection      'key "P"&callVA -> "get"/"put": data-flow direction of a property accessor call (its by-ref local read AFTER = get, else put), since VB6 FuncDesc flags Get and Let identically
Private NVRecentPush(7) As Long    'ring of recent `push imm32/imm8` raw values (to recover __vbaNew's Object Info + @global args)
Private NVRecentTop As Long
Private NVLoopHdr As Collection    'addresses that are loop headers (back-edge targets)
Private NVCallHandled As Boolean   'set by NativeRuntimeCall: True when the call was recognised
Private NVErrHandler As Long       'address of this procedure's On Error handler block (0 = none)
Private NVProcEndWord As String    'closing keyword for this proc: "Sub" / "Function" / "Property"
Private NVAccumRet As Boolean      'the proc returns a simple value in the accumulator (ax/eax) - a module Function whose kind/type is stripped (recovered from the epilogue return-load)
Private NVRetbuf As Boolean        'the proc returns a Variant/String/UDT via a hidden retbuf (first param, ebp+8) - a module Function (from gRetbufFunc); the header drops that param and renders `As Variant`
Private NVAccumRetType As String   'recovered return type from that load: "Integer" (word/ax) or "Long" (dword/eax)
Private NVAccumRetSlot As Long      'the return slot's ebp displacement (negative; 0 = none) - so a constant store `mov [ebp-slot],imm` (0xC7) to it renders `FuncName = imm`
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
    'Build the retbuf-returning-Function map once (a module Function returning a
    'Variant/String/UDT passes a hidden retbuf as its first param; callers render
    '`<dest> = proc(args)` and the proc itself is a Function).
    If gRetbufFunc Is Nothing Then NativeScanRetbufFuncs

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
    NVIsClass = NativeOwnerIsClass(NVForm)
    NVProcEndWord = "Sub"
    NVAccumRet = False: NVAccumRetType = "": NVAccumRetSlot = 0: NVRetbuf = False
    NVLastControl = "": NVLastGuid = "": NVLastImm = "": NVPendingArg = ""
    NVLastLea = 0: NVLastLeaSet = False: NVLastLeaField = False: NVLastCmp = ""
    NVCmpSet = False: NVCmpL = "": NVCmpR = "": NVCmpIsTest = False: NVCmpIsBool = False: NVFpuChk = False
    NVStrCmpPending = False: NVStrCmpP1 = "": NVStrCmpP2 = ""
    NVPendingCall = "": NVErrObjPending = False
    ReDim NVFpu(31): NVFpuTop = 0
    ReDim NVPushImm(31): ReDim NVPushDisp(31): NVPushTop = 0: NVLastPushDisp = 0
    ReDim NVIfTarget(31): NVIfTop = 0: NVIndent = 0
    Dim r As Long
    For r = 0 To 7: NVReg(r) = "": NVR16Val(r) = "": NVRegIsAddr(r) = False: NVRegAddr(r) = "": NVRegAddrDisp(r) = 0: NVRegIsMe(r) = False: NVRegIsFormVt(r) = False: NVRegObjType(r) = "": NVRegObjVt(r) = "": NVRegObjGuid(r) = "": NVRegObjVtGuid(r) = "": NVRegObjInst(r) = "": NVRegFieldCls(r) = "": NVRegFieldRecv(r) = "": Next
    Set NVObjClass = New Collection
    For r = 0 To 7: NVRecentPush(r) = 0: Next
    NVRecentTop = 0
    Set NVLocal = New Collection
    Set NVLocalGuid = New Collection
    Set NVStrLits = New Collection
    Set NVVSlot = New Collection
    NVLastVarData = "": NVLastVarBase = ""
    ReDim NVVarArgList(15): ReDim NVVarArgBase(15): NVVarArgN = 0
    Set NVSuppressVarBuild = New Collection
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
    Set NVForStart = New Collection
    Set NVForLimit = New Collection
    NVForCnt = 0
    Set NVFpCmp = New Collection
    Set NVFpSkip = New Collection
    Set NVStrCmpReg = New Collection
    Set NVStrCmpDirect = New Collection
    Set NVCounterSlot = New Collection
    Set NVWhileCond = New Collection
    Set NVWhileLoop = New Collection
    Set NVVarForInitLink = New Collection
    Set NVVarForFor = New Collection
    Set NVVarForNext = New Collection
    Set NVVarForSuppress = New Collection
    Set NVElemIdx = New Collection
    Set NVLateDispid = New Collection
    Set NVSuppressObjSet = New Collection
    Set NVCtlArrElem = New Collection
    Set NVCtlArrGuid = New Collection
    Set NVCtlArrRetbuf = New Collection
    Set NVKeyCmp = New Collection
    Set NVAbsGlobalCmp = New Collection
    Set NVSelConst = New Collection
    Set NVSelConstSkip = New Collection
    Set NVLocalVtName = New Collection
    Set NVLocalVtGuid = New Collection
    Set NVByteClamp = New Collection
    Set NVByteClampSkip = New Collection
    Set NVLocalObjType = New Collection
    Set NVNewEmitted = New Collection
    Set NVPropDir = New Collection
    NVCurVa = 0
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

    'A FRAMELESS function (no `push ebp` entry; args at [esp+N], result in eax) is
    'invisible to the framed decoder - recover its simple arithmetic/accessor body
    'directly and emit it as a Function.  Bails to the normal (empty) path on any
    'shape it does not fully understand.
    If Not (b(0) = &H55 And b(1) = &H8B And b(2) = &HEC) Then
        Dim flIsFunc As Boolean, flExpr As String
        flExpr = NativeFramelessBody(col, flIsFunc)
        If Len(flExpr) > 0 Then
            Dim flName As String, flParams As String, flpp As Long
            flName = NativeProcName(addr)
            flpp = InStr(flName, "(")
            If flpp > 0 Then flName = Trim$(Left$(flName, flpp - 1))
            flParams = NativeProcParams("Sub", NVHasMe)        'no hidden-retslot drop (returns via eax)
            output = "Private Function " & flName & "(" & flParams & ")   '" & Hex$(addr) & vbCrLf
            output = output & Space$(4) & flName & " = " & flExpr & vbCrLf
            output = output & "End Function" & vbCrLf
            DecompileNativeProcToVB = NativeInsertLocalDims(NativeSubstituteArgNames(output))
            Exit Function
        End If
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

    'Recover the logical index of each SAFEARRAY element access (Player(i).field) by
    'back-tracing the byte-offset register through the addressing chain, so the access
    'renders global_X(12)(i)(field) instead of dropping the index.
    NativeDetectElemIndices col

    'Decode late-bound dispatch calls (__vbaLateIdCall): collect each call's DISPID
    'so it can resolve to obj.Member via the control's OCX typelib at render time.
    NativeDetectLateCalls col

    'Reconstruct control-array element accesses (lblSkillName(i)): for each element
    'accessor `call [arrayVt + 0x40]` whose receiver back-traces to an is-array form
    'control, recover the index + the element retbuf local so it renders as
    'Set var_X = Form.ctrl(idx) and the following .Caption/.ToolTipText puts resolve.
    NativeDetectControlArrays col

    'Resolve Select-Case-on-Integer-parameter compares (e.g. Form_KeyDown's KeyCode
    'cases): bind the per-case `cmp di,<const>` to `<param> = <const>` by VA so the
    'conditions render `If KeyCode = 97` instead of `<cond>`.  Strictly scoped (only
    'procs with the `mov di,word[param]` anchor; only the 16-bit case compares).
    NativeDetectKeyCompares col
    NativeDetectAbsGlobalTests col

    'Recognise the branchless select-of-two-constants idiom (setcc/dec/and/add) so a
    'two-constant If/Else (modMap_Direction = 4 / = 1) reconstructs as IIf(cond, c1, c2)
    'rather than the misleading bare relational.
    NativeDetectSelectConst col

    'Reconstruct the Byte-field +/-1-with-clamp idiom (frmCreate body-part Up/Down
    'buttons) into `field = field +/- 1 : If field <op> N Then field = M / End If`,
    'suppressing the raw movzx/__vbaUI1I2/store/jcc scaffolding.
    NativeDetectByteClamp col

    'Determine each vtable call's data-flow direction (its by-ref local read AFTER the
    'call = a value-out GET, else a value-in PUT) so a class-instance property access
    'renders `var = obj.Prop` vs `obj.Prop = value` - the FuncDesc invoke-kind can't tell.
    NativeDetectPropDir col

    'Detect an unclassified module Function: a simple value returned in the accumulator
    '(`mov ax/eax,[ebp-retSlot]` right before the SEH restore).  Marks the proc a
    'Function, recovers its return TYPE (Integer/Long), and renames the return slot to
    'the proc name so `FuncName = value` reads as VB's implicit return.  MUST run before
    'NativeProcHeader (which reads NVAccumRet to emit `Function ... As <type>`).
    NativeDetectAccumReturn b, addr
    'A module Function returning a Variant/String/UDT via a hidden retbuf (first param):
    'mark it so NativeProcHeader emits `Function ... As Variant` and drops the retbuf
    'param.  Mutually exclusive with the accumulator return (different epilogue).
    If Not NVAccumRet Then NVRetbuf = (Len(NativeColGet(gRetbufFunc, "V" & addr)) > 0)

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

        '--- Byte-field +/-1-with-clamp idiom ---
        'At the anchor (the movzx), emit the reconstructed statements; every idiom
        'instruction (including the anchor) is then suppressed.
        Dim bcRec As String
        bcRec = NativeColGet(NVByteClamp, "B" & inst.va)
        If Len(bcRec) > 0 Then
            Dim bcP() As String, bcFld As String, bcInd As String
            bcP = Split(bcRec, "|")
            bcFld = NativeFieldName(CLng(bcP(4)))
            bcInd = NativeIndentStr()
            output = output & bcInd & bcFld & " = (" & bcFld & " " & bcP(0) & " 1)" & vbCrLf
            output = output & bcInd & "If " & bcFld & " " & bcP(1) & " " & bcP(2) & " Then" & vbCrLf
            NVIndent = NVIndent + 1
            output = output & NativeIndentStr() & bcFld & " = " & bcP(3) & vbCrLf
            NVIndent = NVIndent - 1
            output = output & bcInd & "End If" & vbCrLf
        End If
        If Len(NativeColGet(NVByteClampSkip, "B" & inst.va)) > 0 Then GoTo nextInst

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

        'Suppress the dec/and/add tail of a branchless select-of-two-constants idiom -
        'its value is reconstructed as IIf(cond, base, base+mask) at the setcc, so these
        'must not fold onto (and corrupt) the bound IIf expression.
        If Len(NativeColGet(NVSelConstSkip, "S" & inst.va)) > 0 Then GoTo nextInst

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
            Dim fStoreL As String
            fCtr = CLng(fCtrS)
            fStoreL = NativeColGet(NVForLimit, "F" & inst.va)
            If Len(fStoreL) > 0 Then
                'Register-counter loop: the back-edge targets a limit-register reload
                'so the header isn't the plain cmp the render-time decode can read.
                'start/limit were captured at detect time.
                fStart = NativeColGet(NVForStart, "F" & inst.va): fLimit = fStoreL
            Else
                NativeDecodeCompare inst, "CMP"          'fills NVCmpL (start=counter init) / NVCmpR (limit)
                fStart = NVCmpL: fLimit = NVCmpR
                NVCmpSet = False: NVCmpL = "": NVCmpR = ""
                If Len(fStart) = 0 Then fStart = NativeRegVal(fCtr)
                If Len(fStart) = 0 Then fStart = "1"
                If Len(fLimit) = 0 Then fLimit = "?"
            End If
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
        Dim stmtTxt As String
        stmtTxt = NativeProcessInst(inst)
        output = output & stmtTxt
        'Reset the pending Variant-method-arg list when a NON-build statement is
        'emitted (a Variant field-build line var_X(<digit>) = .. keeps accumulating;
        'anything else - a call that consumed them, or unrelated code - clears them so
        'a stale sequence never bleeds into the next control-method call).
        If Len(Trim$(stmtTxt)) > 0 And Not NativeIsVarBuildLine(stmtTxt) Then NativeResetVarArgs
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
            NVR16Val(0) = "": NVR16Val(1) = "": NVR16Val(2) = ""   'caller-saved 16-bit word shadow invalid after a call
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
    output = NativeStripOrphanLabels(NativeSubstituteConstants(NativeSubstituteArgNames(NativeStripVarBuild(output))))
    output = NativeMergeElseIf(output)
    DecompileNativeProcToVB = NativeInsertLocalDims(output)
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

Private Function NativeMergeElseIf(ByVal src As String) As String
    'Merge a chain of sibling If-blocks into If / ElseIf.  When an `If <c> Then ... End If`
    'block's body ends in an UNCONDITIONAL transfer (GoTo / Exit Sub|Function|Property|Do|
    'For / End), the IMMEDIATELY-following sibling `If` at the same indent is reachable
    'ONLY when <c> was false - i.e. it is logically the Else.  Rewriting `End If` + the
    'next `If <c2> Then` to `ElseIf <c2> Then` is therefore SEMANTICS-PRESERVING (the
    'control flow is identical), and If/End-If counts stay balanced (both drop by 1).
    'STRICTLY gated so it never changes behaviour:
    '  - the End If trims to exactly "End If",
    '  - the very next line (no label / blank / statement between) is a BLOCK `If <c> Then`
    '    (ends in "Then") at the SAME indent,
    '  - the line before the End If is at indent+4 and is an unconditional transfer
    '    (so the block cannot fall through to the next If).
    'Anything off leaves the blocks exactly as they were.
    On Error GoTo done
    Dim lines() As String, i As Long, n As Long, ind As Long
    lines = Split(src, vbCrLf)
    n = UBound(lines)
    If n < 2 Then Exit Function
    Dim skip() As Boolean
    ReDim skip(n)
    For i = 1 To n - 1
        If Trim$(lines(i)) = "End If" Then
            ind = NativeLineIndent(lines(i))
            Dim nxtT As String
            nxtT = Trim$(lines(i + 1))
            If NativeLineIndent(lines(i + 1)) = ind _
               And Left$(nxtT, 3) = "If " And Right$(nxtT, 5) = " Then" _
               And NativeLineIndent(lines(i - 1)) = ind + 4 _
               And NativeIsUncondTransfer(Trim$(lines(i - 1))) Then
                'Drop this End If; rewrite the following If as ElseIf (in place, so a
                'longer chain keeps collapsing on later iterations of this same pass).
                lines(i + 1) = Space$(ind) & "Else" & nxtT
                skip(i) = True
            End If
        End If
    Next i
    'Rejoin, dropping the merged-away End If lines (Join preserves the exact format).
    Dim res() As String, rc As Long
    ReDim res(n): rc = 0
    For i = 0 To n
        If Not skip(i) Then res(rc) = lines(i): rc = rc + 1
    Next i
    ReDim Preserve res(rc - 1)
    NativeMergeElseIf = Join(res, vbCrLf)
    Exit Function
done:
    NativeMergeElseIf = src
End Function

Private Function NativeLineIndent(ByVal s As String) As Long
    NativeLineIndent = Len(s) - Len(LTrim$(s))
End Function

Private Function NativeIsUncondTransfer(ByVal t As String) As Boolean
    'An unconditional control transfer that prevents falling through to the next line.
    NativeIsUncondTransfer = (Left$(t, 5) = "GoTo ") _
        Or t = "Exit Sub" Or t = "Exit Function" Or t = "Exit Property" _
        Or t = "Exit Do" Or t = "Exit For" Or t = "End"
End Function

Private Function NativeInsertLocalDims(ByVal body As String) As String
    'Task A: declare local stack slots (var_XX) with a USAGE-inferred type, emitted
    'as a `Dim` block after the proc header.  Native compilation strips local names
    'AND types, but the type is often inferable from how the slot is assigned - which
    'beats the commercial decompiler's blanket `As Variant`:
    '  String   - assigned a string function (Left$/Trim$/UCase$...), a `&` concat,
    '             a string literal, or CStr.
    '  Long/Integer - Len/UBound/CLng/CInt/InStr/Asc or an arithmetic expression.
    '  As <class>   - `Set v = New clsX` (the class is already resolved in the body).
    '  Object       - other `Set v = ...`.
    '  Variant      - no signal, or conflicting signals (also the commercial default).
    'A post-pass over the FINISHED body text (after arg/const substitution, so the
    'return-slot local is already the proc name and parameters are arg_XX/real names -
    'neither is a var_XX, so neither gets a bogus Dim).
    On Error GoTo done
    Dim lines() As String, i As Long, lt As String, work As String
    Dim seen As Collection, typ As Collection
    lines = Split(body, vbCrLf)
    If UBound(lines) < 0 Then GoTo done
    'Only operate on a real proc body (header line is Private/Public Sub/Function/...).
    If Not NativeIsProcHeaderLine(lines(0)) Then GoTo done
    Set seen = New Collection: Set typ = New Collection
    For i = 1 To UBound(lines)
        lt = Trim$(lines(i))
        If Len(lt) = 0 Then GoTo nextLine
        NativeScanVarTokens lt, seen
        'For <var> = ... -> a Long loop counter.
        If Left$(lt, 4) = "For " Then
            Dim fv As String
            fv = NativeFirstVarTokAt(Mid$(lt, 5))
            If Len(fv) > 0 Then NativeMergeType typ, fv, "Long"
            GoTo nextLine
        End If
        'An assignment whose LHS is a lone var_XX: infer the slot's type from the RHS.
        Dim isSet As Boolean, eqp As Long, lhs As String, rhs As String, t As String
        work = lt: isSet = False
        If Left$(work, 4) = "Set " Then isSet = True: work = Mid$(work, 5)
        eqp = InStr(work, " = ")
        If eqp > 0 Then
            lhs = Trim$(Left$(work, eqp - 1))
            rhs = Trim$(Mid$(work, eqp + 3))
            If NativeIsLoneVarTok(lhs) Then
                If isSet Then t = NativeInferSetType(rhs) Else t = NativeInferValType(rhs)
                If Len(t) > 0 Then NativeMergeType typ, lhs, t
            End If
        End If
nextLine:
    Next
    'A lone var_X passed to a record helper (RecAssign/RecDestruct/Rec*ToUni) is a
    'whole UDT record of that descriptor's type - authoritative over the usage guess.
    NativeTypeUDTLocals lines, typ, seen
    If seen.Count = 0 Then GoTo done
    'Order the locals by their ebp offset (the hex in var_XX) ascending.
    Dim names() As String, offs() As Long, c As Long, v As Variant, j As Long
    ReDim names(seen.Count - 1): ReDim offs(seen.Count - 1)
    c = 0
    For Each v In seen
        names(c) = CStr(v): offs(c) = NativeHexVal(Mid$(CStr(v), 5)): c = c + 1
    Next
    For i = 0 To c - 2
        For j = 0 To c - 2 - i
            If offs(j) > offs(j + 1) Then
                Dim ts As Long, tn As String
                ts = offs(j): offs(j) = offs(j + 1): offs(j + 1) = ts
                tn = names(j): names(j) = names(j + 1): names(j + 1) = tn
            End If
        Next
    Next
    Dim block As String, ty As String
    For i = 0 To c - 1
        ty = NativeColGet(typ, names(i))
        If Len(ty) = 0 Then ty = "Variant"
        block = block & "    Dim " & names(i) & " As " & ty & vbCrLf
    Next
    'Splice the block in after the header line.
    Dim res As String
    res = lines(0) & vbCrLf & block
    For i = 1 To UBound(lines)
        res = res & lines(i)
        If i < UBound(lines) Then res = res & vbCrLf
    Next
    NativeInsertLocalDims = res
    Exit Function
done:
    NativeInsertLocalDims = body
End Function

Private Sub NativeTypeUDTLocals(ByRef lines() As String, typ As Collection, seen As Collection)
    'Force the type of any lone var_X local passed to a record (__vbaRec*) helper to
    'the UDT recovered from that call's descriptor argument.  A record helper only
    'takes whole-record pointers, so such a local is definitively that UDT - this is
    'authoritative (the usage inference otherwise mis-types it As String/Variant, and
    'an element-address operand like (i - 1) is not a lone var so it never matches).
    Dim i As Long, ln As String, p As Long, k As Long
    Dim kws As Variant, kw As Variant
    kws = Array("RecAssign(", "RecDestruct(", "RecAnsiToUni(", "RecUniToAnsi(", "RecDestructAnsi(")
    On Error Resume Next
    For i = 0 To UBound(lines)
        ln = lines(i)
        For Each kw In kws
            p = InStr(ln, CStr(kw))
            Do While p > 0
                Dim argStr As String
                argStr = NativeParenArgs(ln, p + Len(CStr(kw)) - 1)   'text inside the call parens
                If Len(argStr) > 0 Then NativeTypeUDTArgList argStr, typ, seen
                p = InStr(p + 1, ln, CStr(kw))
            Loop
        Next
    Next
End Sub

Private Sub NativeTypeUDTArgList(ByVal argStr As String, typ As Collection, seen As Collection)
    Dim parts() As String, k As Long, a As String, descVa As Long, udtName As String, udtKey As String
    parts = NativeSplitTopComma(argStr)
    If UBound(parts) < 1 Then Exit Sub
    a = Trim$(parts(0))
    If Not IsNumeric(a) Then Exit Sub                       'first arg = descriptor address (decimal)
    descVa = CLng(a)
    udtKey = "H" & Right$("00000000" & Hex$(descVa), 8)
    If gUDTDesc Is Nothing Then Exit Sub
    If Len(NativeColGet(gUDTDesc, udtKey)) = 0 Then Exit Sub  'only when the Type was actually emitted
    udtName = "UDT_" & Right$("00000000" & Hex$(descVa), 8)
    For k = 1 To UBound(parts)
        a = Trim$(parts(k))
        If NativeIsLoneVarTok(a) Then
            NativeColPut typ, a, udtName
            NativeColPut seen, a, a
        End If
    Next
End Sub

Private Function NativeParenArgs(ByVal s As String, ByVal openPos As Long) As String
    'Return the text between the '(' at openPos and its matching ')'.
    Dim i As Long, depth As Long, n As Long, ch As String
    n = Len(s)
    If openPos < 1 Or openPos > n Then Exit Function
    If Mid$(s, openPos, 1) <> "(" Then Exit Function
    depth = 0
    For i = openPos To n
        ch = Mid$(s, i, 1)
        If ch = "(" Then
            depth = depth + 1
        ElseIf ch = ")" Then
            depth = depth - 1
            If depth = 0 Then NativeParenArgs = Mid$(s, openPos + 1, i - openPos - 1): Exit Function
        End If
    Next
End Function

Private Function NativeSplitTopComma(ByVal s As String) As String()
    'Split s on top-level ", " (depth-0 commas), so a nested arg like (i - 1) stays whole.
    Dim res() As String, n As Long, i As Long, depth As Long, ch As String, cur As String, ln As Long
    ReDim res(64): n = 0: cur = "": depth = 0: ln = Len(s)
    For i = 1 To ln
        ch = Mid$(s, i, 1)
        If ch = "(" Then
            depth = depth + 1: cur = cur & ch
        ElseIf ch = ")" Then
            depth = depth - 1: cur = cur & ch
        ElseIf ch = "," And depth = 0 Then
            res(n) = Trim$(cur): n = n + 1: cur = ""
        Else
            cur = cur & ch
        End If
    Next
    res(n) = Trim$(cur): n = n + 1
    ReDim Preserve res(n - 1)
    NativeSplitTopComma = res
End Function

Private Function NativeIsProcHeaderLine(ByVal ln As String) As Boolean
    Dim s As String
    s = Trim$(ln)
    If Left$(s, 8) <> "Private " And Left$(s, 7) <> "Public " And Left$(s, 7) <> "Friend " Then Exit Function
    NativeIsProcHeaderLine = (InStr(s, " Sub ") > 0 Or InStr(s, " Function ") > 0 Or InStr(s, " Property ") > 0)
End Function

Private Sub NativeScanVarTokens(ByVal line As String, seen As Collection)
    'Record every whole-identifier var_<hex> token in line into seen (deduped).
    Dim q As Long, j As Long, n As Long, before As Long, ch As Long, tok As String
    n = Len(line)
    q = InStr(1, line, "var_")
    Do While q > 0
        before = 0
        If q > 1 Then before = Asc(Mid$(line, q - 1, 1))
        If Not NativeIsIdentChar(before) Then
            j = q + 4
            Do While j <= n
                ch = Asc(Mid$(line, j, 1))
                If (ch >= 48 And ch <= 57) Or (ch >= 65 And ch <= 70) Or (ch >= 97 And ch <= 102) Then
                    j = j + 1
                Else
                    Exit Do
                End If
            Loop
            If j > q + 4 Then
                Dim okEnd As Boolean: okEnd = True
                If j <= n Then If NativeIsIdentChar(Asc(Mid$(line, j, 1))) Then okEnd = False
                If okEnd Then
                    tok = Mid$(line, q, j - q)
                    NativeColPut seen, tok, tok        'value = token (For Each yields values)
                End If
            End If
        End If
        q = InStr(q + 4, line, "var_")
    Loop
End Sub

Private Function NativeIsLoneVarTok(ByVal s As String) As Boolean
    'True when s is exactly one var_<hex> token (a simple slot LHS, not an element
    'store like global_X(12)(244) or an array index expression).
    Dim i As Long, ch As Long
    If Left$(s, 4) <> "var_" Or Len(s) <= 4 Then Exit Function
    For i = 5 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        If Not ((ch >= 48 And ch <= 57) Or (ch >= 65 And ch <= 70) Or (ch >= 97 And ch <= 102)) Then Exit Function
    Next
    NativeIsLoneVarTok = True
End Function

Private Function NativeFirstVarTokAt(ByVal s As String) As String
    Dim seen As Collection
    Set seen = New Collection
    NativeScanVarTokens s, seen
    'Return the first token by appearance: re-scan for position.
    Dim q As Long
    q = InStr(1, s, "var_")
    If q > 0 Then
        Dim j As Long, n As Long, ch As Long
        n = Len(s): j = q + 4
        Do While j <= n
            ch = Asc(Mid$(s, j, 1))
            If (ch >= 48 And ch <= 57) Or (ch >= 65 And ch <= 70) Or (ch >= 97 And ch <= 102) Then j = j + 1 Else Exit Do
        Loop
        If j > q + 4 Then NativeFirstVarTokAt = Mid$(s, q, j - q)
    End If
End Function

Private Function NativeInferValType(ByVal rhs As String) As String
    'Infer a local's type from a value (non-Set) assignment's right-hand side.
    'Strong signals only; returns "" when the RHS carries no reliable type.
    Dim s As String
    s = Trim$(rhs)
    If Len(s) = 0 Then Exit Function
    If InStr(s, " & ") > 0 Then NativeInferValType = "String": Exit Function     'concat
    If Left$(s, 1) = """" Then NativeInferValType = "String": Exit Function       'literal
    If NativeStartsWithFn(s, "CStr|Left$|Right$|Mid$|Trim$|LTrim$|RTrim$|UCase$|LCase$|Chr$|ChrW$|Space$|String$|Format$|Hex$|Oct$|Str$") Then NativeInferValType = "String": Exit Function
    If NativeStartsWithFn(s, "CInt|Asc") Then NativeInferValType = "Integer": Exit Function
    If NativeStartsWithFn(s, "CLng|Len|UBound|LBound|InStr") Then NativeInferValType = "Long": Exit Function
    If NativeStartsWithFn(s, "CBool") Then NativeInferValType = "Boolean": Exit Function
    If NativeStartsWithFn(s, "CDbl") Then NativeInferValType = "Double": Exit Function
    If NativeStartsWithFn(s, "CSng") Then NativeInferValType = "Single": Exit Function
    'A parenthesised arithmetic expression (has + - * but no concat/quote) -> numeric.
    If Left$(s, 1) = "(" And InStr(s, """") = 0 Then
        If InStr(s, " + ") > 0 Or InStr(s, " - ") > 0 Or InStr(s, " * ") > 0 Then NativeInferValType = "Long"
    End If
End Function

Private Function NativeInferSetType(ByVal rhs As String) As String
    'Infer an object reference's type from `Set v = <rhs>`.
    Dim s As String
    s = Trim$(rhs)
    If Left$(s, 4) = "New " Then
        Dim nm As String, i As Long, ch As Long
        nm = Trim$(Mid$(s, 5))
        'Class name = leading identifier run.
        Dim outn As String
        For i = 1 To Len(nm)
            ch = Asc(Mid$(nm, i, 1))
            If NativeIsIdentChar(ch) Then outn = outn & Mid$(nm, i, 1) Else Exit For
        Next
        If Len(outn) > 0 Then NativeInferSetType = outn Else NativeInferSetType = "Object"
    Else
        'A control reference (`Set v = frmMain.picView`) -> the intrinsic control
        'class (PictureBox / Label / ...), looked up by the control's name; falls
        'back to Object for a non-control object or an ambiguous/unknown name.
        Dim ct As String
        ct = NativeControlTypeOf(s)
        If Len(ct) > 0 Then NativeInferSetType = ct Else NativeInferSetType = "Object"
    End If
End Function

Private Function NativeControlTypeOf(ByVal rhs As String) As String
    'Map a control reference (the last `.`-segment is the control name) to its VB
    'intrinsic control class via gControlOffset (populated during form parsing).
    'Returns "" unless exactly one control class matches the name.
    On Error GoTo done
    Dim name As String, p As Long, i As Long, ch As Long
    name = rhs
    p = InStrRev(name, ".")
    If p > 0 Then name = Mid$(name, p + 1)
    'Trim to a leading identifier run (drop any trailing `(idx)` etc.).
    Dim clean As String
    For i = 1 To Len(name)
        ch = Asc(Mid$(name, i, 1))
        If NativeIsIdentChar(ch) Then clean = clean & Mid$(name, i, 1) Else Exit For
    Next
    If Len(clean) = 0 Then Exit Function
    Dim found As String, cn As String
    For i = 0 To UBound(gControlOffset)
        If gControlOffset(i).ControlName = clean Then
            cn = NativeControlClassName(gControlOffset(i).ControlType)
            If Len(cn) > 0 Then
                If Len(found) = 0 Then
                    found = cn
                ElseIf found <> cn Then
                    Exit Function                'ambiguous name across forms -> Object
                End If
            End If
        End If
    Next
    NativeControlTypeOf = found
done:
End Function

Private Function NativeControlClassName(ByVal t As Long) As String
    'VB intrinsic control type byte (modControls.ControlType enum) -> class name.
    Select Case t
        Case 0: NativeControlClassName = "PictureBox"
        Case 1: NativeControlClassName = "Label"
        Case 2: NativeControlClassName = "TextBox"
        Case 3: NativeControlClassName = "Frame"
        Case 4: NativeControlClassName = "CommandButton"
        Case 5: NativeControlClassName = "CheckBox"
        Case 6: NativeControlClassName = "OptionButton"
        Case 7: NativeControlClassName = "ComboBox"
        Case 8: NativeControlClassName = "ListBox"
        Case 9: NativeControlClassName = "HScrollBar"
        Case 10: NativeControlClassName = "VScrollBar"
        Case 11: NativeControlClassName = "Timer"
        Case 16: NativeControlClassName = "DriveListBox"
        Case 17: NativeControlClassName = "DirListBox"
        Case 18: NativeControlClassName = "FileListBox"
        Case 22: NativeControlClassName = "Shape"
        Case 23: NativeControlClassName = "Line"
        Case 24: NativeControlClassName = "Image"
        Case 37: NativeControlClassName = "Data"
        Case 38: NativeControlClassName = "OLE"
        Case Else: NativeControlClassName = ""   'menu/form/usercontrol/unknown -> leave Object
    End Select
End Function

Private Function NativeStartsWithFn(ByVal s As String, ByVal pipeList As String) As Boolean
    'True when s begins with one of the pipe-delimited function names followed by "(".
    Dim parts() As String, i As Long
    parts = Split(pipeList, "|")
    For i = 0 To UBound(parts)
        If Left$(s, Len(parts(i)) + 1) = parts(i) & "(" Then NativeStartsWithFn = True: Exit Function
    Next
End Function

Private Sub NativeMergeType(typ As Collection, ByVal var As String, ByVal newType As String)
    'Merge a new type signal for var into the running inference, reconciling
    'compatible families and degrading genuine conflicts to Variant.
    Dim cur As String
    cur = NativeColGet(typ, var)
    If Len(cur) = 0 Then NativeColPut typ, var, newType: Exit Sub
    If cur = newType Or cur = "Variant" Then Exit Sub
    If NativeIsObjType(cur) And NativeIsObjType(newType) Then
        If cur = "Object" Then NativeColPut typ, var, newType: Exit Sub      'prefer the specific class
        If newType = "Object" Then Exit Sub                                  'keep the specific class
        NativeColPut typ, var, "Object": Exit Sub                            'two classes -> Object
    End If
    If (cur = "Integer" Or cur = "Long") And (newType = "Integer" Or newType = "Long") Then NativeColPut typ, var, "Long": Exit Sub
    NativeColPut typ, var, "Variant"
End Sub

Private Function NativeIsObjType(ByVal t As String) As Boolean
    Select Case t
        Case "String", "Integer", "Long", "Boolean", "Double", "Single", "Byte", "Variant", "Currency", "Date"
            NativeIsObjType = False
        Case Else
            NativeIsObjType = True
    End Select
End Function

Private Function NativeHexVal(ByVal h As String) As Long
    'Parse a hex string (e.g. "1C", "100") to a Long.  Hand-rolled: CLng("&H..&")
    'errors on the string form and CLng("&HFFFF") sign-extends as 16-bit.
    Dim i As Long, ch As Long, acc As Long
    For i = 1 To Len(h)
        ch = Asc(UCase$(Mid$(h, i, 1)))
        If ch >= 48 And ch <= 57 Then
            acc = acc * 16 + (ch - 48)
        ElseIf ch >= 65 And ch <= 70 Then
            acc = acc * 16 + (ch - 55)
        Else
            Exit For
        End If
    Next
    NativeHexVal = acc
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
    '
    'Two header shapes are handled.  (1) STACK counter: the header IS the `cmp` and
    'the limit operand resolves at render time (memory limits / variables included).
    '(2) REGISTER counter: VB6 keeps the counter (and the limit) in registers and
    'RELOADS the limit register at the top of the loop, so the back-edge targets a
    '`mov limitReg, imm` and the cmp - often a 16-bit `cmp di,cx` - is a couple of
    'instructions later.  For shape (2) the render-time decode can't read the cmp, so
    'start/limit are captured here into NVForStart/NVForLimit.
    On Error GoTo done
    Dim n As Long, k As Long, hi As Long, ci As Long, bi As Long, i As Long, inst As CInstruction
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
        'Locate the header instruction (the back-edge target) in the array.
        hi = -1
        For i = 0 To bi - 1
            If arr(i).va = hdrVA Then hi = i: Exit For
        Next i
        If hi < 0 Or hi + 1 >= bi Then GoTo nextb
        'Find the compare: the header itself, or - for a register-counter loop - the
        'first cmp after one or more leading limit-setup `mov reg,imm` instructions.
        ci = hi
        Do While ci < bi And NativeMnem(arr(ci)) <> "CMP"
            If Not NativeIsMovRegImm(arr(ci)) Then GoTo nextb
            ci = ci + 1
        Loop
        If ci + 1 >= bi Then GoTo nextb
        'Compare = `cmp <ctr>, <limit>` immediately followed by an exit Jcc that
        'leaves the loop forward (past the back-edge).
        Dim ctrReg As Long
        ctrReg = NativeForCounterReg(arr(ci))
        If ctrReg < 0 Then GoTo nextb
        If (arr(ci + 1).cmdType And C_TYPEMASK) <> C_JMC Then GoTo nextb
        If Not NativeIsExitJcc(NativeMnem(arr(ci + 1))) Then GoTo nextb
        If arr(ci + 1).jmpConst <= arr(bi).va Then GoTo nextb
        'The counter must be incremented within the body.
        If Not NativeForHasIncrement(arr, ci + 2, bi - 1, ctrReg) Then GoTo nextb
        'The counter must be initialised to a CONSTANT before the compare (a real
        'For starts at 0/1/2...).  Rejects false matches where the "counter"
        'register was just loaded from memory, which produced degenerate output
        'like `For i = var_38 To var_38` or `For i = ebx To edi`.
        Dim startVal As String
        startVal = ""
        If Not NativeForStartIsConst(arr, ci, ctrReg, startVal) Then GoTo nextb
        'The header must be reached ONLY by this back-edge (a single, clean
        'counted loop - no extra continue-jumps that would need a second Next).
        Dim refs As Long
        refs = 0
        For i = 0 To n - 1
            If ((arr(i).cmdType And C_TYPEMASK) = C_JMP Or (arr(i).cmdType And C_TYPEMASK) = C_JMC) _
               And arr(i).jmpConst = hdrVA Then refs = refs + 1
        Next i
        If refs <> 1 Then GoTo nextb
        'Register-counter shape (the cmp isn't the header, or it is a word reg-reg
        'compare): capture start/limit now since the render-time decode can't read it.
        If ci <> hi Or NativeHas66(Replace(arr(ci).dump, " ", "")) Then
            Dim limitVal As String
            limitVal = NativeForLimitVal(arr, ci, ctrReg)
            If Len(limitVal) = 0 Then GoTo nextb           'unresolved limit -> leave for Do While
            If Len(startVal) = 0 Then startVal = "0"
            NativeColPut NVForStart, "F" & hdrVA, startVal
            NativeColPut NVForLimit, "F" & hdrVA, limitVal
        End If
        'Record.
        NativeColPut NVForHdr, "F" & hdrVA, CStr(ctrReg)
        NativeColPut NVForJmp, "F" & arr(bi).va, CStr(hdrVA)
        'Suppress everything between the header and the exit Jcc inclusive: the
        'limit-setup movs that aren't the header itself, the cmp, and the Jcc.  (When
        'the header IS the cmp this is just the Jcc, as before.)  The header
        'instruction at hi is replaced by the emitted For line.
        For i = hi + 1 To ci + 1
            NativeColPut NVForSkip, "F" & arr(i).va, "1"
        Next i
        NativeAddUnique NVSkipLabels, hdrVA                   'header label -> the For line
nextb:
    Next bi
done:
End Sub

Private Function NativeIsMovRegImm(inst As CInstruction) As Boolean
    'True for `mov reg, imm32` (B8+r) or `mov r/m32, imm32` with r/m a register
    '(C7 /0, md=3) - the limit-register setup that can precede a counted loop's cmp.
    Dim dump As String, n As Long, p As Long, op As Long, modrm As Long, md As Long, reg As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    If n < 1 Then Exit Function
    p = NativeOpStart(dump, n)
    op = NativeDumpByte(dump, p)
    If op >= &HB8 And op <= &HBF Then NativeIsMovRegImm = True: Exit Function
    If op = &HC7 Then
        modrm = NativeDumpByte(dump, p + 1)
        md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7
        If md = 3 And reg = 0 Then NativeIsMovRegImm = True
    End If
End Function

Private Function NativeForLimitVal(arr() As CInstruction, ByVal ci As Long, ByVal ctrReg As Long) As String
    'Resolve the limit operand of the header compare arr(ci) to a constant value.
    'Handles `cmp r,imm` (immediate) and `cmp r,r` / `cmp r16,r16` (the limit lives
    'in a register loaded by an earlier `mov limitReg, imm`).  Empty when the limit
    'is not a recoverable constant (e.g. a memory variable).
    Dim dump As String, n As Long, p As Long, op As Long, modrm As Long, md As Long, reg As Long, rm As Long
    Dim limReg As Long
    On Error Resume Next
    dump = Replace(arr(ci).dump, " ", "")
    n = Len(dump) \ 2
    p = NativeOpStart(dump, n)
    op = NativeDumpByte(dump, p)
    modrm = NativeDumpByte(dump, p + 1)
    md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
    Select Case op
        Case &H3D                       'cmp eax, imm32
            NativeForLimitVal = NativeNumFromBits(NativeDumpInt32(dump, p + 1)): Exit Function
        Case &H83                       'cmp r/m, imm8 (sign-extended)
            If md = 3 Then NativeForLimitVal = CStr(NativeDumpInt8(dump, n - 1))
            Exit Function
        Case &H81                       'cmp r/m, imm32
            If md = 3 Then NativeForLimitVal = NativeNumFromBits(NativeDumpInt32(dump, n - 4))
            Exit Function
        Case &H3B                       'cmp reg, r/m  -> counter is reg, limit is r/m
            If md <> 3 Then Exit Function              'memory limit (a variable) - not a constant
            limReg = rm
        Case &H39                       'cmp r/m, reg  -> counter is r/m, limit is reg
            limReg = reg
        Case Else
            Exit Function
    End Select
    'Limit lives in a register: find the nearest preceding `mov limReg, imm32`.
    Dim i As Long, lo As Long, d2 As String, n2 As Long, p2 As Long, o2 As Long
    Dim mr As Long, mmd As Long, mrg As Long, mrm As Long
    lo = ci - 16: If lo < 0 Then lo = 0
    For i = ci - 1 To lo Step -1
        d2 = Replace(arr(i).dump, " ", "")
        n2 = Len(d2) \ 2
        If n2 < 1 Then GoTo prev2
        p2 = NativeOpStart(d2, n2)
        o2 = NativeDumpByte(d2, p2)
        If o2 = (&HB8 + limReg) Then              'mov limReg, imm32
            NativeForLimitVal = NativeNumFromBits(NativeDumpInt32(d2, p2 + 1)): Exit Function
        End If
        mr = NativeDumpByte(d2, p2 + 1)
        mmd = (mr \ &H40) And 3: mrg = (mr \ 8) And 7: mrm = mr And 7
        If o2 = &H8B And mrg = limReg Then Exit Function                          'mov limReg, r/m
        If o2 = &H89 And mmd = 3 And mrm = limReg Then Exit Function              'mov limReg, reg
        If o2 = &H33 And mmd = 3 And mrg = limReg And mrm = limReg Then NativeForLimitVal = "0": Exit Function  'xor limReg,limReg
prev2:
    Next i
End Function

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
                            'Variant For loop: header preceded by __vbaVarForInit, the
                            'back-edge by __vbaVarForNext -> render For/Next, drop the calls.
                            Dim vi As Long, vn As Long, lo As Long, z As Long
                            vi = -1: vn = -1
                            lo = hi - 5: If lo < 0 Then lo = 0
                            For z = hi - 1 To lo Step -1
                                If InStr(NativeApiName(arr(z)), "__vbaVarForInit") > 0 Then vi = z: Exit For
                            Next z
                            lo = bi - 5: If lo < 0 Then lo = 0
                            For z = bi - 1 To lo Step -1
                                If InStr(NativeApiName(arr(z)), "__vbaVarForNext") > 0 Then vn = z: Exit For
                            Next z
                            If vi >= 0 And vn >= 0 Then
                                NativeColPut NVVarForSuppress, "V" & arr(vn).va, "1"
                                NativeColPut NVVarForInitLink, "V" & arr(vi).va, arr(j + 1).va & "|" & arr(bi).va
                            End If
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
    'field; 39 (cmp r/m32,r32 with r/m a register) -> rm; 83/81 (cmp r/m32,imm with
    'r/m a register) -> rm; 3D -> eax.  A 16-bit (0x66) compare is accepted ONLY for
    'the reg-reg form `cmp r16,r16` (VB6's register loop counter vs a limit register);
    'other 0x66 compares juggle low-word partials we cannot model, so stay rejected.
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long, md As Long
    NativeForCounterReg = -1
    If NativeMnem(inst) <> "CMP" Then Exit Function
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    modrm = NativeDumpByte(dump, i + 1)
    md = (modrm \ &H40) And 3
    If NativeHas66(dump) Then
        If Not ((op = &H3B Or op = &H39) And md = 3) Then Exit Function
    End If
    Select Case op
        Case &H3D                       'cmp eax, imm32
            NativeForCounterReg = 0
        Case &H3B                       'cmp r32, r/m32 -> counter is the reg field
            NativeForCounterReg = (modrm \ 8) And 7
        Case &H39                       'cmp r/m32, r32 -> counter is the r/m (register only)
            If md = 3 Then NativeForCounterReg = modrm And 7
        Case &H83, &H81                 'cmp r/m32, imm -> counter is the r/m (register only)
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

Private Function NativeForStartIsConst(arr() As CInstruction, ByVal hi As Long, ByVal ctrReg As Long, Optional ByRef startVal As String) As Boolean
    'A counted For initialises its counter to a constant (mov ctr,imm or xor
    'ctr,ctr) before the header.  Scan backward to the NEAREST writer of the
    'counter and require it to be such an init; any other writer first (or none
    'within range) means this is not a clean counted loop, so reject it.  When it is
    'such an init, startVal is set to the constant ("0" for xor/sub-to-self).
    Dim i As Long, lo As Long, dump As String, nn As Long, p As Long, op As Long, modrm As Long, md As Long, rg As Long, rm As Long
    On Error Resume Next
    lo = hi - 12: If lo < 0 Then lo = 0
    For i = hi - 1 To lo Step -1
        dump = Replace(arr(i).dump, " ", "")
        nn = Len(dump) \ 2
        If nn < 1 Then GoTo prev
        p = NativeOpStart(dump, nn)
        op = NativeDumpByte(dump, p)
        If op = (&HB8 + ctrReg) Then                                                   'mov ctr, imm
            startVal = NativeNumFromBits(NativeDumpInt32(dump, p + 1))
            NativeForStartIsConst = True: Exit Function
        End If
        modrm = NativeDumpByte(dump, p + 1)
        md = (modrm \ &H40) And 3: rg = (modrm \ 8) And 7: rm = modrm And 7
        If (op = &H33 Or op = &H31 Or op = &H2B Or op = &H29) And md = 3 And rg = ctrReg And rm = ctrReg Then startVal = "0": NativeForStartIsConst = True: Exit Function  'xor/sub ctr,ctr -> 0
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

Private Sub NativeDetectByteClamp(col As Collection)
    'VB6's Byte-member +/-1-with-clamp idiom (frmCreate's body-part Up/Down buttons,
    'e.g. `iHead = iHead - 1 : If iHead <= 0 Then iHead = 1`).  A Byte field is loaded,
    'incremented/decremented, range-coerced through __vbaUI1I2, stored back, then
    'clamped to a bound:
    '    movzx cx,byte[esi+off] ; mov edi,[__vbaUI1I2] ; (add|sub) cx,1 ; jo ovf ;
    '    call edi ; (cmp al,N | test al,al) ; mov [esi+off],al ; jcc skip ;
    '    mov ecx,M ; call edi ; mov [esi+off],al ; skip:
    'We otherwise drop the lot (the store has no NVReg source, the cmp/jcc on the
    'byte register yields a blank <cond>, and the body store is lost) - leaving an
    'empty `If <cond> Then / End If`.  Reconstruct it exactly; the unconditional
    'decrement store (which the commercial decompiler omits) is recovered too.  The
    'signature is rigid (esi = Me, __vbaUI1I2 confirmed via the IAT), so it cannot
    'misfire on unrelated code; a non-match leaves the proc untouched.
    On Error GoTo done
    Dim n As Long, k As Long, inst As CInstruction
    n = col.Count
    If n < 11 Then Exit Sub
    Dim arr() As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next
    For k = 0 To n - 11
        Dim d0 As String, off As Long
        d0 = Replace(arr(k).dump, " ", "")
        'movzx cx, byte[esi+disp32]  (66 0F B6 8E <d32>)
        If Len(d0) < 16 Then GoTo nextk
        If Not (NativeDumpByte(d0, 0) = &H66 And NativeDumpByte(d0, 1) = &HF _
                And NativeDumpByte(d0, 2) = &HB6 And NativeDumpByte(d0, 3) = &H8E) Then GoTo nextk
        off = NativeDumpInt32(d0, 4)
        'mov edi, [abs]  (8B 3D <abs32>) - confirm it is __vbaUI1I2
        Dim d1 As String, iat As Long
        d1 = Replace(arr(k + 1).dump, " ", "")
        If Not (NativeDumpByte(d1, 0) = &H8B And NativeDumpByte(d1, 1) = &H3D) Then GoTo nextk
        iat = NativeDumpInt32(d1, 2)
        If InStr(dsmNative.GetApiByIatVa(iat), "__vbaUI1I2") = 0 Then GoTo nextk
        '(add|sub) cx, 1  (66 83 C1 01 = add ecx ; 66 83 E9 01 = sub ecx)
        Dim d2 As String, opSign As String
        d2 = Replace(arr(k + 2).dump, " ", "")
        If Not (NativeDumpByte(d2, 0) = &H66 And NativeDumpByte(d2, 1) = &H83 _
                And NativeDumpByte(d2, 3) = &H1) Then GoTo nextk
        Select Case NativeDumpByte(d2, 2)
            Case &HC1: opSign = "+"
            Case &HE9: opSign = "-"
            Case Else: GoTo nextk
        End Select
        'jo ovf  (70 rel8)
        If NativeDumpByte(Replace(arr(k + 3).dump, " ", ""), 0) <> &H70 Then GoTo nextk
        'call edi  (FF D7)
        Dim d4 As String
        d4 = Replace(arr(k + 4).dump, " ", "")
        If Not (NativeDumpByte(d4, 0) = &HFF And NativeDumpByte(d4, 1) = &HD7) Then GoTo nextk
        'cmp al,imm8 (3C N) | test al,al (84 C0)
        Dim d5 As String, cmpVal As Long
        d5 = Replace(arr(k + 5).dump, " ", "")
        If NativeDumpByte(d5, 0) = &H3C Then
            cmpVal = NativeDumpByte(d5, 1)
        ElseIf NativeDumpByte(d5, 0) = &H84 And NativeDumpByte(d5, 1) = &HC0 Then
            cmpVal = 0
        Else
            GoTo nextk
        End If
        'mov [esi+off], al  (88 86 <d32>)
        Dim d6 As String
        d6 = Replace(arr(k + 6).dump, " ", "")
        If Not (NativeDumpByte(d6, 0) = &H88 And NativeDumpByte(d6, 1) = &H86 _
                And NativeDumpInt32(d6, 2) = off) Then GoTo nextk
        'jcc skip (short) - the body runs on the NEGATED condition
        Dim cmpOp As String
        cmpOp = NativeNegJccRel(NativeDumpByte(Replace(arr(k + 7).dump, " ", ""), 0))
        If Len(cmpOp) = 0 Then GoTo nextk
        'mov ecx, imm32  (B9 <imm32>)
        Dim d8 As String, bodyVal As Long
        d8 = Replace(arr(k + 8).dump, " ", "")
        If NativeDumpByte(d8, 0) <> &HB9 Then GoTo nextk
        bodyVal = NativeDumpInt32(d8, 1)
        'call edi (FF D7) ; mov [esi+off], al (88 86 <d32>)
        Dim d9 As String, d10 As String
        d9 = Replace(arr(k + 9).dump, " ", "")
        d10 = Replace(arr(k + 10).dump, " ", "")
        If Not (NativeDumpByte(d9, 0) = &HFF And NativeDumpByte(d9, 1) = &HD7) Then GoTo nextk
        If Not (NativeDumpByte(d10, 0) = &H88 And NativeDumpByte(d10, 1) = &H86 _
                And NativeDumpInt32(d10, 2) = off) Then GoTo nextk
        'Record the reconstruction at the anchor; suppress the 11 idiom instructions.
        NativeColPut NVByteClamp, "B" & arr(k).va, _
            opSign & "|" & cmpOp & "|" & cmpVal & "|" & bodyVal & "|" & off
        Dim j As Long
        For j = k To k + 10
            NativeColPut NVByteClampSkip, "B" & arr(j).va, "1"
        Next
nextk:
    Next k
done:
End Sub

Private Function NativeNegJccRel(ByVal op As Long) As String
    'The VB relational for the block a short Jcc SKIPS (i.e. the NEGATION of the Jcc's
    'taken condition), used to render the reconstructed `If <field> <rel> N`.
    Select Case op
        Case &H72: NativeNegJccRel = ">="   'jb  -> not below
        Case &H73: NativeNegJccRel = "<"    'jae -> not (>=)
        Case &H74: NativeNegJccRel = "<>"   'je  -> not equal
        Case &H75: NativeNegJccRel = "="    'jne -> equal
        Case &H76: NativeNegJccRel = ">"    'jbe -> above
        Case &H77: NativeNegJccRel = "<="   'ja  -> not above
        Case &H7C: NativeNegJccRel = ">="   'jl  -> not less
        Case &H7D: NativeNegJccRel = "<"    'jge -> not (>=)
        Case &H7E: NativeNegJccRel = ">"    'jle -> greater
        Case &H7F: NativeNegJccRel = "<="   'jg  -> not greater
    End Select
End Function

Private Function NativeEbpOp(ByVal dump As String, ByVal wantOp As Long, ByRef disp As Long) As Boolean
    'True when the instruction is `wantOp` (8D lea / 8B mov-from / 89 mov-to) with an
    '[ebp +/- disp] operand (ModR/M rm = 5 = ebp, mod = 01/10).  Sets disp (signed).
    Dim n As Long, i As Long, op As Long, modrm As Long, md As Long, rm As Long
    On Error GoTo no
    dump = Replace(dump, " ", "")
    n = Len(dump) \ 2
    i = NativeOpStart(dump, n)
    op = NativeDumpByte(dump, i)
    If op <> wantOp Then Exit Function
    If i + 1 >= n Then Exit Function
    modrm = NativeDumpByte(dump, i + 1)
    md = (modrm \ &H40) And 3: rm = modrm And 7
    If rm <> 5 Then Exit Function                'ebp base only
    If md = 1 Then
        disp = NativeDumpInt8(dump, i + 2): NativeEbpOp = True
    ElseIf md = 2 Then
        disp = NativeDumpInt32(dump, i + 2): NativeEbpOp = True
    End If
no:
End Function

Private Sub NativeDetectPropDir(col As Collection)
    'Per indirect vtable CALL, decide whether its by-ref stack local is a value-OUT
    '(read AFTER the call before any overwrite/next call -> a GET, `var = obj.Prop`) or a
    'value-IN (not read after -> a PUT, `obj.Prop = value`).  VB6 flags a property Get and
    'Let identically in the FuncDesc, so this use-after data-flow is the reliable signal.
    'Recorded for every call; consumed only by the class-instance property render path.
    On Error GoTo done
    Dim n As Long, k As Long, inst As CInstruction
    n = col.Count
    If n < 2 Then Exit Sub
    Dim arr() As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next
    For k = 0 To n - 1
        If (arr(k).cmdType And C_TYPEMASK) <> C_CAL Then GoTo nextk
        'The by-ref local: nearest `lea reg,[ebp-D]` in the few preceding instructions
        '(the out-param / value pointer pushed to the accessor).
        Dim d As Long, j As Long, leaJ As Long, dd As Long
        leaJ = -1
        For j = k - 1 To k - 7 Step -1
            If j < 0 Then Exit For
            If NativeEbpOp(arr(j).dump, &H8D, d) Then leaJ = j: Exit For
            If (arr(j).cmdType And C_TYPEMASK) = C_CAL Then Exit For
        Next j
        If leaJ < 0 Then GoTo nextk
        'PUT vs GET by WRITE-BEFORE: a put STORES its value into the local before the
        'call (mov [ebp-D],r8/r32 = 88/89, or mov [ebp-D],imm = C7); a get's retbuf is a
        'fresh out-param (no prior store).  Scan the window before the call for a store to
        'D.  More robust than read-after, which a Variant-returning get defeats (its result
        'is moved out by a helper CALL, not a plain mov).
        Dim isPut As Boolean
        isPut = False
        For j = k - 1 To k - 10 Step -1
            If j < 0 Then Exit For
            If NativeEbpOp(arr(j).dump, &H89, dd) Then
                If dd = d Then isPut = True: Exit For
            End If
            If NativeEbpOp(arr(j).dump, &H88, dd) Then
                If dd = d Then isPut = True: Exit For
            End If
            If NativeEbpOp(arr(j).dump, &HC7, dd) Then
                If dd = d Then isPut = True: Exit For
            End If
        Next j
        NativeColPut NVPropDir, "P" & arr(k).va, IIf(isPut, "put", "get")
nextk:
    Next k
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

Private Sub NativeDetectElemIndices(col As Collection)
    'SAFEARRAY UDT-element access `Player(i).field` is addressed as
    '[pvData + (i-lBound)*cbElements + fieldOff].  The element byte-offset
    '(i-lBound)*cbElements lives in a register VB computes with lea/shl/imul; we
    'otherwise drop it (rendering global_X(12)(field), which collapses two different
    'elements to the same text).  Here, for every SIB element access, back-trace the
    'index register through the addressing chain to the LOGICAL index (the param/var
    'that was dereferenced, e.g. arg_8) and record it by VA, so the renderers emit
    'global_X(12)(arg_8)(field).  A trace only succeeds when it passes a scaling step
    '(shl/imul/lea-scale) or the lBound subtract - so it never fires on a plain
    'pointer register (the pvData side), only on a genuine element offset.
    On Error GoTo done
    Dim n As Long, k As Long, inst As CInstruction
    n = col.Count
    If n < 2 Then Exit Sub
    Dim arr() As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next
    For k = 0 To n - 1
        Dim dump As String, sb As Long, si As Long
        dump = Replace(arr(k).dump, " ", "")
        sb = NativeMemBase(dump): si = NativeMemIndex(dump)
        If sb < 0 Or si < 0 Then GoTo nextk          'not a [base+index] SIB access
        Dim origin As String
        origin = NativeTraceIndexOrigin(arr, n, k, si)
        If Len(origin) > 0 Then NativeColPut NVElemIdx, "E" & arr(k).va, origin
nextk:
    Next k
done:
End Sub

Private Sub NativeDetectLateCalls(col As Collection)
    'VB6 late-bound dispatch `obj.Member(args)` compiles to __vbaLateIdCall(obj, DISPID,
    'flags) (cdecl).  The DISPID is a `push imm` among the args; the nested object-getter
    '/ __vbaObjSet reset our push stack, so the dispid is gone by render time.  Here we
    'scan a backward window from each __vbaLateIdCall and record the candidate immediate
    'pushes (DISPIDs) by call VA; the renderer resolves them against the OCX typelib.
    On Error GoTo done
    Dim n As Long, k As Long, inst As CInstruction
    n = col.Count
    If n < 2 Then Exit Sub
    Dim arr() As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next
    For k = 0 To n - 1
        If (arr(k).cmdType And C_TYPEMASK) <> C_CAL Then GoTo nextk
        Dim an As String: an = NativeResolveCallApi(arr, k)
        'Id family (member by DISPID): __vbaLateIdCall/CallLd/CallSt/St/StAd/NamedCall.
        'Mem family (member by NAME string): __vbaLateMem*.
        If InStr(an, "__vbaLateId") = 0 And InStr(an, "__vbaLateMem") = 0 Then GoTo nextk
        'For an Id call collect the candidate immediate pushes (the DISPID); a Mem call
        'carries the member NAME as a string arg, so it needs no DISPID.
        If InStr(an, "__vbaLateId") > 0 Then
            Dim cand As String, j As Long, lim As Long, pv As Long
            cand = ""                          'reset per call (VB6 Dim-in-loop does NOT)
            lim = k - 30: If lim < 0 Then lim = 0
            For j = k - 1 To lim Step -1
                If (arr(j).cmdType And C_TYPEMASK) = C_CAL Then
                    If InStr(NativeResolveCallApi(arr, j), "__vbaLateId") > 0 Then Exit For   'previous late call -> stop
                End If
                If NativeIsPushImm(arr(j), pv) Then
                    If pv <> 0 Then cand = cand & CStr(pv) & ","     '0 is the flags arg, never a member id
                End If
            Next j
            If Len(cand) > 0 Then NativeColPut NVLateDispid, "L" & arr(k).va, cand
        End If
        'A `__vbaObjSet` that is the FIRST call scanning back from this late call (only
        'push/lea/mov plumbing between them) exists solely to pass the just-fetched
        'control into the call - mark it so the redundant `Set temp = control` drops.
        Dim jb As Long, blim As Long
        blim = k - 10: If blim < 0 Then blim = 0
        For jb = k - 1 To blim Step -1
            If (arr(jb).cmdType And C_TYPEMASK) = C_CAL Then
                If NativeBackCallIsObjSet(arr, jb) Then NativeColPut NVSuppressObjSet, "O" & arr(jb).va, "1"
                Exit For
            End If
        Next jb
nextk:
    Next k
done:
End Sub

Private Sub NativeDetectControlArrays(col As Collection)
    'VB6 control-array element access `ctrl(i)` compiles to the array object's Item
    'accessor at vtable offset 0x40:
    '    call [meVt + 0xNNN]          ; form accessor -> the ARRAY control object (eax)
    '    push eax / lea ecx,[arrLoc] / push ecx / call __vbaObjSet  ; store array obj
    '    mov esi,eax / mov edi,[esi]                                ; esi=array, edi=arrVt
    '    lea  REG_RB, [ebp-DISP_RB]   ; element retbuf (out-param)
    '    push REG_RB
    '    mov  ecx, [ebp-DISP_IDX]     ; the index
    '    call [IAT]                   ; __vbaI4 (coerce index)
    '    push eax                     ; index
    '    push esi                     ; this (array obj)
    '    call [edi + 0x40]            ; element accessor -> element returned in [ebp-DISP_RB]
    'We recover (array control name+GUID, index local, element retbuf local) ENTIRELY
    'from the disassembly - not from the live push/register state, which the nested
    '__vbaI4 / __vbaObjSet calls corrupt - and record it by the 0x40 call's VA.  At
    'render the call becomes `Set var_<rb> = Form.ctrl(var_<idx>)` (var_<rb> tagged with
    'the array's control GUID), so the following `mov reg,[ebp-DISP_RB]; mov vt,[reg];
    'call [vt+0x54/0x19C]` resolves to var_<rb>.Caption / .ToolTipText via NativeControlProp.
    'Gated strictly (is-array control + offset 0x40 + a fully recovered index & retbuf);
    'a miss leaves the raw UnkVCall form rather than risk a wrong control name.
    On Error GoTo done
    Dim n As Long, k As Long, inst As CInstruction
    n = col.Count
    If n < 4 Then Exit Sub
    Dim arr() As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next

    Dim baseReg As Long, disp As Long
    For k = 3 To n - 1
        If (arr(k).cmdType And C_TYPEMASK) <> C_CAL Then GoTo nextk
        If Not NativeIsIndirectMemCall(arr(k), baseReg, disp) Then GoTo nextk
        If disp <> &H40 Then GoTo nextk          'the element Item accessor offset

        'Collect the 3 stack args scanning back: pushReg(0)=this (closest to the call),
        '(1)=index, (2)=retbuf (pushed first / bottom of stack).
        Dim pushIdx(2) As Long, pushReg(2) As Long, pc As Long, j As Long, lo As Long
        pc = 0
        lo = k - 16: If lo < 0 Then lo = 0
        Dim preg As Long
        For j = k - 1 To lo Step -1
            If NativePushReg(arr(j), preg) Then
                If pc <= 2 Then pushReg(pc) = preg: pushIdx(pc) = j
                pc = pc + 1
                If pc >= 3 Then Exit For
            End If
        Next j
        If pc < 3 Then GoTo nextk

        'Retbuf local: the lea that set the bottom-push register.
        Dim rbDisp As Long, rbReg As Long, found As Boolean, ldReg As Long, ldDisp As Long
        rbReg = pushReg(2): found = False
        For j = pushIdx(2) - 1 To pushIdx(2) - 4 Step -1
            If j < 0 Then Exit For
            If NativeLeaLocal(arr(j), ldReg, ldDisp) Then
                If ldReg = rbReg And ldDisp < 0 Then rbDisp = ldDisp: found = True: Exit For
            End If
            If NativeInstDestReg(arr(j)) = rbReg Then Exit For   'register reused for something else
        Next j
        If Not found Then GoTo nextk

        'Index token: the local feeding the index coerce (__vbaI4).  The coerce input
        'is loaded into ecx as `mov ecx,[ebp-Y]` (-> var_Y) or `mov ecx,REG` where REG
        'is the loop counter mirroring a spill local (-> that var_Y).
        Dim idxTok As String
        idxTok = NativeArrIndexTok(arr, k, lo)
        If Len(idxTok) = 0 Then GoTo nextk

        'Array control: the nearest preceding indirect form-control accessor
        '`call [reg+offN]` (the __vbaObjSet / __vbaI4 calls are `call [abs]`, skipped).
        'It must resolve to a control that is a control ARRAY.
        Dim accDisp As Long, accReg As Long, ctlName As String, ctlGuid As String, gotArr As Boolean, accJ As Long
        gotArr = False: accJ = -1
        For j = k - 1 To lo Step -1
            If (arr(j).cmdType And C_TYPEMASK) = C_CAL Then
                If NativeIsIndirectMemCall(arr(j), accReg, accDisp) Then
                    If NVBase >= 0 And accDisp >= NVBase Then
                        ctlName = NativeControlByOffset(accDisp)
                        If Len(ctlName) > 0 And NativeControlIsArrayByOffset(accDisp) Then
                            ctlGuid = NativeGuidByOffset(accDisp): gotArr = True: accJ = j
                        End If
                    End If
                    Exit For        'nearest indirect-reg call is the form accessor; stop either way
                End If
            End If
        Next j
        If Not gotArr Then GoTo nextk
        If Len(ctlGuid) = 0 Then GoTo nextk

        'Record the reconstruction for the render pass.
        Dim elemExpr As String
        elemExpr = NVForm & "." & ctlName & "(" & idxTok & ")"
        NativeColPut NVCtlArrElem, "K" & arr(k).va, elemExpr
        NativeColPut NVCtlArrGuid, "K" & arr(k).va, ctlGuid
        NativeColPut NVCtlArrRetbuf, "K" & arr(k).va, CStr(rbDisp)

        'Drop the redundant `Set <arrTemp> = Form.ctrl` that stores the array object:
        'the __vbaObjSet just after the form accessor exists only to feed this element
        'call (we reconstruct the element directly).  Mark it for the ObjSet handler.
        Dim jf As Long, fhi As Long
        fhi = accJ + 5: If fhi > k - 1 Then fhi = k - 1
        For jf = accJ + 1 To fhi
            If (arr(jf).cmdType And C_TYPEMASK) = C_CAL Then
                If InStr(NativeApiName(arr(jf)), "__vbaObjSet") > 0 Then
                    NativeColPut NVSuppressObjSet, "O" & arr(jf).va, "1": Exit For
                End If
            End If
        Next jf
nextk:
    Next k
done:
End Sub

Private Sub NativeDetectKeyCompares(col As Collection)
    'VB6 `Select Case <IntegerByRefParam>` (e.g. Form_KeyDown's KeyCode) loads the param
    'ONCE into a callee-saved 16-bit register (`mov base,[ebp+P]; mov di,word[base]`),
    'then tests each case via `mov ecx,const; call __vbaI2I4; cmp di,ax` or
    '`cmp di,imm16`.  The generic decoder bails on these 16-bit register compares (the
    'boolean-juggling guard), leaving `<cond>`.  This pre-pass recognises ONLY this shape
    'and records, per cmp VA, the resolved operands (paramTok | const) so the Jcc renders
    '`If KeyCode = 97`.  It changes NO global register state (the earlier whole-proc
    '16-bit-deref tracking bled wrong values into other procs - reverted); the 16-bit
    'compares are EXCLUSIVE to the case chain (the 32-bit `cmp edi,eax` SAFEARRAY
    'bounds-checks in the case bodies are a different opcode and never match).
    On Error GoTo done
    Dim n As Long, k As Long, inst As CInstruction
    n = col.Count
    If n < 3 Then Exit Sub
    Dim arr() As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next

    'Anchor: mov <r16>, word[base] where base <- mov base,[ebp+P] (a ByRef param).
    Dim keyReg As Long, keyParam As String, baseReg As Long, foundAnchor As Boolean
    Dim j As Long, P As Long, gotP As Boolean
    foundAnchor = False
    For k = 1 To n - 1
        If NativeIsMovR16FromMem(arr(k), keyReg, baseReg) Then
            gotP = False
            For j = k - 1 To 0 Step -1
                If NativeIsMovRegFromParam(arr(j), baseReg, P) Then gotP = True: Exit For
                If NativeInstDestReg(arr(j)) = baseReg Then Exit For   'base reused for something else
            Next j
            If gotP Then keyParam = "arg_" & Hex$(P): foundAnchor = True: Exit For
        End If
    Next k
    If Not foundAnchor Then Exit Sub

    'Record every 16-bit case compare of the key register.
    Dim rhs As String
    For k = 0 To n - 1
        If NativeKeyCaseRhs(arr, k, keyReg, rhs) Then
            NativeColPut NVKeyCmp, "K" & arr(k).va, keyParam & "|" & rhs
        End If
    Next k
done:
End Sub

Private Sub NativeDetectAbsGlobalTests(col As Collection)
    'A `mov eax,[abs-global]; test eax,eax; jcc` condition where the load uses the
    'short-form opcode 0xA1 (mov eax, moffs32).  That form is NOT register-tracked
    '(broad 0xA1 tracking regressed the SIB array-pointer detection - Gap A, reverted),
    'so the test leaked as a raw `If eax > 0`.  Record the global token at the
    '`test eax,eax` VA so the compare renders `global_X <op> 0` (e.g. `If Game.pIndex > 0`,
    'global_0042304C).  VA-scoped: it touches NO register state, so the Gap-A cascade
    'cannot recur.
    On Error GoTo done
    Dim n As Long, k As Long, arr() As CInstruction, inst As CInstruction
    n = col.Count
    If n < 2 Then Exit Sub
    ReDim arr(n - 1): k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next
    For k = 0 To n - 2
        Dim d As String, gva As Long
        d = Replace(arr(k).dump, " ", "")
        If Len(d) >= 10 And Left$(d, 2) = "A1" Then          'A1 <abs32> = mov eax,[abs32]
            gva = NativeDumpInt32(d, 1)
            If NativeIsGlobalAddr(gva) Then
                If NativeIsTestEaxEax(arr(k + 1)) Then       'immediately tested
                    NativeColPut NVAbsGlobalCmp, "T" & arr(k + 1).va, NativeGlobalName(gva)
                End If
            End If
        End If
    Next k
done:
End Sub

Private Function NativeIsMovR16FromMem(inst As CInstruction, ByRef destReg As Long, ByRef baseReg As Long) As Boolean
    'Match `mov r16, word[base]` (0x66 prefix + 0x8B + memory operand); set destReg
    '(the loaded 16-bit register) and baseReg (the memory base register).
    Dim dump As String, n As Long, i As Long, op As Long, modrm As Long, md As Long
    On Error GoTo no
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    If Not NativeHas66(dump) Then GoTo no
    i = NativeOpStart(dump, n)
    op = NativeDumpByte(dump, i)
    If op <> &H8B Then GoTo no
    modrm = NativeDumpByte(dump, i + 1)
    md = (modrm \ &H40) And 3
    If md = 3 Then GoTo no
    destReg = (modrm \ 8) And 7
    baseReg = NativeMemBase(dump)
    If baseReg < 0 Then GoTo no
    NativeIsMovR16FromMem = True
    Exit Function
no:
    NativeIsMovR16FromMem = False
End Function

Private Function NativeIsMovRegFromParam(inst As CInstruction, ByVal reg As Long, ByRef P As Long) As Boolean
    'Match `mov <reg>, [ebp + P]` with P > 0 (a procedure parameter slot); set P.
    Dim dump As String, n As Long, i As Long, op As Long, modrm As Long, md As Long, rf As Long, disp As Long, isAbs As Boolean
    On Error GoTo no
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    If NativeHas66(dump) Then GoTo no
    i = NativeOpStart(dump, n)
    op = NativeDumpByte(dump, i)
    If op <> &H8B Then GoTo no
    modrm = NativeDumpByte(dump, i + 1)
    md = (modrm \ &H40) And 3: rf = (modrm \ 8) And 7
    If md = 3 Or rf <> reg Then GoTo no
    If NativeMemIndex(dump) >= 0 Then GoTo no
    If NativeMemBase(dump) <> 5 Then GoTo no          'base must be ebp
    If Not NativeDecodeDisp(dump, disp, isAbs) Then GoTo no
    If isAbs Or disp <= 0 Then GoTo no
    P = disp
    NativeIsMovRegFromParam = True
    Exit Function
no:
    NativeIsMovRegFromParam = False
End Function

Private Function NativeKeyCaseRhs(arr() As CInstruction, ByVal k As Long, ByVal keyReg As Long, ByRef rhs As String) As Boolean
    'A 16-bit case compare of the key register: `cmp keyReg,ax` (const from the preceding
    '`mov ecx,imm` feeding __vbaI2I4) or `cmp keyReg,imm16` (const inline).  Sets rhs.
    Dim dump As String, n As Long, i As Long, op As Long, modrm As Long, md As Long, reg As Long, rm As Long
    On Error GoTo no
    dump = Replace(arr(k).dump, " ", "")
    n = Len(dump) \ 2
    If Not NativeHas66(dump) Then GoTo no             'only the 16-bit case compares
    i = NativeOpStart(dump, n)
    op = NativeDumpByte(dump, i)
    modrm = NativeDumpByte(dump, i + 1)
    md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
    If op = &H3B Then                                  'cmp reg, r/m  -> cmp keyReg, ax
        If reg = keyReg And md = 3 And rm = 0 Then
            Dim jj As Long, c As Long
            For jj = k - 1 To k - 4 Step -1
                If jj < 0 Then Exit For
                If NativeIsMovEcxImm(arr(jj), c) Then rhs = CStr(c And &HFFFF&): NativeKeyCaseRhs = True: Exit Function
            Next jj
        End If
    ElseIf op = &H81 Then                              'grp1 r/m, imm16 ; /7 = cmp
        If md = 3 And rm = keyReg And reg = 7 Then
            rhs = CStr(NativeDumpInt16(dump, n - 2) And &HFFFF&)
            NativeKeyCaseRhs = True
        End If
    End If
    Exit Function
no:
    NativeKeyCaseRhs = False
End Function

Private Function NativeIsMovEcxImm(inst As CInstruction, ByRef val As Long) As Boolean
    'Match `mov ecx, imm32` (0xB9 id); set val.
    Dim dump As String, n As Long, i As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    i = NativeOpStart(dump, n)
    If NativeDumpByte(dump, i) = &HB9 Then val = NativeDumpInt32(dump, i + 1): NativeIsMovEcxImm = True
End Function

Private Function NativeIsIndirectMemCall(inst As CInstruction, ByRef baseReg As Long, ByRef disp As Long) As Boolean
    'True when inst is `call [reg + disp]` (FF /2, memory via a base register, NOT an
    'absolute [disp32] nor `call reg`).  Returns the base register (0..7) and disp.
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long, regf As Long, isAbs As Boolean
    baseReg = -1: disp = 0
    On Error GoTo no
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op <> &HFF Then GoTo no
    modrm = NativeDumpByte(dump, i + 1)
    regf = (modrm \ 8) And 7
    If regf <> 2 Then GoTo no                 'call r/m is /2
    If (modrm \ &H40) = 3 Then GoTo no        'call reg (register-direct), not memory
    baseReg = NativeMemBase(dump)
    If baseReg < 0 Then GoTo no               'absolute / no-base SIB
    If Not NativeDecodeDisp(dump, disp, isAbs) Then GoTo no
    If isAbs Then GoTo no
    NativeIsIndirectMemCall = True
    Exit Function
no:
    baseReg = -1: NativeIsIndirectMemCall = False
End Function

Private Function NativePushReg(inst As CInstruction, ByRef reg As Long) As Boolean
    'Match a single-byte `push r32` (0x50..0x57); set reg to the pushed register index.
    Dim dump As String, nn As Long, i As Long, op As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op >= &H50 And op <= &H57 Then reg = op - &H50: NativePushReg = True
End Function

Private Function NativeLeaLocal(inst As CInstruction, ByRef reg As Long, ByRef disp As Long) As Boolean
    'Match `lea r32, [ebp + disp]` (8D /r, base ebp, no SIB index); set reg and disp.
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long, isAbs As Boolean
    On Error GoTo no
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op <> &H8D Then GoTo no
    modrm = NativeDumpByte(dump, i + 1)
    If (modrm \ &H40) = 3 Then GoTo no
    reg = (modrm \ 8) And 7
    If NativeMemIndex(dump) >= 0 Then GoTo no
    If NativeMemBase(dump) <> 5 Then GoTo no       'base must be ebp
    If Not NativeDecodeDisp(dump, disp, isAbs) Then GoTo no
    If isAbs Then GoTo no
    NativeLeaLocal = True
    Exit Function
no:
    NativeLeaLocal = False
End Function

Private Function NativeArrIndexTok(arr() As CInstruction, ByVal callIdx As Long, ByVal loBound As Long) As String
    'Recover the array index token for a control-array element accessor `call
    '[arrVt+0x40]` at callIdx.  VB coerces the index with __vbaI4 (a `call [abs]`) just
    'before pushing it: `mov ecx,SRC ; call [__vbaI4] ; push eax`.  Resolve SRC (the ecx
    'load feeding the coerce) to a local token - either a direct local [ebp-Y] or a
    'register holding the loop counter that mirrors a spill local.  "" on any miss.
    Dim j As Long, cc As Long, dd As Long, ab As Boolean
    cc = -1
    For j = callIdx - 1 To loBound Step -1
        If (arr(j).cmdType And C_TYPEMASK) = C_CAL Then
            If NativeDecodeDisp(arr(j).dump, dd, ab) Then
                If ab Then cc = j               'absolute call = the index coerce
            End If
            Exit For                            'first call back is the coerce (abs) or the form accessor; stop
        End If
    Next j
    If cc < 0 Then Exit Function
    Dim sd As Long, sr As Long, kind As Long, yd As Long
    For j = cc - 1 To loBound Step -1
        kind = NativeMovEcxSrc(arr(j), sd, sr)
        If kind = 1 Then
            If sd < 0 Then NativeArrIndexTok = "var_" & Hex$(Abs(sd))
            Exit Function
        ElseIf kind = 2 Then
            'The index register mirrors the loop counter, spilled at the loop top -
            'often well before the tight element-access window - so search a wider
            'backward range for its nearest spill/fill local.
            Dim wlo As Long
            wlo = j - 80: If wlo < 0 Then wlo = 0
            If NativeRegSpillLocal(arr, j, sr, wlo, yd) Then NativeArrIndexTok = "var_" & Hex$(Abs(yd))
            Exit Function
        End If
    Next j
End Function

Private Function NativeMovEcxSrc(inst As CInstruction, ByRef localDisp As Long, ByRef srcReg As Long) As Long
    'Classify a `mov ecx, SRC` instruction: 1 = ecx <- [ebp+disp] (sets localDisp),
    '2 = ecx <- register (sets srcReg), 0 = not a mov-into-ecx.
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long, md As Long, reg As Long, isAbs As Boolean
    On Error GoTo no
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op <> &H8B Then GoTo no               'mov r32, r/m32
    modrm = NativeDumpByte(dump, i + 1)
    md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7
    If reg <> 1 Then GoTo no                  'dest must be ecx
    If md = 3 Then
        srcReg = modrm And 7
        NativeMovEcxSrc = 2
        Exit Function
    End If
    If NativeMemIndex(dump) >= 0 Then GoTo no
    If NativeMemBase(dump) <> 5 Then GoTo no   'base must be ebp
    If Not NativeDecodeDisp(dump, localDisp, isAbs) Then GoTo no
    If isAbs Then GoTo no
    NativeMovEcxSrc = 1
    Exit Function
no:
    NativeMovEcxSrc = 0
End Function

Private Function NativeRegSpillLocal(arr() As CInstruction, ByVal fromIdx As Long, ByVal reg As Long, ByVal loBound As Long, ByRef disp As Long) As Boolean
    'Resolve a register to the stack local it mirrors, by finding the nearest preceding
    'spill `mov [ebp-Y], reg` (store) or fill `mov reg, [ebp-Y]` (load).  Used to map a
    'register-held loop counter back to its var_Y for a control-array index.
    Dim j As Long, dmp As String, nn As Long, i As Long, op As Long, modrm As Long
    Dim md As Long, rf As Long, rm As Long, dd As Long, ab As Boolean
    For j = fromIdx - 1 To loBound Step -1
        dmp = Replace(arr(j).dump, " ", "")
        nn = Len(dmp) \ 2
        i = NativeOpStart(dmp, nn)
        op = NativeDumpByte(dmp, i)
        If op = &H89 Or op = &H8B Then          '89=store reg->r/m, 8B=load r/m->reg
            modrm = NativeDumpByte(dmp, i + 1)
            md = (modrm \ &H40) And 3: rf = (modrm \ 8) And 7: rm = modrm And 7
            If md <> 3 And rf = reg And NativeMemIndex(dmp) < 0 And NativeMemBase(dmp) = 5 Then
                If NativeDecodeDisp(arr(j).dump, dd, ab) Then
                    If Not ab And dd < 0 Then disp = dd: NativeRegSpillLocal = True: Exit Function
                End If
            End If
        End If
    Next j
End Function

Private Function NativeCallReg(inst As CInstruction) As Long
    'Register index of an indirect `call reg` (FF /2, mod=3), else -1.
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long
    NativeCallReg = -1
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op <> &HFF Then Exit Function
    modrm = NativeDumpByte(dump, i + 1)
    If (modrm \ &H40) = 3 And ((modrm \ 8) And 7) = 2 Then NativeCallReg = modrm And 7
End Function

Private Function NativeResolveCallApi(arr() As CInstruction, ByVal callIdx As Long) As String
    'API name of a call, resolving an INDIRECT `call reg` to the IAT the register was
    'loaded from (VB caches a helper IAT into a callee-saved reg, then calls it many
    'times - e.g. `mov edi,[__vbaLateIdSt]; ...; call edi`).  Without this, the
    'late-call pre-pass misses every register-cached helper call (its DISPID is never
    'collected), so the property put renders as a bare `Call LateIdSt()`.
    NativeResolveCallApi = NativeApiName(arr(callIdx))
    If Len(NativeResolveCallApi) > 0 Then Exit Function
    Dim rg As Long, j As Long
    rg = NativeCallReg(arr(callIdx))
    If rg < 0 Then Exit Function
    For j = callIdx - 1 To 0 Step -1
        If NativeInstDestReg(arr(j)) = rg Then
            NativeResolveCallApi = NativeApiName(arr(j))
            Exit Function
        End If
    Next j
End Function

Private Function NativeBackCallIsObjSet(arr() As CInstruction, ByVal callIdx As Long) As Boolean
    'True when the call at callIdx is __vbaObjSet - either direct (call [iat]) or via a
    'register VB cached it into (`mov edi,[__vbaObjSet iat]; call edi`).
    If InStr(NativeApiName(arr(callIdx)), "__vbaObjSet") > 0 Then NativeBackCallIsObjSet = True: Exit Function
    Dim rg As Long
    rg = NativeCallReg(arr(callIdx))
    If rg < 0 Then Exit Function
    'Find the most recent load of that register (VB caches the helper IAT into a
    'callee-saved register once near the proc top, then calls it many times), so scan
    'the whole preceding range, not a short window.
    Dim j As Long
    For j = callIdx - 1 To 0 Step -1
        If NativeInstDestReg(arr(j)) = rg Then
            NativeBackCallIsObjSet = (InStr(NativeApiName(arr(j)), "__vbaObjSet") > 0)
            Exit Function
        End If
    Next j
End Function

Private Function NativeIsPushImm(inst As CInstruction, ByRef val As Long) As Boolean
    'Match `push imm8` (6A ib) or `push imm32` (68 id); set val to the pushed value.
    Dim dump As String, nn As Long, i As Long, op As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op = &H6A Then
        val = NativeDumpInt8(dump, i + 1): NativeIsPushImm = True
    ElseIf op = &H68 Then
        val = NativeDumpInt32(dump, i + 1): NativeIsPushImm = True
    End If
End Function

Private Function NativeCtlBaseName(ByVal s As String) As String
    'From an object expression like "frmClient.Winsock1" extract the control base name
    '"Winsock" (drop the form prefix, any (index), and trailing digits) for matching the
    'right OCX among the project's references.
    Dim t As String, p As Long, ch As Long
    t = Trim$(s)
    p = InStrRev(t, ".")
    If p > 0 Then t = Mid$(t, p + 1)
    p = InStr(t, "(")
    If p > 0 Then t = Left$(t, p - 1)
    Do While Len(t) > 0
        ch = Asc(Right$(t, 1))
        If ch >= 48 And ch <= 57 Then t = Left$(t, Len(t) - 1) Else Exit Do
    Loop
    NativeCtlBaseName = t
End Function

Private Function NativeLateIdCall(inst As CInstruction, ByVal apiName As String, ByRef resolved As Boolean) As String
    'Render a late-bound dispatch call obj.Member by resolving its DISPID (recorded by
    'NativeDetectLateCalls) against the control's OCX typelib.  Sets resolved=True (and
    'returns the statement / "" for a value) when it resolves; resolved stays False so
    'the caller falls back to the raw Call otherwise.  Variants:
    '  __vbaLateIdCall   - a method/sub call, result discarded -> statement obj.M[(a)]
    '  __vbaLateIdCallLd - returns a value (a property Get / function) -> thread the
    '                      expression through eax so the consumer assigns it
    '  __vbaLateIdCallSt - a property/array store -> statement obj.M = value
    resolved = False
    Dim objExpr As String, cand As String, baseName As String, memName As String, invKind As Long
    objExpr = NativeArgList()
    NVPushTop = 0
    If Len(objExpr) = 0 Then Exit Function
    'The object is the control-reference token (form.control) - NOT necessarily the
    'first arg: __vbaLateIdCall is (obj, dispid,..) but __vbaLateIdCallLd is
    '(result, obj, ..).  Pick the token that is a `<...>.<name>` reference.
    Dim toks() As String, ti As Long, objTok As String, resTok As String
    toks = Split(objExpr, ", ")
    For ti = 0 To UBound(toks)
        Dim tkv As String: tkv = Trim$(toks(ti))
        If Len(objTok) = 0 And InStr(tkv, ".") > 0 And InStr(tkv, "(") = 0 Then
            objTok = tkv
        ElseIf Len(resTok) = 0 And (Left$(tkv, 4) = "var_" Or Left$(tkv, 4) = "arg_") Then
            resTok = tkv                     'the Ld result buffer (a lea'd local)
        End If
    Next ti
    If Len(objTok) = 0 Then Exit Function
    baseName = NativeCtlBaseName(objTok)
    If Len(baseName) = 0 Then Exit Function
    Dim libClass As String
    libClass = GetControlClass(objTok)               'exact external class (e.g. TabDlg.SSTab)
    cand = NativeColGet(NVLateDispid, "L" & inst.va)
    If Len(cand) = 0 Then Exit Function
    'Which invoke kind this call wants - breaks ties when one memid has both a
    'method and a property on the control's typelib (1=Func, 2=Get, 4=Put).
    Dim wantKind As Long
    If InStr(apiName, "St") > 0 Then
        wantKind = 4
    ElseIf InStr(apiName, "Ld") > 0 Then
        wantKind = 2
    Else
        wantKind = 1
    End If
    Dim parts() As String, pi As Long, d As Long
    parts = Split(cand, ",")
    For pi = 0 To UBound(parts)
        If Len(parts(pi)) > 0 Then
            d = CLng(parts(pi))
            memName = modCOM.LateMemberName(libClass, baseName, d, wantKind, invKind)
            If Len(memName) > 0 Then Exit For
        End If
    Next pi
    If Len(memName) = 0 Then Exit Function
    resolved = True
    Dim mref As String
    mref = objTok & "." & memName
    If InStr(apiName, "St") > 0 Or invKind = 4 Then
        'A property/array STORE (__vbaLateIdSt/StAd/CallSt, or a PROPERTYPUT member):
        'obj.Member = value.  The value is usually a Variant temp; if the only thing
        'available is the object itself (the push stack held the control), don't echo
        'it as the RHS - leave a <value> placeholder.
        Dim stVal As String
        stVal = NativePopValue()
        If Len(stVal) = 0 Or stVal = objTok Or stVal = "<value>" Then stVal = "<value>"
        'The value is usually built into a Variant DATA field just before the call
        '(the push stack only carried the object), so recover it from there - and
        'mark that temp's build statements (var_C(4/8/12)=...) for removal.
        If stVal = "<value>" And Len(NVLastVarData) > 0 Then
            stVal = NVLastVarData
            If Len(NVLastVarBase) > 0 Then NativeSuppressVarBuild NVLastVarBase
        End If
        NVLastVarData = "": NVLastVarBase = ""
        NativeLateIdCall = mref & " = " & stVal
    ElseIf InStr(apiName, "Ld") > 0 Then
        'A value (property Get / function).  __vbaLateIdCallLd writes the result into a
        'lea'd buffer (the var_ token) - assign to it so the value is captured; if no
        'buffer was recovered, thread through eax for the consumer instead.
        If Len(resTok) > 0 Then
            NativeLateIdCall = resTok & " = " & mref
        Else
            NVReg(0) = mref
            NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
            NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjGuid(0) = "": NVRegObjVtGuid(0) = ""
            NVPendingCall = "Call " & mref     'flushed as a statement if the result is unused
            NativeLateIdCall = ""
        End If
    Else
        'A method/sub call.  Its arguments are passed as Variants by pointer (the push
        'stack was reset by the nested object-getter), so use the Variant data-field
        'values built just before the call - e.g. Winsock1.Connect "127.0.0.1", 535 -
        'and strip their build statements.
        NativeLateIdCall = mref & NativeTakeVarArgs()
    End If
End Function

Private Function NativeTakeVarArgs() As String
    'Consume the pending Variant-method-argument list as " a, b, c" (leading space),
    'mark each temp's build statements for removal, and reset.  "" if none.
    Dim s As String, vk As Long
    For vk = 0 To NVVarArgN - 1
        If Len(s) > 0 Then s = s & ", "
        s = s & NVVarArgList(vk)
        NativeSuppressVarBuild NVVarArgBase(vk)
    Next
    NativeResetVarArgs
    If Len(s) > 0 Then NativeTakeVarArgs = " " & s
End Function

Private Function NativeLateMemCall(inst As CInstruction, ByVal apiName As String, ByRef resolved As Boolean) As String
    'Name-based late dispatch: __vbaLateMemCall[Ld/St](result?, obj, "Member", cArgs..)
    'carries the member name as a STRING argument, so it needs no typelib - render
    'obj.Member directly.  Variants mirror the Id form (Ld -> result = obj.Member,
    'St / PROPERTYPUT -> obj.Member = value, plain -> obj.Member).
    resolved = False
    Dim objExpr As String
    objExpr = NativeArgList()
    NVPushTop = 0
    If Len(objExpr) = 0 Then Exit Function
    Dim toks() As String, ti As Long, memName As String, memIdx As Long, objTok As String, resTok As String
    toks = Split(objExpr, ", ")
    memIdx = -1
    For ti = 0 To UBound(toks)                    'the member name is the first quoted token
        If Left$(Trim$(toks(ti)), 1) = Chr$(34) And memIdx < 0 Then memName = Trim$(toks(ti)): memIdx = ti
    Next ti
    If memIdx < 1 Or Len(memName) < 3 Then Exit Function
    memName = Mid$(memName, 2, Len(memName) - 2)   'strip the quotes
    If Len(memName) = 0 Or Not NativeIsIdentName(memName) Then Exit Function
    objTok = Trim$(toks(memIdx - 1))              'object = the token just before the name
    If Len(objTok) = 0 Then Exit Function
    For ti = 0 To memIdx - 2                       'result buffer (Ld) = the first var_/arg_ before the object
        If Left$(Trim$(toks(ti)), 4) = "var_" Or Left$(Trim$(toks(ti)), 4) = "arg_" Then resTok = Trim$(toks(ti)): Exit For
    Next ti
    resolved = True
    Dim mref As String
    mref = objTok & "." & memName
    If InStr(apiName, "Ld") > 0 Then
        If Len(resTok) > 0 Then
            NativeLateMemCall = resTok & " = " & mref
        Else
            NVReg(0) = mref
            NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
            NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjGuid(0) = "": NVRegObjVtGuid(0) = ""
            NVPendingCall = "Call " & mref
            NativeLateMemCall = ""
        End If
    ElseIf InStr(apiName, "St") > 0 Then
        Dim mstVal As String
        mstVal = NativePopValue()
        If Len(mstVal) = 0 Or mstVal = objTok Or mstVal = "<value>" Then mstVal = "<value>"
        If mstVal = "<value>" And Len(NVLastVarData) > 0 Then
            mstVal = NVLastVarData
            If Len(NVLastVarBase) > 0 Then NativeSuppressVarBuild NVLastVarBase
        End If
        NVLastVarData = "": NVLastVarBase = ""
        NativeLateMemCall = mref & " = " & mstVal
    Else
        NativeLateMemCall = mref & NativeTakeVarArgs()
    End If
End Function

Private Function NativeIsIdentName(ByVal s As String) As Boolean
    'A plausible member identifier (letters/digits/underscore, leading letter) - guards
    'against treating a non-name string literal as a member name.
    Dim i As Long, ch As Long
    If Len(s) = 0 Then Exit Function
    ch = Asc(UCase$(Left$(s, 1)))
    If Not ((ch >= 65 And ch <= 90) Or ch = 95) Then Exit Function
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        If Not ((ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Or (ch >= 48 And ch <= 57) Or ch = 95) Then Exit Function
    Next
    NativeIsIdentName = True
End Function

Private Function NativeTraceIndexOrigin(arr() As CInstruction, ByVal n As Long, ByVal startK As Long, ByVal idxReg As Long) As String
    'Walk backwards from the SIB access at startK, following the register that holds
    'the element byte-offset, through scaling (shl/lea/imul), the lBound subtract, and
    'reg-reg / deref movs, to the named source of the logical index.  Returns "" unless
    'a scaling/lBound step was seen (proof it is an element offset, not a raw pointer).
    Dim k As Long, cur As Long, steps As Long, sawScale As Boolean
    cur = idxReg
    For k = startK - 1 To 0 Step -1
        steps = steps + 1
        If steps > 80 Then Exit Function
        Dim ins As CInstruction: Set ins = arr(k)
        If (ins.cmdType And C_TYPEMASK) = C_CAL Then
            'VB's array bookkeeping (the bounds-error stub on the jumped-over error
            'path, the array lock/unlock) sits inside the addressing sequence but does
            'not really run before the access / preserves the index - trace through it.
            Dim cnm As String: cnm = NativeApiName(ins)
            If InStr(cnm, "BoundsError") > 0 Or InStr(cnm, "AryLock") > 0 Or InStr(cnm, "AryUnlock") > 0 Then GoTo contk
            If cur <= 2 Then Exit Function            'a real call clobbers caller-saved regs
            GoTo contk                                'callee-saved (ebx/esi/edi/ebp) survive
        End If
        If NativeInstDestReg(ins) <> cur Then GoTo contk
        Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long, md As Long, rmf As Long, op2 As Long
        dump = Replace(ins.dump, " ", "")
        nn = Len(dump) \ 2
        i = NativeOpStart(dump, nn)
        op = NativeDumpByte(dump, i)
        modrm = NativeDumpByte(dump, i + 1)
        md = (modrm \ &H40) And 3: rmf = modrm And 7
        Select Case op
            Case &HC1, &HD1, &HD3                     'shl/shr -> scale in place
                sawScale = True
            Case &H8D                                 'lea dst,[base+idx*s+disp] -> trace base
                sawScale = True
                Dim lb As Long: lb = NativeMemBase(dump)
                If lb >= 0 And lb <= 7 Then cur = lb Else Exit Function
            Case &H69, &H6B                           'imul dst, r/m, imm -> trace r/m
                sawScale = True
                If md = 3 Then cur = rmf Else Exit Function
            Case &H2B                                 'sub dst,[reg+14] -> lBound subtract
                Dim sdisp As Long, sAbs As Boolean
                If md = 3 Then Exit Function
                If Not NativeDecodeDisp(dump, sdisp, sAbs) Then Exit Function
                If sAbs Or sdisp <> &H14 Then Exit Function
                sawScale = True                       'drop the lBound adjust, keep cur
            Case &H8B                                 'mov dst, r/m
                If md = 3 Then
                    cur = rmf                          'reg-reg copy
                Else
                    Dim mdisp As Long, mAbs As Boolean, mbase As Long
                    If Not NativeDecodeDisp(dump, mdisp, mAbs) Then Exit Function
                    mbase = NativeMemBase(dump)
                    If mAbs And NativeIsGlobalAddr(mdisp) Then
                        If sawScale Then NativeTraceIndexOrigin = NativeGlobalName(mdisp)
                        Exit Function
                    ElseIf Not mAbs And mbase = 5 And mdisp >= 8 Then
                        If sawScale Then NativeTraceIndexOrigin = "arg_" & Hex$(mdisp)   'ByVal index param
                        Exit Function
                    ElseIf Not mAbs And mbase = 5 And mdisp < 0 Then
                        If sawScale Then NativeTraceIndexOrigin = NativeGetLocalExpr(mdisp)
                        Exit Function
                    ElseIf Not mAbs And mbase >= 0 And mbase <= 7 And mdisp = 0 And NativeMemIndex(dump) < 0 Then
                        cur = mbase                    'deref of a ByRef index pointer -> trace it
                    Else
                        Exit Function
                    End If
                End If
            Case &HF
                op2 = NativeDumpByte(dump, i + 1)
                If op2 = &HAF Then                     'imul dst, r/m
                    sawScale = True
                    Dim m2 As Long: m2 = NativeDumpByte(dump, i + 2)
                    If (m2 \ &H40) = 3 Then cur = m2 And 7 Else Exit Function
                ElseIf op2 = &HBE Or op2 = &HBF Or op2 = &HB6 Or op2 = &HB7 Then   'movsx/movzx
                    Dim m3 As Long: m3 = NativeDumpByte(dump, i + 2)
                    If (m3 \ &H40) = 3 Then cur = m3 And 7 Else Exit Function
                Else
                    Exit Function
                End If
            Case &HB8, &HB9, &HBA, &HBB, &HBC, &HBD, &HBE, &HBF
                'mov dst, imm: the index register originates at a constant.  This is
                'almost always a REGISTER-RESIDENT LOOP COUNTER's stale init value
                '(e.g. `For i = 1` -> `mov edi,1`), NOT a real literal array index, so
                'binding it would mis-render Player(i) as Player(1).  Bail (leave the
                'index dropped) rather than emit a wrong constant.
                Exit Function
            Case Else
                Exit Function                          'unrecognised write to cur -> bail
        End Select
contk:
    Next k
End Function

Private Function NativeInstDestReg(inst As CInstruction) As Long
    'Destination register index for the reg-writing instructions the index back-trace
    'follows (lea/shl/imul/mov-to-reg/sub-to-reg/movsx/movzx/mov-imm), else -1.
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long, op2 As Long
    NativeInstDestReg = -1
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    Select Case op
        Case &H8D, &H8B, &H2B, &H3, &H69, &H6B        'lea/mov-r/sub-r/imul3 -> reg field
            modrm = NativeDumpByte(dump, i + 1)
            NativeInstDestReg = (modrm \ 8) And 7
        Case &HC1, &HD1, &HD3                          'shl/shr group -> rm (md=3 only)
            modrm = NativeDumpByte(dump, i + 1)
            If (modrm \ &H40) = 3 Then NativeInstDestReg = modrm And 7
        Case &HB8, &HB9, &HBA, &HBB, &HBC, &HBD, &HBE, &HBF   'mov reg, imm32
            NativeInstDestReg = op - &HB8
        Case &HF
            op2 = NativeDumpByte(dump, i + 1)
            Select Case op2
                Case &HAF, &HBE, &HBF, &HB6, &HB7      'imul / movsx / movzx -> reg field
                    modrm = NativeDumpByte(dump, i + 2)
                    NativeInstDestReg = (modrm \ 8) And 7
            End Select
    End Select
End Function

Private Function NativeIs16BitAddSub(inst As CInstruction) As Boolean
    'True for a 16-bit (0x66) in-place arithmetic that folds its dest's word shadow
    '(NVR16Val) in NativeTrackReg - the per-instruction shadow-clear must skip these so
    'the read-then-write (`cx = cx * 1000` / `cx = cx - 55`) sees the prior value:
    '  add/sub r16,r/m16   (reg-form 0x03 / 0x2B)
    '  imul r16,r/m16,imm  (0x69 / 0x6B)
    '  add/sub r/m16,imm   (group-1 0x83 / 0x81, reg field 0 = add, 5 = sub)
    Dim dump As String, n As Long, i As Long, op As Long, modrm As Long, rf As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    n = Len(dump) \ 2
    If Not NativeHas66(dump) Then Exit Function
    i = NativeOpStart(dump, n)
    op = NativeDumpByte(dump, i)
    Select Case op
        Case &H3, &H2B, &H69, &H6B
            NativeIs16BitAddSub = True
        Case &H83, &H81
            modrm = NativeDumpByte(dump, i + 1)
            rf = (modrm \ 8) And 7
            NativeIs16BitAddSub = (rf = 0 Or rf = 5)
    End Select
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
        'DIRECT-TEST form: `call __vbaStrCmp; test eax,eax; jcc` - the raw result is
        'tested IMMEDIATELY (no materialisation, no store).  Resolve via NativeResolveCallApi
        'so a register-cached strcmp (`mov edi,[iat]; call edi`, the map parser) is caught
        'too.  Record it so the handler hands the operands to that test (-> `If p1 = p2`).
        'Must be the very next instruction: when the result is first stored to a local
        '(`mov var_B0,eax; ... ; test var_B0`, the config parsers) this does NOT match, so
        'that strcmp keeps its visible Call (no lost comparison).
        Dim cnmAny As String
        cnmAny = NativeResolveCallApi(arr, k)
        If (InStr(cnmAny, "__vbaStrCmp") > 0 Or InStr(cnmAny, "__vbaStrComp") > 0 Or InStr(cnmAny, "__vbaStrTextCmp") > 0) Then
            If k + 1 <= n - 1 Then
                If NativeIsTestEaxEax(arr(k + 1)) Then
                    NativeColPut NVStrCmpDirect, "P" & arr(k).va, "1"
                    GoTo nexts
                End If
            End If
        End If
        'Materialised form (neg/sbb/inc/neg) - direct `call [iat]` only, unchanged.
        Dim disp As Long, isAbs As Boolean
        If Not NativeDecodeDisp(arr(k).dump, disp, isAbs) Then GoTo nexts
        If Not isAbs Then GoTo nexts
        Dim cnm As String
        cnm = dsmNative.GetApiByIatVa(disp)
        If InStr(cnm, "__vbaStrCmp") = 0 And InStr(cnm, "__vbaStrComp") = 0 _
           And InStr(cnm, "__vbaStrTextCmp") = 0 Then GoTo nexts   'Text = Option Compare Text
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

Private Function NativeIsTestEaxEax(inst As CInstruction) As Boolean
    'Match `test eax, eax` (85 C0): the direct test of a __vbaStrCmp tri-state result.
    Dim dump As String, nn As Long, i As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    If NativeHas66(dump) Then Exit Function
    NativeIsTestEaxEax = (NativeDumpByte(dump, i) = &H85 And NativeDumpByte(dump, i + 1) = &HC0)
End Function

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

Private Sub NativeDetectSelectConst(col As Collection)
    'VB compiles a two-constant select `If cond Then x=c1 Else x=c2` BRANCHLESSLY:
    '   xor r,r ; cmp.. ; setcc r8 ; dec r ; and r,mask ; add r,base
    'whose result is base (cond true) or base+mask (cond false).  Record (base, base+mask)
    'at the setcc VA so the renderer binds IIf(cond, base, base+mask), and mark the
    'dec/and/add tail to skip.  The literal `((cond)-1 And mask)+base` is both cryptic and
    'WRONG in VB (setcc is 0/1 but a VB Boolean is -1/0); IIf is correct and readable.
    On Error GoTo done
    Dim n As Long, k As Long, inst As CInstruction
    n = col.Count
    If n < 4 Then Exit Sub
    Dim arr() As CInstruction
    ReDim arr(n - 1)
    k = 0
    For Each inst In col: Set arr(k) = inst: k = k + 1: Next
    Dim ccOp As String, ccReg As Long, mask As Long, base As Long
    For k = 0 To n - 4
        If NativeIsSetcc(arr(k), ccOp, ccReg) Then
            If NativeIsDecReg(arr(k + 1), ccReg) _
               And NativeIsAluRegImm(arr(k + 2), 4, ccReg, mask) _
               And NativeIsAluRegImm(arr(k + 3), 0, ccReg, base) Then
                NativeColPut NVSelConst, "S" & arr(k).va, CStr(base) & "|" & CStr(base + mask)
                NativeColPut NVSelConstSkip, "S" & arr(k + 1).va, "1"
                NativeColPut NVSelConstSkip, "S" & arr(k + 2).va, "1"
                NativeColPut NVSelConstSkip, "S" & arr(k + 3).va, "1"
            End If
        End If
    Next
done:
End Sub

Private Function NativeIsDecReg(inst As CInstruction, ByVal reg As Long) As Boolean
    'Match `dec r32` (single-byte 0x48+reg, or FF /1 md=3 rm=reg).
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op = &H48 + reg Then NativeIsDecReg = True: Exit Function
    If op = &HFF Then
        modrm = NativeDumpByte(dump, i + 1)
        If (modrm \ &H40) = 3 And (((modrm \ 8) And 7) = 1) And ((modrm And 7) = reg) Then NativeIsDecReg = True
    End If
End Function

Private Function NativeIsAluRegImm(inst As CInstruction, ByVal grp As Long, ByVal reg As Long, ByRef imm As Long) As Boolean
    'Match a grp1 ALU op `<grp> r32, imm8/imm32` (0x83 sign-extended imm8, or 0x81 imm32),
    'register-direct, where grp is the /digit (0=add, 4=and, 5=sub).  Returns imm (signed).
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op <> &H83 And op <> &H81 Then Exit Function
    modrm = NativeDumpByte(dump, i + 1)
    If (modrm \ &H40) <> 3 Then Exit Function            'register-direct
    If ((modrm \ 8) And 7) <> grp Then Exit Function     'the /digit selects the ALU op
    If (modrm And 7) <> reg Then Exit Function
    If op = &H83 Then imm = NativeDumpInt8(dump, i + 2) Else imm = NativeDumpInt32(dump, i + 2)
    NativeIsAluRegImm = True
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

Private Function NativeIsCallExpr(ByVal s As String) As Boolean
    'True when s is a tracked function/value-call expression like `IsNumeric(var_4C)`
    'or `Environ$("TEMP")` - first char an identifier char, contains "(", ends ")".
    'Used so a 16-bit `test ax,ax` of a folded predicate result (VARIANT_BOOL, which
    'VB word-tests) resolves to `<expr> <> 0` instead of being dropped as a word-reg
    'partial.  A leading "(" (an already-parenthesised relational) is handled separately.
    Dim c As String
    s = Trim$(s)
    If Len(s) < 3 Then Exit Function
    c = Left$(s, 1)
    If Not ((c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or c = "_") Then Exit Function
    NativeIsCallExpr = (InStr(s, "(") > 0 And Right$(s, 1) = ")")
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
    'Also `dec reg` (48+reg, or FF /1) - the common 1-based index (Select Case on a
    '1..N value compiles to `movsx reg,[v]; dec reg; cmp reg,N-1; jmp [reg*4+tbl]`), so
    'case k corresponds to value k+1.  Without this `dec` the case values were off by one.
    Dim dump As String, nn As Long, i As Long, op As Long, modrm As Long
    On Error Resume Next
    dump = Replace(inst.dump, " ", "")
    nn = Len(dump) \ 2
    i = NativeOpStart(dump, nn)
    op = NativeDumpByte(dump, i)
    If op = &H48 + reg Then imm = 1: NativeSubRegImm = True: Exit Function      'dec reg (1-byte)
    If op = &HFF Then
        modrm = NativeDumpByte(dump, i + 1)
        If (modrm \ &H40) = 3 And (((modrm \ 8) And 7) = 1) And ((modrm And 7) = reg) Then imm = 1: NativeSubRegImm = True
        Exit Function
    End If
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
    NVCurVa = inst.va          'so the SIB read/store renderers can look up NVElemIdx

    'A write to a register invalidates the 16-bit word shadow for that register (the
    'setter, a 16-bit mov-from-memory, re-populates it below).  This keeps a captured
    'array-element / Integer-field word from being read by a later, unrelated
    '`cmp ax,imm16` after the register was clobbered.  EXCEPTION: a 16-bit add/sub
    'reads-then-writes its dest's shadow (the fold in NativeTrackReg updates it in
    'place to `(a + b)`), so it must not be pre-cleared.
    Dim r16Dest As Long
    r16Dest = NativeInstDestReg(inst)
    If r16Dest >= 0 And Not NativeIs16BitAddSub(inst) Then NVR16Val(r16Dest) = ""

    'TEST/CMP set the flags consumed by the next conditional jump.  Record the
    'operands now (the relational operator is resolved later from the Jcc).
    If mn = "TEST" Or mn = "CMP" Then
        'A recognised Select-Case-on-Integer-parameter case compare (`cmp di,<const>`):
        'use the pre-pass-recovered operands (<param> = <const>) instead of the generic
        'decode, which bails on the 16-bit register compare and leaves <cond>.
        Dim kcRec As String, kcBar As Long
        kcRec = NativeColGet(NVKeyCmp, "K" & inst.va)
        If Len(kcRec) > 0 Then
            kcBar = InStr(kcRec, "|")
            NVCmpL = Left$(kcRec, kcBar - 1): NVCmpR = Mid$(kcRec, kcBar + 1)
            NVCmpIsTest = False: NVCmpIsBool = False: NVCmpSet = True
            Exit Function
        End If
        'A `mov eax,[abs-global] (0xA1); test eax,eax` condition: the short-form global
        'load is NOT register-tracked (broad 0xA1 tracking regressed the SIB array-pointer
        'detection - Gap A, reverted), so a VA-scoped pre-pass binds the global token here
        '-> `If global_X > 0` instead of a raw `If eax > 0`.
        Dim agRec As String
        agRec = NativeColGet(NVAbsGlobalCmp, "T" & inst.va)
        If Len(agRec) > 0 Then
            NVCmpL = agRec: NVCmpR = "0": NVCmpIsTest = True: NVCmpIsBool = False: NVCmpSet = True
            Exit Function
        End If
        NativeDecodeCompare inst, mn
        Exit Function
    End If

    'SETcc materialises a relational into a register: bind "(L <op> R)" from the
    'pending compare so the value flows into a following boolean combine / test.
    Dim ccOp As String, ccReg As Long
    If NativeIsSetcc(inst, ccOp, ccReg) Then
        If NVCmpSet And Len(NVCmpL) > 0 Then
            Dim ccR As String, scRec As String
            If NVCmpIsTest Then ccR = "0" Else ccR = NVCmpR
            scRec = NativeColGet(NVSelConst, "S" & inst.va)
            If Len(ccR) = 0 Then
                NVReg(ccReg) = ""
            ElseIf Len(scRec) > 0 Then
                'Branchless select-of-two-constants: this setcc heads a dec/and/add tail
                'that yields base (cond true) or base+mask (cond false) - render IIf.
                Dim scBar As Long
                scBar = InStr(scRec, "|")
                NVReg(ccReg) = "IIf(" & NVCmpL & " " & ccOp & " " & ccR & ", " & Left$(scRec, scBar - 1) & ", " & Mid$(scRec, scBar + 1) & ")"
            Else
                NVReg(ccReg) = "(" & NVCmpL & " " & ccOp & " " & ccR & ")"
            End If
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
                'Control-array element accessor: a deterministic pre-pass
                '(NativeDetectControlArrays) recovered the array control, index and
                'element retbuf for this `call [arrayVt + 0x40]`.  Emit
                'Set var_<rb> = Form.ctrl(idx) and tag the retbuf local with the array's
                'control GUID, so the following `mov reg,[ebp-rb]; mov vt,[reg];
                'call [vt+0x54/0x19C]` resolves to var_<rb>.Caption / .ToolTipText.  The
                'COM HRESULT left in eax is error-checked next, so reset it like UnkVCall.
                Dim caKey As String
                caKey = "K" & inst.va
                If Len(NativeColGet(NVCtlArrElem, caKey)) > 0 Then
                    Dim caElem As String, caGuid As String, caRb As Long, caName As String
                    caElem = NativeColGet(NVCtlArrElem, caKey)
                    caGuid = NativeColGet(NVCtlArrGuid, caKey)
                    caRb = CLng(NativeColGet(NVCtlArrRetbuf, caKey))
                    caName = "var_" & Hex$(Abs(caRb))
                    NVPushTop = 0
                    NativeResetValue
                    NativeSetLocalExpr caRb, caName
                    NativeSetLocalGuid caRb, caGuid
                    NativeProcessInst = ind & "Set " & caName & " = " & caElem & vbCrLf
                    Exit Function
                End If
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
                'Value-returning METHOD on a tracked intrinsic object (Clipboard.GetData):
                'the vtable register carries the intrinsic object's name.  Fold the result
                'to the consumer (var_X = Clipboard.GetData), like a Property Get; the
                'optional format argument (built as a DISP_E_PARAMNOTFOUND Variant) is
                'omitted.
                If ocb >= 0 And ocb <= 7 Then
                    If Len(NVRegObjVt(ocb)) > 0 Then
                        Dim imName As String
                        imName = NativeIntrinsicMethodByOffset(NVRegObjVt(ocb), disp)
                        If Len(imName) > 0 Then
                            Dim imChain As String, imLn As String
                            imChain = NVRegObjVt(ocb) & "." & imName
                            NVPushTop = 0
                            'The method returns via a hidden retbuf (the most recent LEA);
                            'the HRESULT is left in eax and error-checked next, so DON'T
                            'leave imChain in eax (it would leak into the HRESULT compare).
                            'Emit the assignment to the retbuf local so the call is visible
                            'and its value flows on to the consumer (var = Clipboard.GetData).
                            NVReg(0) = ""
                            NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                            NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjGuid(0) = "": NVRegObjVtGuid(0) = ""
                            If NVLastLeaSet Then
                                imLn = "var_" & Hex$(Abs(NVLastLea))
                                NativeSetLocalExpr NVLastLea, imLn
                                NVLastLeaSet = False
                                NativeProcessInst = ind & imLn & " = " & imChain & vbCrLf
                            End If
                            Exit Function
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
                            'Method kind (for the this/retbuf arithmetic below).
                            Dim umKind As String, umVal As String, umIsVal As Boolean
                            NativeTryMethodKind umAddr, umKind
                            'Args (source order) lead with the implicit `this` pointer and
                            'may end with a hidden [out,retval] return buffer.  Whether the
                            'this is PRESENT in the captured list varies (a New-local
                            'receiver pushes it as a stale "0"; an As-New FIELD receiver does
                            'not), so decide by COUNT: a value-returning method has a retbuf,
                            'so expected-without-this = params + retbuf; any EXTRA leading
                            'arg beyond that is the implicit this and is dropped.  (The old
                            'gate umP(0)=umRecv missed the this when it rendered as "0" -
                            '`testPacket.ID = 0` instead of 25 - while unconditionally
                            'dropping arg 0 ate a real arg of the field-receiver calls.)
                            umArgs = NativeArgList()
                            Dim umP() As String, umStart As Long, umKeep As Long, umI As Long, umOut As String
                            Dim umTotal As Long, umRetbuf As String, umNP As Long, umRb As Long, umExp As Long
                            umTotal = 0: umRetbuf = ""
                            If Len(umArgs) > 0 Then
                                umP = Split(umArgs, ", ")
                                umNP = 0
                                If NativeTryMethodSig(umAddr, umSig) Then umNP = NativeArgCount(umSig)
                                umRb = IIf(InStr(umKind, "Function") > 0 Or InStr(umKind, "Get") > 0, 1, 0)
                                umExp = umNP + umRb                        'args expected when `this` is absent
                                umStart = 0
                                If (UBound(umP) + 1) > umExp Then umStart = 1   'extra leading arg = the implicit this
                                umTotal = UBound(umP) - umStart + 1
                                umKeep = umNP
                                If umKeep > umTotal Then umKeep = umTotal
                                For umI = umStart To umStart + umKeep - 1
                                    If umI > UBound(umP) Then Exit For
                                    If Len(umOut) > 0 Then umOut = umOut & ", "
                                    'A by-ref Variant temp built for an untyped (Variant)
                                    'param resolves to the literal value the caller passed:
                                    'SetupCalc(55) instead of SetupCalc(var_34).
                                    umOut = umOut & NativeResolveVarArg(umP(umI))
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
                            'A Property LET/SET on a user-class instance (e.g.
                            'testPacket.ID = 25 / testPacket.Name = "Jonathan").  The
                            'FuncDesc kind is authoritative (b1=0 -> Let/Set), so this is
                            'a STORE, not a value-returning call: render `recv.Prop = value`
                            '(`Set recv.Prop = value` for a Set).  Without this it fell
                            'through to the value path and the put was consumed by the
                            'following HRESULT check as a bogus `If recv.Prop(0) = 0 Then`.
                            'The stored value is the method's single real parameter (umArgs,
                            'after the leading `this` was dropped above).
                            If InStr(umKind, "Property Let") > 0 Or InStr(umKind, "Property Set") > 0 Then
                                Dim umSetKw As String
                                umSetKw = IIf(InStr(umKind, "Property Set") > 0, "Set ", "")
                                If Len(umVal) = 0 Then umVal = umArgs
                                If Len(umVal) = 0 Then umVal = umRetbuf
                                NVPushTop = 0
                                NVReg(0) = ""                       'put returns nothing - clear the HRESULT slot
                                NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjInst(0) = ""
                                NativeProcessInst = ind & umSetKw & umRecv & "." & umName & " = " & umVal & vbCrLf
                                Exit Function
                            End If
                            'A property GET/method on a LOCAL class-instance receiver (an
                            '`As New`/predeclared instance, var_X) - e.g. cmdOK_Click reading
                            'pktCreate.Name/.Life/... into locals.  Emit the read as a VISIBLE
                            'statement `var_<rb> = recv.Prop` so the resolved member shows
                            '(beats commercial's UnkVCall) instead of silently folding into an
                            'unconsumed local and vanishing.  Scoped to local-instance
                            'receivers (umRecv "var_") so As-New FIELD receivers (clsBitmap
                            'field_X, whose gets are consumed inline) keep their current fold.
                            If Len(umRetbuf) > 0 And Left$(umRecv, 4) = "var_" And Left$(umRetbuf, 4) = "var_" _
                               And InStr(umKind, "Property") > 0 Then
                                'A property accessor on a LOCAL class-instance (As New /
                                'predeclared, e.g. pktCreate).  VB6 flags Get and Let alike,
                                'so the DIRECTION comes from the data-flow pre-pass
                                '(NVPropDir): the by-ref local read AFTER the call = a GET
                                '(var = obj.Prop), else a PUT (obj.Prop = value).  Either way
                                'the resolved member NAME shows (beats commercial's UnkVCall)
                                'instead of silently folding into an unconsumed local.
                                Dim umDir As String, umRbD As Long, umPv As String
                                umDir = NativeColGet(NVPropDir, "P" & inst.va)
                                On Error Resume Next
                                umRbD = -CLng("&H" & Mid$(umRetbuf, 5))
                                On Error GoTo 0
                                NVPushTop = 0
                                If umDir = "put" Then
                                    'Value being stored = what the by-ref local holds (its
                                    'tracked expression if meaningful, else the local name).
                                    umPv = NVLocal("L" & umRbD)
                                    If Len(umPv) = 0 Or umPv = "0" Or IsNumeric(umPv) Then umPv = umRetbuf
                                    NativeProcessInst = ind & umRecv & "." & umName & " = " & umPv & vbCrLf
                                Else
                                    NativeProcessInst = ind & umRetbuf & " = " & umRecv & "." & umName & vbCrLf
                                    On Error Resume Next
                                    NativeSetLocalExpr umRbD, umRetbuf      'read result; reference by name
                                    On Error GoTo 0
                                End If
                                NVReg(0) = ""
                                Exit Function
                            End If
                            umIsVal = (Len(umRetbuf) > 0)
                            If Not umIsVal Then
                                umIsVal = (InStr(umKind, "Get") > 0)
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
                                'A Property GET returns its value through the retbuf
                                '(eax = HRESULT) and the value is consumed through that
                                'local; CLEAR eax so the folded expression cannot linger
                                'and leak into a later argument push (picItemBmp.MaskDC
                                'duplicated into BitBlt's ySrc).  A Function/Sub called as
                                'a statement keeps eax (its call surfaces in the following
                                'HRESULT check, e.g. If field_34.LoadBitmap(..) = 0).
                                If Len(umKind) = 0 Then NativeTryMethodKind umAddr, umKind
                                If InStr(umKind, "Get") > 0 Then
                                    NVReg(0) = ""
                                Else
                                    'A Function called as a statement: keep eax for a
                                    'following HRESULT check (`If obj.Method() = 0`), but
                                    'ALSO arm a deferred Call so it EMITS `Call obj.Method(args)`
                                    'when the result is NOT consumed - otherwise the resolved
                                    'call silently dropped (e.g. global_0055D5A0.Send_Data(..)
                                    'before a Me.Hide).  The main loop clears the pending call
                                    'if eax IS consumed, so the HRESULT-check inline (clsBitmap
                                    'If field_34.LoadBitmap(..)=0) is unchanged.
                                    NVPendingCall = "Call " & umVal
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
                        'Resolve on the form's own vtable (Me) OR on the standalone VB
                        '_Global object held in a module global (tagged at __vbaNew2) -
                        'both expose the App/Screen/Clipboard accessors at these offsets.
                        If NVRegIsFormVt(gcb) Or NVRegObjVt(gcb) = "_Global" Then
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
                'A VB6 global STATEMENT routed through the _Global object's vtable
                '(e.g. `Unload <form>` = call [_GlobalVt + 0x10]).  Render the global
                'statement with its argument (the object being acted on), taken from the
                'push below the implicit `this` (the _Global object, pushed topmost).
                Dim gmeth As String
                If ocb >= 0 And ocb <= 7 Then
                    If NVRegObjVt(ocb) = "_Global" Then
                        gmeth = NativeGlobalMethodByOffset(disp)
                        If Len(gmeth) > 0 Then
                            Dim gmArg As String, gmFm As String
                            If NVPushTop >= 2 Then gmArg = NVPushImm(NVPushTop - 2)
                            'The masked event-source param (arg_8) is Me; a form-instance
                            'access renders `New frmX` - strip the New for the Unload arg.
                            If gmArg = "arg_8" Then gmArg = "Me"
                            If Left$(gmArg, 4) = "New " Then gmArg = Mid$(gmArg, 5)
                            'A predeclared form-instance global -> the form name
                            '(Unload frmMainMenu, not Unload global_00423108).
                            If Left$(gmArg, 7) = "global_" Then
                                gmFm = FormNameByInstGlobal(NativeGlobalTokVa(gmArg))
                                If Len(gmFm) > 0 Then gmArg = gmFm
                            End If
                            NVPushTop = 0: NativeResetValue
                            If Len(gmArg) > 0 Then
                                NativeProcessInst = ind & gmeth & " " & gmArg & vbCrLf
                            Else
                                NativeProcessInst = ind & gmeth & vbCrLf
                            End If
                            Exit Function
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
                    'Require the receiver to be the GENUINE Me: the NVRegIsMe heuristic also
                    'tags an abs-global form instance (`mov esi,[0x423108]` = frmMainMenu),
                    'so a method ALSO callable on another form (Show) would mis-render as
                    '`Me.Show` for `frmMainMenu.Show vbModal,Me` (wrong receiver + dropped
                    'args).  The `this` is the topmost push; only fire when it is Me (arg_8),
                    'otherwise fall through to the UnkVCall path (honest, correct receiver).
                    Dim fmThis As String
                    If NVPushTop >= 1 Then fmThis = NVPushImm(NVPushTop - 1)
                    If NVRegIsFormVt(ocb) And (fmThis = "arg_8" Or fmThis = "Me") Then
                        fmeth = NativeFormMethodByOffset(disp)
                        If Len(fmeth) > 0 Then
                            'An arg-taking _Form method (PopupMenu <menu>) leads with the
                            'implicit `this` (Me, pushed topmost) and the real argument just
                            'below it - the same shape as the _Global Load/Unload statements.
                            Dim fmArg As String
                            If NativeFormMethodHasArg(disp) And NVPushTop >= 2 Then fmArg = NVPushImm(NVPushTop - 2)
                            NVPushTop = 0
                            NVReg(0) = "": NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                            NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjGuid(0) = "": NVRegObjVtGuid(0) = ""
                            If Len(fmArg) > 0 Then
                                NativeProcessInst = ind & fmeth & " " & fmArg & vbCrLf
                            Else
                                NativeProcessInst = ind & "Me." & fmeth & vbCrLf
                            End If
                            Exit Function
                        End If
                    End If
                End If
                'A built-in Form method called on ANOTHER form's predeclared instance
                '(`frmAbout.Show`): the `this` is a form-instance global, not Me, so the
                'Me path above (correctly) skipped it.  Resolve the form name from the
                'global (gFormInstGlobal) and render `<form>.<method>` - a strong, specific
                'signal (the global maps to a real form), so it never mis-fires.
                If ocb >= 0 And ocb <= 7 And Len(fmeth) = 0 Then
                    Dim fmgThis As String, fmgVa As Long, fmgForm As String
                    If NVPushTop >= 1 Then fmgThis = NVPushImm(NVPushTop - 1)
                    If Left$(fmgThis, 7) = "global_" Then
                        fmgVa = NativeGlobalTokVa(fmgThis)
                        If fmgVa <> 0 Then fmgForm = FormNameByInstGlobal(fmgVa)
                    End If
                    If Len(fmgForm) > 0 Then
                        fmeth = NativeFormMethodByOffset(disp)
                        If Len(fmeth) > 0 Then
                            Dim fmgArg As String
                            If NativeFormMethodHasArg(disp) And NVPushTop >= 2 Then fmgArg = NVPushImm(NVPushTop - 2)
                            NVPushTop = 0
                            NVReg(0) = "": NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                            NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjGuid(0) = "": NVRegObjVtGuid(0) = ""
                            If Len(fmgArg) > 0 Then
                                NativeProcessInst = ind & fmgForm & "." & fmeth & " " & fmgArg & vbCrLf
                            Else
                                NativeProcessInst = ind & fmgForm & "." & fmeth & vbCrLf
                            End If
                            Exit Function
                        End If
                    End If
                End If
                'A form calling its own method: call [vtable + 0x6F8 + slot*4].
                'Checked first: requiring a real gFormVtable slot is a stronger
                'signal than the NVBase control heuristic (which the same form-method
                'calls can otherwise mis-solve into a bogus control base).
                Dim ftgt As Long, fvThis As String
                If NVPushTop >= 1 Then fvThis = NVPushImm(NVPushTop - 1)
                ftgt = NativeFormVtableTarget(disp)
                'This path resolves against NVForm's OWN vtable, so it must not fire for a
                'call on a DIFFERENT form instance (`mov edi,[global_X]; call [edi_vt+off]`
                'where global_X is e.g. frmClient) - that would mis-resolve to NVForm's
                'same-offset method and leak the real receiver into the args (the
                '`frmCreate.cmdOK_Click(global_0055D5A0, ...)` bug).  A genuine self-call
                'pushes arg_8/Me (or leaves `this` untracked), OR uses the form's OWN
                'predeclared-instance global (global_X == NVForm) - both fine; only a
                'global_ that maps to a DIFFERENT form is excluded.
                'Exclude a call whose `this` is a separate object held in a module global
                '(`mov edi,[global_X]; call [edi_vt + off]`): such a call is NOT a Me-self
                'call, so resolving it against NVForm's vtable is wrong - it fabricates
                'NVForm's same-offset method and leaks the real receiver into the args
                '(client2 `frmCreate.cmdOK_Click(global_0055D5A0, ...)`; Dungeon Form_Load
                'mis-attributing a SOUND-object call to frmMain.Update_Status/Death).  A
                'genuine self-call pushes arg_8/Me (or leaves `this` untracked), never a
                'global_.  Verified against source: both were plausible-but-wrong.
                If ftgt <> 0 And Left$(fvThis, 7) <> "global_" Then
                    'A COM-style form-method call leads with the implicit Me/this (the
                    'form pointer, tracked as arg_8 when it reaches the arg stack); drop
                    'it and the trailing retbuf, keeping the real parameters (count from
                    'the method's typeinfo signature) - so `Update_Status()` no longer
                    'renders as frmMain.Update_Status(arg_8).
                    Dim fpname As String, fargs As String, fRetbuf As String, fkind As String, fIsVal As Boolean, fVal As String
                    fargs = NativeDropThisArgs(NativeArgList(), ftgt, fRetbuf)
                    fpname = NativeCallTargetName(ftgt)
                    NVPushTop = 0
                    'Emit DIRECTLY (not deferred), exactly like the class-vtable path
                    'below: a COM/form method returns its result through a hidden retbuf
                    'and leaves an HRESULT in eax that VB error-checks (cmp/test eax)
                    'on the very next instruction.  A deferred call would be folded into
                    'that check and LOST - which left Timer2_Timer / vsCarry_Change with
                    'empty bodies (their only statements are such form-method calls).
                    'A value-returning method (retbuf present, or a Property Get) folds
                    'recv.method(args) into the retbuf local + eax so the value flows to
                    'its consumer; a void method emits a `Call` statement.
                    fIsVal = (Len(fRetbuf) > 0)
                    If Not fIsVal Then
                        If NativeTryMethodKind(ftgt, fkind) Then fIsVal = (InStr(fkind, "Get") > 0)
                    End If
                    If fIsVal Then
                        fVal = NVForm & "." & fpname
                        If Len(fargs) > 0 Then fVal = fVal & "(" & fargs & ")"
                        NativeResetValue
                        NVReg(0) = fVal
                        NVRegObjType(0) = "": NVRegObjVt(0) = ""
                        If Left$(fRetbuf, 4) = "var_" Then
                            On Error Resume Next
                            NativeSetLocalExpr -CLng("&H" & Mid$(fRetbuf, 5)), fVal
                            On Error GoTo 0
                        End If
                        Exit Function
                    End If
                    NativeResetValue
                    NativeProcessInst = ind & "Call " & NVForm & "." & fpname & "(" & fargs & ")" & vbCrLf
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
                    'A property call on the form's OWN _Form vtable: `Form1.Caption =
                    '"Hey"` compiles to `mov this,[formInstance]; push value; push this;
                    'call [this_vt + off]` (Caption Let = 0x54).  When we are inside a
                    'form and the offset resolves to a real _Form property (GetProperty
                    'returns "" for any non-_Form offset, so unrelated calls fall
                    'through), render it.  Done BEFORE clearing the push stack so the
                    'Let value (a pushed arg) is recovered.
                    'Gate tightly so an UNTRACKED control's property call (e.g.
                    'lblSkillName.Caption, also vtoffset 0x54) is NOT mislabelled as
                    'the form's: the receiver must be the form's own vtable (Me) or a
                    'predeclared form-instance global (rendered global_XXXX) - a
                    'control receiver is a `frmX.ctl`/var_ token, never a bare global_.
                    If disp >= &H40 And disp < &H250 Then
                        Dim formObj As String, thisTok As String, gva As Long
                        'Me's own vtable -> the current form.
                        If ocb >= 0 And ocb <= 7 Then
                            If NVRegIsFormVt(ocb) And Len(NVForm) > 0 Then formObj = NVForm
                        End If
                        'A predeclared form-instance global -> the mapped form (works
                        'from a .bas module too); else the current form if in one.
                        If Len(formObj) = 0 And NVPushTop >= 1 Then
                            thisTok = NVPushImm(NVPushTop - 1)
                            If Left$(thisTok, 7) = "global_" Then
                                gva = NativeGlobalTokVa(thisTok)
                                formObj = FormNameByInstGlobal(gva)
                                If Len(formObj) = 0 And Len(NVForm) > 0 Then formObj = NVForm
                            End If
                        End If
                        If Len(formObj) > 0 Then
                            Dim fpv As String
                            'Pass the Form EVENT/coclass GUID - GetProperty matches it
                            'then searches the NEXT typeinfo (the _Form property iface).
                            fpv = NativeControlProp(formObj, "{33AD4F38-6699-11CF-B70C-00AA0060D393}", disp)
                            If Len(fpv) > 0 Then
                                NVPushTop = 0
                                NativeProcessInst = ind & fpv & vbCrLf
                                Exit Function
                            End If
                        End If
                    End If
                    'A property vtable call that the per-register control tracking
                    'above did NOT resolve.  The old single-slot NVLastControl is
                    'stale-prone (it names whatever control was fetched last, not the
                    'one this call is on), so leave the call raw rather than emit a
                    'wrong Form.Control.Prop.  Correct resolutions come only from the
                    'tracked-vtable path (NativeControlProp).
                    'Drop the COM lifetime/identity plumbing VB6 emits around every
                    'object use - QueryInterface / AddRef / Release at the IUnknown
                    'vtable slots 0 / 4 / 8 - which is pure noise, never user code.
                    If disp = 0 Or disp = 4 Or disp = 8 Then NVPushTop = 0: Exit Function
                    'An unresolved object method/property vtable call (the target is not
                    'in a loaded typelib): render it as `<obj>.UnkVCall_<8hex>h(args)` -
                    'matching the commercial decompiler - instead of dropping it as an
                    'annotated-assembly comment, so the call site stays traceable.  The
                    'COM calling convention pushes the implicit `this` last (the topmost
                    'push), so it is the receiver and the pushes below it are the
                    'arguments (a `this` whose register is an untracked deref renders as
                    'a leading-dot `.UnkVCall`, exactly as the commercial also emits).
                    'Emitted DIRECTLY (not deferred): a COM call returns via a hidden
                    'by-ref buffer and leaves an HRESULT in eax that VB immediately
                    'error-checks, which would otherwise consume a deferred call.
                    NativeResetValue
                    NativeProcessInst = ind & NativeUnkVCall(disp) & vbCrLf
                    Exit Function
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
                    'A module Function returning a Variant/String/UDT via a hidden retbuf:
                    'its FIRST arg (the topmost push, rendered first by NativeArgList) is
                    'the return slot - emit `<dest> = proc(<rest>)` instead of a Call that
                    'leaks the retbuf local as an argument + a raw `= eax` at the consumer.
                    Dim rbN As String, rbDest As String, rbRest As String, rbCp As Long
                    rbN = NativeColGet(gRetbufFunc, "V" & tgt)
                    If Len(rbN) > 0 Then
                        rbCp = InStr(uargs, ", ")
                        If rbCp > 0 Then
                            rbDest = Left$(uargs, rbCp - 1): rbRest = Mid$(uargs, rbCp + 2)
                        Else
                            rbDest = uargs: rbRest = ""
                        End If
                        If Left$(rbDest, 4) = "var_" Or Left$(rbDest, 4) = "arg_" Or Left$(rbDest, 7) = "global_" Then
                            NVReg(0) = "": NVRegIsAddr(0) = False: NVRegObjType(0) = "": NVRegObjVt(0) = ""
                            NativeProcessInst = ind & rbDest & " = " & pname & "(" & rbRest & ")" & vbCrLf
                            Exit Function
                        End If
                    End If
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
            'Variant For loop header (detected pre-pass): this exit Jcc is a `For`.
            Dim vffStr As String
            vffStr = NativeColGet(NVVarForFor, "W" & inst.va)
            If Len(vffStr) > 0 Then
                Dim vffP() As String, vffC As String, vffS As String, vffL As String
                vffP = Split(vffStr, "|")
                vffC = vffP(0)
                vffS = "?": vffL = "?"
                If UBound(vffP) >= 1 And Len(vffP(1)) > 0 Then vffS = vffP(1)
                If UBound(vffP) >= 2 And Len(vffP(2)) > 0 Then vffL = vffP(2)
                NativeProcessInst = ind & "For " & vffC & " = " & vffS & " To " & vffL & vbCrLf
                NVIndent = NVIndent + 1
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
            'Back-edge of a Variant For loop -> Next <counter>.
            Dim vfnStr As String
            vfnStr = NativeColGet(NVVarForNext, "W" & inst.va)
            If Len(vfnStr) > 0 Then
                If NVIndent > 0 Then NVIndent = NVIndent - 1
                NativeProcessInst = NativeIndentStr() & "Next " & vfnStr & vbCrLf
                Exit Function
            End If
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
            'Exit <kind> matching the proc (Function/Property/Sub) so a recovered
            'class Function/Property Get emits `Exit Function`/`Exit Property`, not a
            'mis-typed `Exit Sub` (NVProcEndWord is set by NativeProcHeader).
            NativeProcessInst = ind & "Exit " & NVProcEndWord & vbCrLf
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
                Dim ld As Long, lAbs As Boolean, llBase As Long
                If NativeDecodeDisp(inst.dump, ld, lAbs) Then
                    NVLastLea = ld: NVLastLeaSet = True
                    'A `lea reg,[Me + off]` (positive disp on the Me register, no index)
                    'takes the address of an instance FIELD - a following move/copy into
                    'it is a field store (field_<off> = src), unlike the common
                    '`lea reg,[ebp-X]` which addresses a local.  (Guard the base-register
                    'index before NVRegIsMe - VB6 And is not short-circuit.)
                    NVLastLeaField = False
                    llBase = NativeMemBase(inst.dump)
                    If Not lAbs And ld > 0 And llBase >= 0 And llBase <= 7 And NativeMemIndex(inst.dump) < 0 Then
                        NVLastLeaField = NVRegIsMe(llBase)
                    End If
                End If
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

Private Function NativeIsFoldableArith(ByVal v As String) As Boolean
    'A value worth extending with `+ N` / `- N`:
    '  (1) a FUNCTION-CALL result (Len(s), InStr, Asc) where the immediate is a real
    '      computation, so `Right$(s, Len(s) - 4)` keeps the "- 4"; OR
    '  (2) a fully-dereferenced ARRAY-ELEMENT VALUE that carries a recovered VARIABLE
    '      index group (global_X(12)(arg_8)(252), arr(var_C)(4) - a UDT field read),
    '      so a read-modify-write `Stamina = Stamina - 1` keeps its "- 1".
    'Deliberately EXCLUDES a bare deref/pointer (global_X(12) / arg_8(96)) whose only
    'paren groups are NUMERIC, and a bare local (var_38, often a loop counter): those
    'are address/index math, and folding onto them rendered misleading `(arg_8(96) +
    '4)` / off-by-one `(var_38 + 1)`.  A recovered variable index `(arg_/var_/global_)`
    'is the tell that distinguishes a genuine element VALUE from that pointer math.
    If Len(v) = 0 Then Exit Function
    If Left$(v, 1) = Chr$(34) Then Exit Function                    'a string literal
    'A recovered array-element value is a CLEAN DEREF CHAIN that carries a variable
    'index/arg group (global_X(12)(arg_8)(252)).  It must NOT start with "(" nor
    'contain a top-level operator - else it is a parenthesised arithmetic / relational
    'expression (`(arg_C >= arg_14)`, `(var_38 + var_24)`) that merely begins with a
    'variable, where appending `+ N` is nonsense (folding `+4` onto a Boolean).
    If Left$(v, 1) <> "(" And Not NativeHasArithOp(v) Then
        If InStr(v, "(arg_") > 0 Or InStr(v, "(var_") > 0 Or InStr(v, "(global_") > 0 Then
            NativeIsFoldableArith = True: Exit Function
        End If
    End If
    If InStr(v, "(") < 2 Then Exit Function                         'need a name before "("
    'reject synthetic-local derefs (those are pointer/field math, not a call result)
    If Left$(v, 4) = "var_" Or Left$(v, 4) = "arg_" Or Left$(v, 4) = "loc_" _
       Or Left$(v, 7) = "global_" Or Left$(v, 6) = "field_" Then Exit Function
    NativeIsFoldableArith = True
End Function

Private Function NativeIs16Foldable(ByVal v As String) As Boolean
    'A 16-bit shadow value (NVR16Val) is always an Integer VALUE - a field/local/element
    'read or an Integer expression already built from one - never a pointer (16-bit
    'registers do not hold addresses in this code).  So arithmetic may fold onto any
    'non-empty, non-placeholder value, including a field-read deref arg_8(52) that the
    'pointer-aware NativeIsFoldableArith deliberately rejects.  Length-capped so repeated
    'folds cannot blow up.
    If Len(v) = 0 Then Exit Function
    If Left$(v, 1) = Chr$(34) Then Exit Function       'string literal
    If Left$(v, 1) = "<" Then Exit Function            '<arg>/<cond> placeholder
    If Len(v) > 70 Then Exit Function
    NativeIs16Foldable = True
End Function

Private Function NativeHasArithOp(ByVal v As String) As Boolean
    'True when v contains a rendered (space-delimited) arithmetic / relational / logical
    'operator - i.e. it is an EXPRESSION, not a clean deref-chain value.  A deref chain
    'like global_X(12)(arg_8)(252) has no spaces, so this is False for it.
    NativeHasArithOp = (InStr(v, " + ") > 0 Or InStr(v, " - ") > 0 Or InStr(v, " * ") > 0 _
        Or InStr(v, " / ") > 0 Or InStr(v, " \ ") > 0 Or InStr(v, " Mod ") > 0 _
        Or InStr(v, " >= ") > 0 Or InStr(v, " <= ") > 0 Or InStr(v, " <> ") > 0 _
        Or InStr(v, " > ") > 0 Or InStr(v, " < ") > 0 Or InStr(v, " = ") > 0 _
        Or InStr(v, " And ") > 0 Or InStr(v, " Or ") > 0 Or InStr(v, " Is ") > 0)
End Function

Private Function NativeFoldArith(ByVal val As String, ByVal isSub As Boolean, ByVal imm As Long) As String
    'Render (val - imm) / (val + imm), normalising a negative immediate (so `+ -4`
    'reads `- 4`) and capping length so repeated folds cannot blow up.
    Dim op As String
    NativeFoldArith = val
    If imm = 0 Then Exit Function                                   'no-op add/sub 0
    If isSub Then op = " - " Else op = " + "
    If imm < 0 Then
        imm = -imm
        If isSub Then op = " + " Else op = " - "
    End If
    If Len(val) > 70 Then Exit Function                             'avoid runaway nesting
    NativeFoldArith = "(" & val & op & CStr(imm) & ")"
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

Private Function NativeResolveVarArg(ByVal arg As String) As String
    'A by-ref Variant temp local (var_X) built for an untyped (Variant) parameter, when
    'it holds a SIMPLE VALUE type, renders the literal the caller passed instead of the
    'temp name: a VARIANT at slot -X has its VT tag at -X and its data at -X+8, so
    '`MOV [ebp-34],2 ; MOV [ebp-2C],55` is the Integer 55 -> SetupCalc(55).  Restricted to
    'value VTs (I2/I4/R4/R8/CY/DATE); a BSTR/object Variant's +8 field is a pointer, not
    'the value, so those keep the temp name.
    NativeResolveVarArg = arg
    If Left$(arg, 4) <> "var_" Then Exit Function
    Dim disp As Long, vt As String, dataV As String, vtN As Long
    On Error Resume Next
    disp = -CLng("&H" & Mid$(arg, 5))
    On Error GoTo 0
    If disp >= 0 Then Exit Function
    vt = NativeGetVSlot(disp)
    If Len(vt) = 0 Or Not IsNumeric(vt) Then Exit Function
    vtN = CLng(vt)
    If vtN < 2 Or vtN > 7 Then Exit Function          '2=I2 3=I4 4=R4 5=R8 6=CY 7=DATE
    dataV = NativeGetVSlot(disp + 8)                   'VARIANT data field
    If Len(dataV) > 0 Then NativeResolveVarArg = dataV
End Function

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
    'OCX control? Resolve via its OWN typelib first (authoritative), overriding the
    'VB6.OLB guess - see NativeControlProp.  NVLastControl is the receiver (e.g.
    '"frmCreate.sldrLife"); a tControl-array OCX carries a VB-intrinsic GUID here, so
    'GetProperty would mis-name its members (Slider .Value -> .ClientHeight).
    Dim ocxLib As String, ocxBase As String, ocxInv As Long, ocxName As String, dotP As Long
    ocxLib = GetControlClass(NVLastControl)
    If Len(ocxLib) > 0 Then
        ocxBase = NVLastControl
        dotP = InStrRev(ocxBase, ".")
        If dotP > 0 Then ocxBase = Mid$(ocxBase, dotP + 1)
        ocxName = modCOM.VtableMemberName(ocxLib, ocxBase, vtOffset, 0, ocxInv)
        If Len(ocxName) > 0 Then p = "X_" & ocxName & " (" & cTypeInfo.InvKind2String(ocxInv) & ")"
    End If
    If Len(p) = 0 Then
        p = modPCode.GetProperty(NVLastGuid, vtOffset)
        If InStr(p, "Unknown GUID") > 0 Then Exit Function   'GUID/offset not in a loaded TypeLib - leave raw
    End If
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
    'When the control has a known EXTERNAL class (an OCX: MSCOMCTL Slider/StatusBar,
    'RichTextBox, ...), its OWN typelib is the authoritative source for the vtable
    'offset -> member mapping.  Resolve there FIRST and override any VB6.OLB guess: a
    'named OCX control already present in the tControl array carries a VB-intrinsic GUID,
    'so GetProperty would mis-resolve its members (e.g. Slider .Value -> .ClientHeight).
    'Intrinsic controls have no external class (GetControlClass empty), so this never
    'perturbs them.  Falls through to GetProperty/VB6.OLB when the OCX doesn't resolve.
    Dim ocxLib As String, ocxBase As String, ocxInv As Long, ocxName As String, dotP As Long
    ocxLib = GetControlClass(ctlName)
    If Len(ocxLib) > 0 Then
        ocxBase = ctlName
        dotP = InStrRev(ocxBase, ".")
        If dotP > 0 Then ocxBase = Mid$(ocxBase, dotP + 1)
        ocxName = modCOM.VtableMemberName(ocxLib, ocxBase, vtOffset, 0, ocxInv)
        If Len(ocxName) > 0 Then p = "X_" & ocxName & " (" & cTypeInfo.InvKind2String(ocxInv) & ")"
    End If
    If Len(p) = 0 Then
        p = modPCode.GetProperty(guid, vtOffset)
        If InStr(p, "Unknown GUID") > 0 Then Exit Function   'GUID/offset not in any loaded TypeLib - leave raw
    End If
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

    'Record (UDT) helpers carry the record-layout descriptor address as one of their
    'arguments.  Harvest it (decode -> gUDTDesc) so the Type block is reconstructed.
    'Non-destructive: the call still renders normally below (type recovery only).
    If InStr(nm, "__vbaRec") > 0 Then NativeHarvestUDTArgs

    'Variant For loop: a detected __vbaVarForInit becomes the `For` header (its args are
    'the counter / start / limit), and its paired __vbaVarForNext becomes `Next`; both
    'calls are suppressed.  The structure was paired in NativeDetectWhileLoops.
    If InStr(nm, "__vbaVarForInit") > 0 Then
        Dim vfLink As String
        vfLink = NativeColGet(NVVarForInitLink, "V" & inst.va)
        If Len(vfLink) > 0 Then
            Dim vfA() As String, vfArgs As String, vfL() As String
            Dim vfCtr As String, vfStart As String, vfLimit As String
            vfArgs = NativeArgList()
            vfA = Split(vfArgs, ", ")
            If UBound(vfA) >= 0 Then vfCtr = vfA(0)
            If UBound(vfA) >= 1 Then vfStart = vfA(1)
            If UBound(vfA) >= 2 Then vfLimit = vfA(2)
            vfL = Split(vfLink, "|")                 'jccVA|backedgeVA
            NativeColPut NVVarForFor, "W" & vfL(0), vfCtr & "|" & vfStart & "|" & vfLimit
            NativeColPut NVVarForNext, "W" & vfL(1), vfCtr
            NVPushTop = 0
            NativeRuntimeCall = "": Exit Function
        End If
    ElseIf InStr(nm, "__vbaVarForNext") > 0 Then
        If Len(NativeColGet(NVVarForSuppress, "V" & inst.va)) > 0 Then
            NVPushTop = 0: NativeRuntimeCall = "": Exit Function
        End If
    End If

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
            Dim osGuid As String, osName As String, osSrc As String
            osGuid = NVRegObjGuid(0): osName = NVRegObjType(0)
            'The source control's identity is often cleared off eax by a `lea eax,
            '[dest]` placed just before the call; fall back to the last control
            'accessor (the __vbaObjSet source is always the just-fetched control).
            If Len(osGuid) = 0 And Len(NVLastGuid) > 0 Then osGuid = NVLastGuid: osName = NVLastControl
            '__vbaObjSet(ppDest, pSrc) RETURNS pSrc.  pSrc is the source object - the
            'bottom (first-pushed) argument; ppDest (the &local) is pushed last.  Capture
            'it so eax becomes the real source even when no control identity is on eax
            '(e.g. `Unload <form>`: the form was pushed from a register, leaving a stale
            'tag on eax that otherwise leaked as the Unload argument).
            If NVPushTop >= 1 Then osSrc = NVPushImm(0)
            NVPushTop = 0
            'Re-tag eax with the control identity: __vbaObjSet returns the same
            'object, so a following `mov [tempLocal], eax` (the property LET target)
            'keeps the GUID and the LET through that local resolves.
            If Len(osGuid) > 0 Then
                NVReg(0) = osName: NVRegObjType(0) = osName: NVRegObjGuid(0) = osGuid
                NVRegObjVt(0) = "": NVRegObjVtGuid(0) = ""
            ElseIf Len(osSrc) > 0 Then
                'No control identity: eax is simply the source object the helper returns.
                NVReg(0) = osSrc
                NVRegObjType(0) = "": NVRegObjGuid(0) = "": NVRegObjVt(0) = "": NVRegObjVtGuid(0) = ""
                NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
            End If
            If NVLastLeaSet And Len(osGuid) > 0 Then
                NativeSetLocalGuid NVLastLea, osGuid
                NativeSetLocalExpr NVLastLea, osName
                NVLastLeaSet = False
                'A `Set temp = control` that exists only to pass the control into the
                'immediately-following late call is redundant (the late call renders the
                'control directly) - drop it, keeping the re-tag above so the call resolves.
                If Len(NativeColGet(NVSuppressObjSet, "O" & inst.va)) > 0 Then
                    NativeRuntimeCall = "": Exit Function
                End If
                NativeRuntimeCall = "Set var_" & Hex$(Abs(NVLastLea)) & " = " & osName
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
        Case InStr(nm, "__vbaI4Abs") > 0, InStr(nm, "__vbaI2Abs") > 0
            'Integer/Long Abs(): the value is passed in ECX (not pushed), result in eax.
            'Fold to `Abs(<ecx>)` ONLY when ecx holds a freshly-computed arithmetic
            'expression (a difference like X1-X2) - `modMap_Dist = Abs(X1-X2)` etc.  A
            'bare/stale register or leftover variable is NOT folded: in Open_Sight_Line
            'the first Abs's ecx is a control-flow-merged remnant (stale), so folding it
            'would emit a wrong Abs (verified) - leave those as a visible Call.
            Dim absArg As String, absInner As String, absSp As Long, absParamDiff As Boolean
            absArg = NVReg(1)
            'Skip a difference of two bare PARAMETERS (Abs(arg_8 - arg_10)): that shape is
            'characteristic of the tiny distance/max helpers (modMap_Dist), where the two
            'Abs results feed an `If X > Y Then ret = X Else ret = Y` the linear model
            'can't reconstruct (both branches share eax -> both returns come out identical
            'and one is wrong).  An Abs of data (global/array/local) feeds a guard condition
            'instead (frmMain Display) and reconstructs correctly - fold those.
            absInner = absArg
            If Left$(absInner, 1) = "(" And Right$(absInner, 1) = ")" Then absInner = Mid$(absInner, 2, Len(absInner) - 2)
            absSp = InStr(absInner, " - "): If absSp = 0 Then absSp = InStr(absInner, " + ")
            If absSp > 0 Then
                absParamDiff = (Left$(Trim$(Left$(absInner, absSp - 1)), 4) = "arg_") _
                           And (Left$(Trim$(Mid$(absInner, absSp + 3)), 4) = "arg_")
            End If
            If Len(absArg) > 0 And Left$(absArg, 1) = "(" _
               And (InStr(absArg, " - ") > 0 Or InStr(absArg, " + ") > 0) _
               And Not absParamDiff Then
                NVReg(0) = "Abs" & absArg
                NVRegIsAddr(0) = False: NVRegObjType(0) = "": NVRegObjVt(0) = ""
                NativeRuntimeCall = "": Exit Function
            Else
                NativeRuntimeCall = "Call " & NativeFriendlyName(nm) & "()": Exit Function
            End If
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
        Case InStr(nm, "__vbaI2Var") > 0
            'Variant -> Integer (result in ax).  Folds the narrowing of a Variant result
            'into eax for the consumer, e.g. `Item(i).ID = CInt(var_40)` where var_40 holds
            'a Variant-returning Function's result (modMap_Fix_Walls); was a dropped
            '`Call I2Var(var_40)` + a raw `= eax` store.
            aa = NativeArgPop()
            If Len(aa) = 0 Then aa = "<arg>"
            NVReg(0) = "CInt(" & aa & ")"
            NVRegIsAddr(0) = False: NVRegObjType(0) = "": NVRegObjVt(0) = ""
            NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaI4Var") > 0
            'Variant -> Long (result in eax).
            aa = NativeArgPop()
            If Len(aa) = 0 Then aa = "<arg>"
            NVReg(0) = "CLng(" & aa & ")"
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
            'fold it through to the consumer (e.g. an API argument).  Pop ONLY the
            'conversion's own two args and KEEP the rest of the push stack - those are
            'the surrounding API call's earlier arguments (e.g. LoadImage's uType / cx
            '/ cy / fuLoad pushed before the string), which clearing here truncated.
            Dim adst As String, asrc As String
            adst = NativeArgPop(): asrc = NativeArgPop()
            If Len(asrc) = 0 Then asrc = adst
            NVReg(0) = asrc: NVKeepPushStack = True: NativeRuntimeCall = "": Exit Function
        Case InStr(nm, "__vbaStrCmp") > 0, InStr(nm, "__vbaStrComp") > 0, InStr(nm, "__vbaStrTextCmp") > 0
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
            'DIRECT-TEST form (no neg/sbb/inc/neg materialisation): `call __vbaStrCmp;
            'test eax,eax; jcc` - the raw tri-state is tested directly (0 = equal).  Hand
            'the operands to the following `test eax,eax`, which then behaves like
            '`cmp p1,p2` (identical Jcc polarity: je=equal, jl=p1<p2, ...), so the Jcc
            'renders `If var_124 = "!"` etc. (a char-dispatch chain in the map parser).
            If Len(NativeColGet(NVStrCmpDirect, "P" & inst.va)) > 0 And scN >= 2 Then
                If NativeIsStrOperand(scA(0)) And NativeIsStrOperand(scA(1)) Then
                    NVStrCmpP1 = scA(1): NVStrCmpP2 = scA(0): NVStrCmpPending = True
                    NVReg(0) = ""
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
        Case InStr(nm, "__vbaVarTstNe") > 0, InStr(nm, "__vbaVarTstEq") > 0
            'Variant comparison: __vbaVarTstNe/Eq(a, b) returns a<>b / a=b (a VARIANT_BOOL
            'in ax).  Bind the relational into eax so the following test/jcc renders
            'If (a <> b).  Symmetric ops only (order-independent) - Lt/Gt would need the
            'operand order pinned.
            Dim vtA() As String, vtN As Long, vtOp As String
            NativeArgsSnapshot vtA, vtN
            If vtN >= 2 Then
                vtOp = IIf(InStr(nm, "TstNe") > 0, "<>", "=")
                If Len(vtA(0)) > 0 And vtA(0) <> "<arg>" And Len(vtA(1)) > 0 And vtA(1) <> "<arg>" Then
                    NVReg(0) = "(" & vtA(0) & " " & vtOp & " " & vtA(1) & ")"
                    NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                    NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjGuid(0) = "": NVRegObjVtGuid(0) = ""
                    NativeRuntimeCall = "": Exit Function
                End If
            End If
            Dim vtList As String, vtK As Long
            For vtK = 0 To vtN - 1
                If Len(vtList) > 0 Then vtList = vtList & ", "
                vtList = vtList & vtA(vtK)
            Next
            NativeRuntimeCall = "Call " & NativeFriendlyName(nm) & "(" & vtList & ")": Exit Function
        Case InStr(nm, "__vbaI2I4") > 0, InStr(nm, "__vbaUI1I2") > 0, _
             InStr(nm, "__vbaUI1I4") > 0, InStr(nm, "__vbaI4UI1") > 0, _
             InStr(nm, "__vbaI2UI1") > 0
            'Implicit integer widening/narrowing.  Usually the value already sits in
            'eax (NVReg(0)) and we just suppress the Call.  But VB6 also passes the
            'value in ECX (`mov ecx,1; call __vbaI2I4`, e.g. setting a small Integer
            'field); there eax is untracked, so fold the clean ecx value into eax -
            'else the consuming store leaks a raw `field = eax`.
            'A fresh `mov ecx, imm; call __vbaI2I4` (the Select-Case-on-Integer arm:
            'coerce the Long case-label to Integer for `cmp KeyCode, ax`) puts the
            'INPUT in ecx and returns it in eax - so the numeric literal in ecx is the
            'result regardless of any stale value eax held.  A clean NAMED ecx value is
            'only the input when eax is otherwise untracked (the original case).
            If NativeIsNumLit(NVReg(1)) Then
                NVReg(0) = NVReg(1)
            ElseIf Len(NVReg(0)) = 0 And NativeIsCleanNamedVal(NVReg(1)) Then
                NVReg(0) = NVReg(1)
            End If
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
            Dim ncls As String, ngObj As Long, rp As Long, rv As Long, isGlobalObj As Boolean
            For rp = 0 To 7
                rv = NVRecentPush(rp)
                If rv <> 0 And Len(ncls) = 0 Then ncls = NativeClassFromObjInfo(rv)
                If rv <> 0 And Not isGlobalObj Then isGlobalObj = NativeIsVBGlobalDesc(rv)
                If rv <> 0 And NativeIsGlobalAddr(rv) Then ngObj = rv
            Next
            'A __vbaNew2 into a LOCAL (`lea reg,[ebp-X]; push reg` as @dest) - type that
            'local with the class so its later vtable calls resolve to var_X.Member.
            Dim ndp As Long, ndDisp As Long
            ndDisp = 0
            If Len(ncls) > 0 Then
                For ndp = 0 To NVPushTop - 1
                    If NVPushDisp(ndp) < 0 Then ndDisp = NVPushDisp(ndp): Exit For
                Next
                If ndDisp < 0 Then NativeColPut NVLocalObjType, "D" & ndDisp, ncls
            End If
            NVPushTop = 0
            If Len(ncls) > 0 Then
                If ngObj <> 0 Then NativeColPut NVObjClass, "G" & ngObj, ncls
                If ndDisp < 0 And ngObj = 0 Then
                    'Creation into a LOCAL -> emit `Set var_X = New <class>` (commercial:
                    'Set var = vbaNew2(class)) - we were dropping it, leaving an untyped
                    'var_X whose members were the only trace.  The object goes to the local
                    'via the @dest ptr; eax is the HRESULT the next instr error-checks, so
                    'clear it.  (`Set var = New clsX` also lets the Dim inference type it.)
                    'An `As New` local re-instantiates (If Is Nothing Then Set..New) before
                    'EACH use - emit the creation only ONCE per local; suppress the repeats
                    'so the proc matches the single `Set` commercial shows, not 10 of them.
                    NVReg(0) = "": NVRegObjType(0) = "": NVRegObjVt(0) = ""
                    If Len(NativeColGet(NVNewEmitted, "N" & ndDisp)) > 0 Then
                        NativeRuntimeCall = ""                   'repeat guard - suppress
                    Else
                        NativeColPut NVNewEmitted, "N" & ndDisp, "1"
                        NativeRuntimeCall = "Set var_" & Hex$(Abs(ndDisp)) & " = New " & ncls
                    End If
                    Exit Function
                End If
                NVReg(0) = "New " & ncls: NVRegObjType(0) = ncls
            ElseIf isGlobalObj Then
                'The VB6 _Global intrinsic-objects holder, lazily New'd into a module
                'global.  Tag the global as "_Global" so a following `[global+0x14]`
                'deref-call resolves to .App / .Screen / .Clipboard exactly like the
                'form's own vtable does (NativeGlobalObjByOffset), and App.Path & co.
                'then resolve via the existing intrinsic-property path.
                If ngObj <> 0 Then NativeColPut NVObjClass, "G" & ngObj, "_Global"
                NVReg(0) = "_Global": NVRegObjType(0) = "_Global"
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
            'A dynamic array of a user RECORD has an element size (arg index 1) larger
            'than any primitive; recover a byte-buffer UDT from it (descriptor-less UDTs
            'have no __vbaRec* descriptor - this is the only size signal for them).  Name
            'it after the array variable's address (arg index 2) so it reads like the
            'descriptor UDTs (UDT_<va>), falling back to the ReDim call site.
            Dim rdArr As String
            If rdN > 2 Then rdArr = rdA(2)
            If rdN > 1 Then NativeRegisterUDTBySize rdA(1), rdArr, inst.va
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
            Case "EOF", "FREEFILE"
                'EOF feeds a `Do While Not EOF(f)` loop test, reconstructed by a
                'separate pre-pass that doesn't (yet) consume a folded predicate, so
                'folding would drop the call and leave the loop cond blank.
                'FreeFile is the file number; it is used in many places (the condition,
                'and every Open/Put/Get/Close `#<n>`) but often has NO single clean
                '`var = FreeFile()` store to anchor a variable name, so folding leaks
                'the expression into all of them (`Open ... As #FreeFile(...)`).
                'Keep them visible Calls until the loop pre-pass / file-number tracking
                'handle them.  (Simple-If Boolean predicates Is*/IsNumeric DO fold now -
                'the 16-bit `test ax,ax` of their VARIANT_BOOL result resolves to
                '`<pred> <> 0` via NativeIsCallExpr in NativeDecodeCompare.)
                NativeRuntimeCall = "Call " & vbName & "(" & NativeArgList() & ")": Exit Function
            Case Else
                'Value-returning intrinsic (Environ$/Command$/Now/Timer/Rnd/
                'QBColor/Format*/financial/date funcs/TypeName/...).  Fold the result
                'into eax AND defer a "Call X()" - exactly the string-transform path:
                'if the next instruction consumes eax the value threads into the
                'consumer (`sTemp = Environ$(...)`, `lColor = QBColor(4)`), otherwise the
                'deferred call is flushed as a statement, so an unconsumed result is
                'never silently dropped.  (The old "always emit a Call" lost every such
                'assignment; the even older "fold into eax with no flush" dropped
                'unconsumed DateAdd/DateValue/CDate - the NVPendingCall net fixes both.)
                Dim ceArgs As String
                ceArgs = NativeArgList()
                NVReg(0) = vbName & "(" & ceArgs & ")"
                NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjGuid(0) = "": NVRegObjVtGuid(0) = ""
                NVPendingCall = "Call " & NVReg(0)
                NativeRuntimeCall = "": Exit Function
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

    'Late-bound dispatch: __vbaLateIdCall(obj, DISPID, ...) -> obj.Member[(args)].
    'Resolve the DISPID to the member name via the control's OCX typelib; on failure
    'fall through to the generic Call below (current behaviour, no regression).
    If InStr(nm, "__vbaLateId") > 0 Then
        Dim lcRes As String, lcResolved As Boolean
        lcRes = NativeLateIdCall(inst, nm, lcResolved)
        If lcResolved Then NativeRuntimeCall = lcRes: Exit Function
    ElseIf InStr(nm, "__vbaLateMem") > 0 Then
        Dim lmRes As String, lmResolved As Boolean
        lmRes = NativeLateMemCall(inst, nm, lmResolved)
        If lmResolved Then NativeRuntimeCall = lmRes: Exit Function
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
    'A move/copy whose fastcall DEST register (ecx) holds a module-global ADDRESS
    '(`mov ecx, &global; call __vbaStrMove`) stores into that global - e.g.
    '`Module1.packetValue = CStr(value)`.  Use it instead of the last LEA, which is often
    'an earlier cleanup lea (the __vbaFreeVar of the call's Variant arg) and would
    'mis-target a local (var_34).  Checked first; the global dest is an unambiguous signal.
    If Left$(NVReg(1), 7) = "global_" And (NativeIsExprValue(src) Or NativeIsCleanNamedVal(src)) Then
        NativeMoveAssign = NVReg(1) & " = " & src
        NVReg(0) = NVReg(1)
        NVLastLeaSet = False: NVLastLeaField = False: NVPendingArg = "": NVKeepPushStack = True
        Exit Function
    End If
    'A move/copy into a Me-FIELD ([Me+off], NVLastLeaField) is a field store
    'field_<off> = src - e.g. a String Property Let `packetName = vData` copies the
    'value param into [Me+0x3C] via __vbaStrCopy.  The store target has no ebp local
    'slot, and the value is often a plain param/field (vData) rather than a folded
    'expression, so accept a named source (arg_/field_/global_/var_) here too.
    If NVLastLeaSet And NVLastLeaField Then
        If Not NativeIsExprValue(src) And NativeIsCleanNamedVal(NVReg(2)) Then src = NVReg(2)
        If NativeIsExprValue(src) Or NativeIsCleanNamedVal(src) Then
            dn = "field_" & Hex$(NVLastLea)
            NativeMoveAssign = dn & " = " & src
            NVReg(0) = dn                   'the helper returns the moved value in eax
            NativeRecordFieldType NVLastLea, "String"   'a strcopy store -> a String field
        End If
        NVLastLeaSet = False: NVLastLeaField = False: NVPendingArg = "": NVKeepPushStack = True
        Exit Function
    End If
    If NVLastLeaSet And NativeIsExprValue(src) Then
        dn = "var_" & Hex$(Abs(NVLastLea))
        NativeMoveAssign = dn & " = " & src
        NativeSetLocalExpr NVLastLea, dn
        NVReg(0) = dn                       'the helper returns the moved value in eax
    ElseIf NVLastLeaSet And NVLastLea < 0 And NativeIsCleanNamedVal(NVReg(2)) Then
        'A plain param/field/global copied into a local via the fastcall (edx) source -
        'typically a compiler temp that is immediately copied onward (a String Property
        'Let compiles `temp = vData` then `field = temp`).  Track the source value
        'silently so the downstream store folds to it (field_<off> = vData) instead of
        'referencing the bare temp; no noisy `var_X = vData` line is emitted.
        NativeSetLocalExpr NVLastLea, NVReg(2)
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
    ElseIf NativeOwnerIsStdForm(NVForm) Then
        'A standard FORM always lays its control accessors at 0x2F8 (verified across
        'binaries).  If 0x2F8 maps NOTHING in this proc, the proc only touches controls
        'missing from our list (e.g. an OCX like CommonDialog1 whose tControl is not in
        'the parsed control array, ControlCount undercounts it) - the per-proc solver
        'would then fabricate a base that maps the unknown offset onto a REAL control,
        'mis-naming it (CommonDialog1 -> cmdSkillRaise).  Keep 0x2F8 so the unknown
        'control stays unresolved rather than wrong.  Per-proc solving is reserved for
        'genuinely non-standard layouts (UserControls), handled below.
        NativeSolveControlBase = FORM_CONTROL_BASE
    Else
        NativeSolveControlBase = NativeSolveControlBasePerProc(col)
    End If
End Function

Private Function NativeOwnerIsStdForm(ByVal owner As String) As Boolean
    'True when `owner` is a standard VB Form (not a class / UserControl / module),
    'so its control vtable base is the fixed 0x2F8.  Object-type values match the
    'form-detection set used when the control array is parsed (frmMain.OpenVBExe).
    Dim i As Long, ot As Long
    On Error Resume Next
    For i = 0 To UBound(gObjectNameArray)
        If gObjectNameArray(i) = owner Then
            ot = gObject(i).ObjectType
            NativeOwnerIsStdForm = (ot = 98435 Or ot = 17926147 Or ot = 98467 Or ot = 98499)
            Exit Function
        End If
    Next
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
        nm = NativeParamName(parts(i))
        If Len(nm) > 0 And NativeIsIdent(nm) Then
            NVArgTok(NVArgN) = "arg_" & Hex$(base + i * 4)
            NVArgNm(NVArgN) = nm
            NVArgN = NVArgN + 1
        End If
    Next
End Sub

Private Function NativeParamName(ByVal p As String) As String
    'Extract the bare parameter NAME from one parameter declaration. Handles both a
    'plain reconstructed name ("strName") and a fully-typed event signature part
    '("ByVal Number As Integer" -> "Number"): strip leading modifiers, then take the
    'identifier before " As ".
    p = Trim$(p)
    If InStr(1, p, "ByVal ", vbTextCompare) = 1 Then p = Trim$(Mid$(p, 7))
    If InStr(1, p, "ByRef ", vbTextCompare) = 1 Then p = Trim$(Mid$(p, 7))
    If InStr(1, p, "Optional ", vbTextCompare) = 1 Then p = Trim$(Mid$(p, 10))
    If InStr(1, p, "ParamArray ", vbTextCompare) = 1 Then p = Trim$(Mid$(p, 12))
    Dim asP As Long
    asP = InStr(1, p, " As ", vbTextCompare)
    If asP > 0 Then p = Trim$(Left$(p, asP - 1))
    'Defensive: if anything still has spaces, keep the last token.
    If InStr(p, " ") > 0 Then p = Mid$(p, InStrRev(p, " ") + 1)
    NativeParamName = p
End Function

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

Private Sub NativeDetectAccumReturn(b() As Byte, ByVal addr As Long)
    'A standard-module (or otherwise unclassified) Function returns a SIMPLE value in the
    'accumulator: the epilogue copies it from the return local into ax/eax right before
    'the SEH teardown:  mov ax/eax,[ebp-retSlot] ; mov reg,[ebp-Z] ; mov fs:[0],reg.
    'Detect that to (a) mark the proc a Function, (b) recover its return TYPE - word=ax=
    'Integer, dword=eax=Long, (c) rename the return slot var_X to the proc name so its
    'assignments read `FuncName = value` (VB's implicit return) instead of var_X = value.
    'Unlike NativeDetectReturnSlot (a CLASS method's hidden [out,retval] retbuf) this is
    'the plain accumulator return modules use - so there is NO hidden return-slot param.
    'Scans the RAW bytes (not the instruction collection): this epilogue lives in the
    'SEH tail PAST the proc`s `ret`, which proc-bounding excludes from the collection.
    On Error GoTo done
    NVAccumRet = False: NVAccumRetType = ""
    'Skip anything already classified by typeinfo / the proc list (class Function /
    'Property / a known Sub) - those are handled by the existing header path.
    Dim kind As String, mi As Long
    If NativeTryMethodKind(addr, kind) Then
        If InStr(kind, "Function") > 0 Or InStr(kind, "Get") > 0 Or InStr(kind, "Property") > 0 Or InStr(kind, "Sub") > 0 Then Exit Sub
    End If
    mi = NativeProcMatchIdx(addr)
    If mi >= 0 Then
        If InStr(SubNamelist(mi).kind, "Function") > 0 Or InStr(SubNamelist(mi).kind, "Property") > 0 Then Exit Sub
    End If

    'Bound the scan to this proc - the 8KB buffer overruns into the next procedure.
    Dim hi As Long, pp As Long
    hi = 8190
    For pp = 0 To UBound(gNativeProcArray) - 1
        If gNativeProcArray(pp).offset > addr Then
            Dim dd As Long: dd = gNativeProcArray(pp).offset - addr
            If dd > 0 And dd < hi Then hi = dd
        End If
    Next

    'Find the SEH-frame RESTORE `mov fs:[0],reg` (64 89 <m> 00 00 00 00) where reg is a
    'GP register (the setup writes esp = reg field 4, so skip that).  The 3 bytes before
    'it are `mov reg,[ebp-Z]` (8b <m2> <disp8>, m2 = md1/rm5); before THAT, for a
    'Function, is the return-value load `mov ax/eax,[ebp-retSlot]`.  Take the FIRST such
    'restore: it is this proc's epilogue (a Sub has no return load before its restore).
    Dim j As Long, ret8 As Long, isWord As Boolean
    For j = 8 To hi - 7
        If b(j) = &H64 And b(j + 1) = &H89 And (b(j + 2) And &HC7) = 5 _
           And ((b(j + 2) \ 8) And 7) <> 4 _
           And b(j + 3) = 0 And b(j + 4) = 0 And b(j + 5) = 0 And b(j + 6) = 0 Then
            'SEH-ptr load `mov reg,[ebp-Z]` (8b <m2:md1,rm5> <disp8>) immediately before.
            If j >= 3 Then
                If b(j - 3) = &H8B And (b(j - 2) And &HC7) = &H45 Then
                    'Return-value load ending at j-3.  Three encodings:
                    isWord = False: ret8 = 0
                    If j >= 7 And b(j - 7) = &H66 And b(j - 6) = &H8B And b(j - 5) = &H45 Then
                        isWord = True: ret8 = b(j - 4)                 'mov ax, word[ebp-d8]
                    ElseIf j >= 7 And b(j - 7) = &HF And (b(j - 6) = &HBF Or b(j - 6) = &HB7) And b(j - 5) = &H45 Then
                        isWord = True: ret8 = b(j - 4)                 'movsx/movzx eax, word[ebp-d8]
                    ElseIf j >= 6 And b(j - 6) = &H8B And b(j - 5) = &H45 And b(j - 7) <> &H66 Then
                        isWord = False: ret8 = b(j - 4)                'mov eax, dword[ebp-d8]
                    Else
                        Exit Sub        'no return load -> a Sub (don't scan into the next proc)
                    End If
                    Dim retDisp As Long
                    retDisp = ret8: If retDisp >= 128 Then retDisp = retDisp - 256   'signed disp8
                    If retDisp < 0 Then
                        NVAccumRet = True
                        NVAccumRetSlot = retDisp
                        If isWord Then NVAccumRetType = "Integer" Else NVAccumRetType = "Long"
                        Dim funcName As String, fp As Long
                        funcName = NativeProcName(addr)
                        fp = InStr(funcName, "("): If fp > 0 Then funcName = Trim$(Left$(funcName, fp - 1))
                        If NativeIsIdent(funcName) Then
                            ReDim Preserve NVArgTok(NVArgN): ReDim Preserve NVArgNm(NVArgN)
                            NVArgTok(NVArgN) = "var_" & Hex$(Abs(retDisp))
                            NVArgNm(NVArgN) = funcName
                            NVArgN = NVArgN + 1
                        End If
                    End If
                End If
            End If
            Exit Sub                    'first restore decides this proc
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

Private Function NativeGlobalTokVa(ByVal tok As String) As Long
    'Parse the VA out of a global_XXXXXXXX token (0 if malformed). On Error is
    'function-scoped, so a bad/overflowing parse can't disturb the caller.
    On Error Resume Next
    If Left$(tok, 7) = "global_" Then NativeGlobalTokVa = CLng("&H" & Mid$(tok, 8))
End Function

Private Sub NativeAddVarArg(ByVal v As String, ByVal base As String)
    'Append a Variant data-field value (a method argument being built) to the
    'pending argument list for a following control-method call.
    On Error Resume Next
    If NVVarArgN > UBound(NVVarArgList) Then ReDim Preserve NVVarArgList(NVVarArgN + 15): ReDim Preserve NVVarArgBase(NVVarArgN + 15)
    NVVarArgList(NVVarArgN) = v
    NVVarArgBase(NVVarArgN) = base
    NVVarArgN = NVVarArgN + 1
End Sub

Private Sub NativeResetVarArgs()
    NVVarArgN = 0
End Sub

Private Function NativeIsVarBuildLine(ByVal s As String) As Boolean
    'True when a statement is a Variant field-build line: var_<hex>(<digits>) = ...
    Dim t As String, p As Long, q As Long
    t = Trim$(Replace(s, vbCrLf, ""))
    If Left$(t, 4) <> "var_" Then Exit Function
    p = InStr(t, "(")
    If p < 5 Then Exit Function
    q = InStr(p, t, ")")
    If q <= p + 1 Then Exit Function
    If Not IsNumeric(Mid$(t, p + 1, q - p - 1)) Then Exit Function
    NativeIsVarBuildLine = (InStr(Mid$(t, q), ") = ") = 1)
End Function

Private Sub NativeSuppressVarBuild(ByVal base As String)
    'Mark a Variant-temp base (var_C) whose numeric-field build statements
    '(var_C(4/8/12)=...) should be stripped - its data was consumed by a late put.
    If Len(base) = 0 Or Left$(base, 4) <> "var_" Then Exit Sub
    If NVSuppressVarBuild Is Nothing Then Set NVSuppressVarBuild = New Collection
    On Error Resume Next
    NVSuppressVarBuild.Add base, base                'keyed -> deduped
End Sub

Private Function NativeStripVarBuild(ByVal src As String) As String
    'Remove the Variant-construction statements (base(<digit>) = ...) for any temp
    'whose value was folded into a late-bound property put, leaving just the put.
    NativeStripVarBuild = src
    If NVSuppressVarBuild Is Nothing Then Exit Function
    If NVSuppressVarBuild.Count = 0 Then Exit Function
    Dim lines() As String, i As Long, t As String, drop As Boolean, b As Variant, rest As String, cp As Long
    Dim outp As String
    lines = Split(src, vbCrLf)
    For i = 0 To UBound(lines)
        t = Trim$(lines(i))
        drop = False
        For Each b In NVSuppressVarBuild
            If Left$(t, Len(b) + 1) = b & "(" Then
                rest = Mid$(t, Len(b) + 2)
                cp = InStr(rest, ")")
                If cp > 1 Then
                    If IsNumeric(Left$(rest, cp - 1)) And InStr(Mid$(rest, cp), ") = ") = 1 Then drop = True: Exit For
                End If
            End If
        Next b
        If Not drop Then
            If Len(outp) > 0 Then outp = outp & vbCrLf
            outp = outp & lines(i)
        End If
    Next i
    NativeStripVarBuild = outp
End Function

Private Function NativeSubstituteConstants(ByVal src As String) As String
    'Replace distinctive Win32 magic numbers with their constant names.  The ternary
    'raster-ops (SRCCOPY etc.) are 24-bit structured codes that can't plausibly be an
    'ordinary number, but to be safe the substitution is scoped to procedures that
    'call a blit API (where the dwRop argument lives) - so a coincidental identical
    'value elsewhere (e.g. an RGB colour) is never renamed.  Whole-number-token match.
    NativeSubstituteConstants = src
    If InStr(src, "BitBlt") = 0 And InStr(src, "StretchBlt") = 0 And InStr(src, "PatBlt") = 0 _
       And InStr(src, "MaskBlt") = 0 And InStr(src, "PlgBlt") = 0 Then Exit Function
    Dim r As String
    r = src
    r = NativeSubConst(r, "13369376", "SRCCOPY", "&HCC0020")
    r = NativeSubConst(r, "15597702", "SRCPAINT", "&HEE0086")
    r = NativeSubConst(r, "8913094", "SRCAND", "&H8800C6")
    r = NativeSubConst(r, "6684742", "SRCINVERT", "&H660046")
    r = NativeSubConst(r, "4456232", "SRCERASE", "&H440328")
    r = NativeSubConst(r, "3342344", "NOTSRCCOPY", "&H330008")
    r = NativeSubConst(r, "1114278", "NOTSRCERASE", "&H1100A6")
    r = NativeSubConst(r, "12583114", "MERGECOPY", "&HC000CA")
    r = NativeSubConst(r, "12255782", "MERGEPAINT", "&HBB0226")
    r = NativeSubConst(r, "15728673", "PATCOPY", "&HF00021")
    r = NativeSubConst(r, "16452617", "PATPAINT", "&HFB0A09")
    r = NativeSubConst(r, "5898313", "PATINVERT", "&H5A0049")
    r = NativeSubConst(r, "5570569", "DSTINVERT", "&H550009")
    r = NativeSubConst(r, "16711778", "WHITENESS", "&HFF0062")
    NativeSubstituteConstants = r
End Function

Private Function NativeSubConst(ByVal src As String, ByVal decVal As String, ByVal name As String, ByVal hexVal As String) As String
    'Replace decVal -> name (whole-number token); when something was replaced, record
    'the constant so its `Public Const name = hexVal` declaration is reconstructed.
    NativeSubConst = NativeReplaceToken(src, decVal, name)
    If NativeSubConst <> src Then
        On Error Resume Next
        If gUsedWin32Const Is Nothing Then Set gUsedWin32Const = New Collection
        gUsedWin32Const.Add name & " = " & hexVal, name        'keyed by name -> dedup
        On Error GoTo 0
    End If
End Function

Public Function GetWin32ConstBlock(ByVal scope As String) As String
    'The `<scope> Const NAME = &Hvalue` block for the Win32 constants recognised by
    'value during this decompile (raster-ops etc.).  Emitted once in the first
    'standard module, like the API Declare block.
    If gUsedWin32Const Is Nothing Then Exit Function
    Dim v As Variant, s As String
    For Each v In gUsedWin32Const
        s = s & scope & " Const " & v & vbCrLf
    Next
    GetWin32ConstBlock = s
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

Private Function NativeControlIsArrayByOffset(ByVal offset As Long) As Boolean
    'True when the control accessed at form-vtable `offset` is a control ARRAY (its
    'element accessor, vtable 0x40, returns element i).  Used to gate the control-array
    'element reconstruction strictly so a non-array control call never misfires.
    'A control array surfaces in gControlNameArray either as MULTIPLE entries sharing
    'the same (form, name) - one per defined element - or, for a single-element array,
    'as one entry with the is-array IID flag set.  Resolve the name at this offset and
    'test both signals (so any element accessor of the array qualifies, not only those
    'whose own IID happened to be the array-member IID+1).
    Dim nm As String
    nm = NativeControlByOffset(offset)
    If Len(nm) = 0 Then Exit Function
    NativeControlIsArrayByOffset = NativeNameIsControlArray(nm)
End Function

Private Function NativeNameIsControlArray(ByVal nm As String) As Boolean
    'True when control `nm` on the current form is a control array: more than one
    'gControlNameArray entry shares its name (one per element), or any such entry
    'carries the array-member IID flag.
    Dim k As Long, cnt As Long
    On Error Resume Next
    For k = 0 To UBound(gControlNameArray)
        If gControlNameArray(k).strParentForm = NVForm And gControlNameArray(k).strControlName = nm Then
            cnt = cnt + 1
            If gControlNameArray(k).bIsArray <> 0 Then NativeNameIsControlArray = True: Exit Function
            If cnt > 1 Then NativeNameIsControlArray = True: Exit Function
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
        'For a SIB element store the pre-pass may have recovered the logical index;
        'render base(i)(off) so `Player(dst).f = ...` keeps its element index instead
        'of collapsing to base(off) (which made distinct elements look identical).
        Dim eix As String
        eix = NativeColGet(NVElemIdx, "E" & NVCurVa)
        If Len(eix) > 0 Then
            NativeFieldStoreLHS = NVReg(base) & "(" & eix & ")(" & CStr(off) & ")"
        Else
            NativeFieldStoreLHS = NVReg(base) & "(" & CStr(off) & ")"
        End If
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

Private Function NativeIsRelationalExpr(ByVal s As String) As Boolean
    'A parenthesised relational/boolean expression (a comparison result), e.g.
    '`(var_7C <> var_4C)` or `(a Is Nothing)` - the materialised result of
    '__vbaVarTstNe / __vbaStrCmp / __vbaObjIs.  Used to gate the narrow 16-bit
    'reg-reg propagation (a Boolean is -1/0, so the low word is the whole value).
    If Left$(s, 1) <> "(" Then Exit Function
    If InStr(s, " <> ") > 0 Or InStr(s, " = ") > 0 Or InStr(s, " Is ") > 0 _
       Or InStr(s, " < ") > 0 Or InStr(s, " > ") > 0 _
       Or InStr(s, " <= ") > 0 Or InStr(s, " >= ") > 0 Then NativeIsRelationalExpr = True
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
            If Len(sv) > 0 Then
                NVReg(op - &HB8) = sv
            ElseIf NativeIsGlobalAddr(immv) Then
                'mov reg, &global - the immediate IS a module-global's ADDRESS, loaded to
                'pass it by reference (a store DEST for __vbaStrMove / a ByRef arg).  Name
                'it global_<va> so a following `__vbaStrMove(dest=ecx)` renders the real
                'target (Module1.packetValue) instead of the bare numeric literal / a
                'stale lea.  (`mov reg, [global]` - the VALUE - is opcode 0x8B/0xA1, not
                'this imm form, so this only fires for the address-of idiom.)
                NVReg(op - &HB8) = NativeGlobalName(immv)
            Else
                NVReg(op - &HB8) = NativeNumFromBits(immv)
            End If
            NVRegIsAddr(op - &HB8) = False: NVRegIsMe(op - &HB8) = False: NVRegIsFormVt(op - &HB8) = False
            NVRegObjType(op - &HB8) = "": NVRegObjVt(op - &HB8) = "": NVRegObjGuid(op - &HB8) = "": NVRegObjVtGuid(op - &HB8) = ""
        Case &H8B                       'mov r32, r/m32
            modrm = NativeDumpByte(dump, i + 1)
            md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
            If NativeHas66(dump) And md = 3 And NativeIsRelationalExpr(NVReg(rm)) Then
                'NARROW exception to the 16-bit clear below: the source register holds a
                'RELATIONAL/boolean expression (a `(a <> b)` from __vbaVarTstNe /
                '__vbaStrCmp / __vbaObjIs).  A 16-bit `mov bx,ax` here copies the
                'VARIANT_BOOL result whose VALUE is the low word (-1/0 - no high half to
                'lose), so propagate it: it must survive to the deferred `test bx,bx`/jcc
                'that forms the loop/If condition (VB moves the result into a callee-saved
                'register across the Variant cleanup calls).  Only relational sources
                'qualify, so the general truncation hazard below is untouched.
                NVReg(reg) = NVReg(rm)
                NVRegIsAddr(reg) = False: NVRegIsMe(reg) = False: NVRegIsFormVt(reg) = False
                NVRegObjType(reg) = "": NVRegObjVt(reg) = "": NVRegObjGuid(reg) = "": NVRegObjVtGuid(reg) = ""
            ElseIf NativeHas66(dump) Then
                'A 16-bit move (mov si,ax) writes only the LOW WORD of the dest; we
                'model whole 32-bit values, so the dest is now unknown.  Copying the
                'source clobbered the high half with a stale value and collapsed
                'conditions like `SendMessage(...) <> 0` into `0 <> 0`.
                NVReg(reg) = "": NVRegIsAddr(reg) = False: NVRegIsMe(reg) = False: NVRegIsFormVt(reg) = False
                NVRegObjType(reg) = "": NVRegObjVt(reg) = "": NVRegObjGuid(reg) = "": NVRegObjVtGuid(reg) = ""
                'BUT when the source is MEMORY (not a register low-word partial), the
                'loaded word IS the whole value of an Integer/Boolean local, parameter,
                'array element or struct field.  Capture its symbolic value in the
                'compare-only shadow so a following `cmp ax,imm16` resolves to e.g.
                '`Item(i).ID >= 1` (a `Select Case ... Case 1 To 20` range check) or
                '`var_X = 5`, instead of the generic decoder dropping it to <cond>.
                'Kept OUT of NVReg on purpose (see the shadow's declaration).
                If md <> 3 Then
                    Dim w16 As String
                    w16 = NativeRmVal(dump, md, rm)
                    If Len(w16) > 0 Then NVR16Val(reg) = w16
                End If
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
                ElseIf Not isAbs And NativeMemIndex(dump) >= 0 Then
                    'SIB element READ [pvData + idx + field]: track the element value
                    'so a field-to-field copy `field = REG` reconstructs the source
                    'element (global_X(12)(i)(off)) instead of leaking the bare
                    'register.  NativeRmVal mirrors the compare/store renderers and
                    'injects the recovered index (NVElemIdx, keyed by NVCurVa); it
                    'returns "" when the base is not a tracked array pointer, so a
                    'non-element SIB read stays untracked as before.
                    NVReg(reg) = NativeRmVal(dump, md, rm)
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
                Dim hadFieldCls As String, hadFieldRecv As String
                hadFieldCls = "": hadFieldRecv = ""
                If Not isAbs And disp = 0 And bse >= 0 And bse <= 7 And Len(NVRegFieldCls(bse)) > 0 Then
                    hadFieldCls = NVRegFieldCls(bse): hadFieldRecv = NVRegFieldRecv(bse)
                End If
                NVRegFieldCls(reg) = "": NVRegFieldRecv(reg) = ""   'the dest holds a new value
                If Len(hadFieldCls) > 0 Then
                    'Deref of an As-New field ADDRESS (&Me.<field>) -> the field's OBJECT.
                    'Tag the object identity (class + field_<off> receiver) so a following
                    '`mov vt,[obj]; call [vt+off]` resolves to field_<off>.Method
                    '(e.g. picBackBmp.LoadBitmap, accessed via lea&deref not a direct mov).
                    'NVReg is deliberately left empty: the object pointer has no meaningful
                    'VALUE to render, and putting "field_<off>" there leaked the receiver
                    'into a later argument push (field_40 into a BitBlt coordinate slot).
                    NVRegObjType(reg) = hadFieldCls: NVRegObjInst(reg) = hadFieldRecv
                ElseIf Not isAbs And disp < 0 Then
                    If NativeIsIntrinsicObj(NVReg(reg)) Then NVRegObjType(reg) = NVReg(reg)
                    'Loading a local that __vbaNew2 typed with a user class (an `As New`
                    'instance, e.g. pktCreate): tag it so var_X.Member resolves.
                    Dim lcls As String
                    lcls = NativeColGet(NVLocalObjType, "D" & disp)
                    If Len(lcls) > 0 Then NVRegObjType(reg) = lcls: NVRegObjInst(reg) = "var_" & Hex$(Abs(disp))
                    'Loading a local that holds a control object tags this register
                    'with the control's identity + GUID (for a later property call).
                    Dim lguid As String
                    lguid = NativeGetLocalGuid(disp)
                    If Len(lguid) > 0 Then NVRegObjGuid(reg) = lguid: NVRegObjType(reg) = NVReg(reg)
                    'Reloading an object's cached VTABLE (stored by the 0x89 handler):
                    'restore its identity (control name+GUID, or an intrinsic vtable name
                    'like _Global) so a deferred call `call [reg + off]` through it
                    'resolves (e.g. the cached _Global vtable for `Unload Me`).
                    Dim lvn As String
                    lvn = NativeColGet(NVLocalVtName, "L" & disp)
                    If Len(lvn) > 0 Then
                        NVRegObjVt(reg) = lvn
                        NVRegObjVtGuid(reg) = NativeColGet(NVLocalVtGuid, "L" & disp)
                    End If
                ElseIf isAbs And NativeIsGlobalAddr(disp) Then
                    'Loading a module-global that holds a user-class instance (typed
                    'at its __vbaNew auto-instantiation) tags the register as that
                    'object pointer, so a following `mov vt,[obj]; call [vt+off]`
                    'resolves to obj.Method.
                    'Type the register holding this module-global so a following
                    '`mov vt,[obj]; call [vt+off]` resolves to obj.Method.  Priority:
                    '  1. FormNameByInstGlobal - a REGISTERED form-instance global; definitive
                    '     (and overrides a per-proc NVObjClass a coincidental same-address push
                    '     may have mis-set, now that BSS globals are recognised).
                    '  2. NVObjClass - a user-class instance typed at its __vbaNew in THIS proc.
                    '  3. gNativeGlobalClass - a `Public X As <Class>` proven by a resolved call
                    '     elsewhere (cross-proc), e.g. global_0055D5A0 As frmClient.
                    Dim gcls As String
                    gcls = FormNameByInstGlobal(disp)
                    If Len(gcls) = 0 Then gcls = NativeColGet(NVObjClass, "G" & disp)
                    If Len(gcls) = 0 And Not gNativeGlobalClass Is Nothing Then gcls = NativeColGet(gNativeGlobalClass, "g" & disp)
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
        Case &H8A                       'mov r8, r/m8  (byte load)
            modrm = NativeDumpByte(dump, i + 1)
            md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
            'A byte deref of a ByRef parameter pointer that was just zero-extended
            '(`xor ecx,ecx; mov cl,[edx]`) yields the parameter's (small) value - VB6
            'reads a small Integer ubound this way.  Only the low-byte registers
            'al/cl/dl/bl (0..3) zero-extend a full register, and only when it already
            'tracks "0" (the xor), so the loaded byte IS the whole value.
            If md <> 3 And reg <= 3 And NVReg(reg) = "0" Then
                If NativeDecodeDisp(dump, disp, isAbs) Then
                    Dim baseB As Long
                    baseB = NativeMemBase(dump)
                    If Not isAbs And disp = 0 And baseB >= 0 And baseB <= 7 And Left$(NVReg(baseB), 4) = "arg_" Then
                        NVReg(reg) = NVReg(baseB)
                        NVRegIsAddr(reg) = False: NVRegIsMe(reg) = False: NVRegIsFormVt(reg) = False
                        NVRegObjType(reg) = "": NVRegObjVt(reg) = "": NVRegObjGuid(reg) = "": NVRegObjVtGuid(reg) = ""
                    End If
                End If
            End If
        Case &H8D                       'lea r32, [mem]  (address-of / ptr arithmetic)
            modrm = NativeDumpByte(dump, i + 1)
            reg = (modrm \ 8) And 7
            NVRegFieldCls(reg) = "": NVRegFieldRecv(reg) = ""   'cleared unless this lea is an As-New field address
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
                ElseIf Not isAbs And disp > 0 And lbse >= 0 And lbse <= 7 And NVRegIsMe(lbse) And NativeMemIndex(dump) < 0 _
                       And Len(NativeColGet(gFormFieldClass, NVForm & ":" & disp)) > 0 Then
                    'lea reg,[Me + fieldOff] = the ADDRESS of an As-New object field
                    '(taken to pass to __vbaNew2 in the auto-instantiation guard, then
                    'dereferenced).  Remember the field's class + a field_<off> receiver
                    'so the following `mov obj,[reg]` tags obj as that clsX object and a
                    'method call on it resolves to field_<off>.Method (picBackBmp.LoadBitmap).
                    NVReg(reg) = "field_" & Hex$(disp): NVRegIsAddr(reg) = False
                    NVRegFieldCls(reg) = NativeColGet(gFormFieldClass, NVForm & ":" & disp)
                    NVRegFieldRecv(reg) = "field_" & Hex$(disp)
                    NVRegIsAddr(reg) = False
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
                    'Effective stored value: the register's tracked 32-bit value, or -
                    'when that was cleared because the register was last loaded as a
                    '16-bit memory WORD (mov cx,[Me+off]; mov [ebp-X],ecx) - the 16-bit
                    'shadow, which carries the Integer field/var value (an Integer
                    'Property Get reads `temp = mvarID` this way; without it the body
                    'vanished entirely).
                    Dim stv89 As String
                    stv89 = NVReg(reg)
                    If Len(stv89) = 0 And Len(NVR16Val(reg)) > 0 Then stv89 = NVR16Val(reg)
                    'Record the register's value against the slot (a Variant data or
                    'VT field filled from a register, e.g. `mov [ebp-X], esi` with
                    'esi = 0xA for a missing optional argument).
                    NativeSetVSlot disp, stv89
                    'A stored call result (expression containing a call) is worth
                    'surfacing as a real assignment; bind the local to its name so
                    'later uses reference the variable rather than re-expanding.  A
                    'loop induction slot (NVCounterSlot) is treated the same way even
                    'for a plain value, so the counter shows as var_X (not a stale
                    'per-iteration constant) and its init / increment surface.
                    Dim isCtrSlot As Boolean
                    isCtrSlot = Len(NativeColGet(NVCounterSlot, "C" & disp)) > 0
                    lname = "var_" & Hex$(Abs(disp))
                    If (InStr(stv89, "(") > 0 Or Left$(stv89, 6) = "field_") And Left$(stv89, 1) <> Chr$(34) Then
                        'Surface a meaningful value as `var_X = <expr>`: a call/deref
                        '(parenthesised) OR a bare instance-field read (field_<off>) - so an
                        'Integer Property Get's `temp = field_34` return-slot copy shows and
                        'the post-pass renames it to `ID = field_34`.
                        NativeTrackReg = lname & " = " & stv89
                        NativeSetLocalExpr disp, lname
                    ElseIf isCtrSlot And Len(stv89) > 0 And stv89 <> lname Then
                        NativeTrackReg = lname & " = " & stv89
                        NativeSetLocalExpr disp, lname
                    Else
                        NativeSetLocalExpr disp, stv89
                    End If
                    'Storing a tracked control object to a local (plain mov, not
                    '__vbaObjSet) - remember its GUID so a later property access
                    'through that local resolves (e.g. the LET target temp).
                    If Len(NVRegObjGuid(reg)) > 0 Then NativeSetLocalGuid disp, NVRegObjGuid(reg)
                    'Storing an object's VTABLE to a local (VB caches it before a call
                    'when intervening helpers would clobber the register - e.g. the first
                    'lblSkillName(i).Caption put, or the _Global vtable before `Unload`).
                    'Thread the vtable's identity (name, and GUID when a control) through
                    'the slot so the reload resolves.
                    On Error Resume Next
                    NVLocalVtName.Remove "L" & disp: NVLocalVtGuid.Remove "L" & disp
                    On Error GoTo 0
                    If Len(NVRegObjVt(reg)) > 0 Then
                        NVLocalVtName.Add NVRegObjVt(reg), "L" & disp
                        If Len(NVRegObjVtGuid(reg)) > 0 Then NVLocalVtGuid.Add NVRegObjVtGuid(reg), "L" & disp
                    End If
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
                    Dim fb89 As Long, fbMe89 As Boolean
                    fb89 = NativeMemBase(dump)
                    lhs89 = NativeFieldStoreLHS(fb89, disp)
                    If Len(lhs89) > 0 Then
                        'Harvest the field's type for the class field-decl block: a
                        '0x66-prefixed store writes a 16-bit Integer, a plain store a Long.
                        '(Guard the base index before NVRegIsMe - VB6 And is not short-circuit.)
                        fbMe89 = False
                        If fb89 >= 0 And fb89 <= 7 Then fbMe89 = (NVHasMe And NVRegIsMe(fb89))
                        If fbMe89 Then NativeRecordFieldType disp, IIf(NativeHas66(dump), "Integer", "Long")
                        fv89 = NVReg(reg)
                        'Prefer the 16-bit word shadow when set: the stored register was
                        'sign-extended from a tracked 16-bit expression (movsx of a folded
                        '`(Index + vsCarry.Value)`), which NVReg does not carry.  Only set
                        'on a fresh 16-bit chain, so a normal 32-bit store is unaffected.
                        If Len(NVR16Val(reg)) > 0 Then fv89 = NVR16Val(reg)
                        If Len(fv89) = 0 Then fv89 = NativeRegName(reg)
                        'A tracked RHS identical to the LHS is a same-element
                        'read-modify-write whose intervening arithmetic did NOT fold
                        'onto the element value (NativeIsFoldableArith deliberately
                        'excludes deref chains, e.g. `Stamina = Stamina - 1`): rather
                        'than emit a misleading no-op `x = x`, fall back to the honest
                        'raw register.
                        If fv89 = lhs89 Then fv89 = NativeRegName(reg)
                        NativeTrackReg = lhs89 & " = " & fv89
                        'A write to a Variant's DATA field (offset +8) is the value a
                        'following late put / method call consumes - remember it (and
                        'the temp's base, to strip its build statements if consumed).
                        If disp = 8 And Len(fv89) > 0 And Left$(fv89, 1) <> "e" Then
                            NVLastVarData = fv89: NVLastVarBase = NVReg(NativeMemBase(dump))
                            NativeAddVarArg fv89, NVLastVarBase
                        End If
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
                ElseIf Not isAbs And disp < 0 And disp = NVAccumRetSlot Then
                    'A constant stored to the FUNCTION RETURN slot: `mov [ebp-retSlot],
                    'imm` (e.g. modMap_Direction = 7).  Render `var_<slot> = imm` (the
                    'return-slot rename then makes it `FuncName = imm`).  Gated to the
                    'return slot so the SEH-frame / Variant-VT 0xC7 stores to other
                    'locals stay suppressed (they are not user assignments).
                    Dim rcImm As Long, rcs As String
                    If NativeHas66(dump) Then rcImm = NativeDumpInt16(dump, n - 2) Else rcImm = NativeDumpInt32(dump, n - 4)
                    rcs = ""
                    If rcImm >= OptHeader.ImageBase Then rcs = NativeStringAt(rcImm)
                    If Len(rcs) = 0 Then rcs = NativeNumFromBits(rcImm)
                    NativeTrackReg = "var_" & Hex$(Abs(disp)) & " = " & rcs
                    NativeSetLocalExpr disp, rcs
                ElseIf Not isAbs And disp > 0 And disp < &H2000 Then
                    'Store an immediate to a struct FIELD: mov [base + off], imm.
                    'A 0x66-prefixed store writes a word (Boolean True = 0xFFFF -> -1).
                    Dim lhsC7 As String, fimm As Long, fis As String, fbC7 As Long, fbMeC7 As Boolean
                    fbC7 = NativeMemBase(dump)
                    lhsC7 = NativeFieldStoreLHS(fbC7, disp)
                    If Len(lhsC7) > 0 Then
                        If NativeHas66(dump) Then fimm = NativeDumpInt16(dump, n - 2) Else fimm = NativeDumpInt32(dump, n - 4)
                        If fimm >= OptHeader.ImageBase Then fis = NativeStringAt(fimm)
                        If Len(fis) = 0 Then fis = NativeNumFromBits(fimm)
                        'Harvest the field type for the class field-decl block.
                        fbMeC7 = False
                        If fbC7 >= 0 And fbC7 <= 7 Then fbMeC7 = (NVHasMe And NVRegIsMe(fbC7))
                        If fbMeC7 Then
                            If Left$(fis, 1) = Chr$(34) Then
                                NativeRecordFieldType disp, "String"
                            Else
                                NativeRecordFieldType disp, IIf(NativeHas66(dump), "Integer", "Long")
                            End If
                        End If
                        NativeTrackReg = lhsC7 & " = " & fis
                        'Variant DATA field (offset +8) immediate -> remember as the
                        'value (and base temp) for a following late put / method call.
                        If disp = 8 And Len(fis) > 0 Then
                            NVLastVarData = fis: NVLastVarBase = NVReg(NativeMemBase(dump))
                            NativeAddVarArg fis, NVLastVarBase
                        End If
                    End If
                End If
            End If
        Case &HF                        'two-byte opcode: movsx / movzx r32, r/m16
            Dim mxop2 As Long
            mxop2 = NativeDumpByte(dump, i + 1)
            If mxop2 = &HBF Or mxop2 = &HBE Or mxop2 = &HB7 Or mxop2 = &HB6 Then
                modrm = NativeDumpByte(dump, i + 2)
                md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
                'Sign/zero-extend of a register whose 16-bit word shadow we tracked (e.g.
                'the folded `(arg_C + frmMain.vsCarry.Value)` in cx) - carry that
                'expression into the DEST's word shadow (NOT NVReg: putting it in NVReg
                'leaked the 16-bit expr into unrelated 32-bit arithmetic folds and
                'corrupted them).  The field store below prefers NVR16Val, so
                'FocusCarryIndex = Index + vsCarry.Value still reconstructs.  Only reg->reg
                'with a live word shadow; otherwise leave everything as-is.
                If md = 3 And rm >= 0 And rm <= 7 And Len(NVR16Val(rm)) > 0 Then
                    NVR16Val(reg) = NVR16Val(rm)
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
        Case &H83, &H81                 'add/sub r/m32, imm8/imm32 -> fold onto a tracked value
            modrm = NativeDumpByte(dump, i + 1)
            md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
            'add <Me>, <fieldOffset> computes the ADDRESS of an instance field (VB6
            'does this to pass a field's array by reference to __vbaRedim).  Re-tag the
            'register as that field so the by-ref push renders the field name (the
            'array variable) instead of the stale Me/arg_8.  Done before the arithmetic
            'fold so `add esi,0x38` (esi = Me) becomes the field, not `(arg_8 + 56)`.
            If op = &H83 Or op = &H81 Then
                If md = 3 And reg = 0 And rm <> 4 And rm <> 5 And NVRegIsMe(rm) Then
                    Dim fOff As Long
                    If op = &H83 Then fOff = NativeDumpInt8(dump, n - 1) Else fOff = NativeDumpInt32(dump, n - 4)
                    If fOff > 0 And fOff < &H2000 Then
                        NVReg(rm) = NativeFieldName(fOff)
                        NVRegIsAddr(rm) = False: NVRegIsMe(rm) = False: NVRegIsFormVt(rm) = False
                        NVRegObjType(rm) = "": NVRegObjVt(rm) = "": NVRegObjGuid(rm) = "": NVRegObjVtGuid(rm) = ""
                        Exit Function
                    End If
                End If
            End If
            'reg-field 0 = ADD, 5 = SUB; rm = the destination register (md=3).  Extend a
            'meaningful tracked value (a call result like Len(s), or a clean variable)
            'with the immediate so `Right$(s, Len(s) - 4)` keeps the "- 4".  Skip esp/ebp.
            If md = 3 And (reg = 0 Or reg = 5) And rm <> 4 And rm <> 5 Then
                Dim aimm As Long
                If op = &H83 Then aimm = NativeDumpInt8(dump, n - 1) Else aimm = NativeDumpInt32(dump, n - 4)
                If NativeHas66(dump) Then
                    'A 16-bit add/sub (sub cx,55) folds onto the 16-bit shadow, which
                    'carries an Integer field/expression (an Integer Function builds its
                    'result this way: `mvarID * 1000 - 55`).  NVReg's 16-bit value was
                    'cleared, so fold NVR16Val and leave NVReg alone.  A 16-bit register
                    'is always a VALUE (never a pointer), so the permissive
                    'NativeIs16Foldable gate is used (NativeIsFoldableArith would reject
                    'the field-read deref arg_8(52) as pointer math).
                    If NativeIs16Foldable(NVR16Val(rm)) Then
                        NVR16Val(rm) = NativeFoldArith(NVR16Val(rm), (reg = 5), aimm)
                    End If
                ElseIf NativeIsFoldableArith(NVReg(rm)) Then
                    NVReg(rm) = NativeFoldArith(NVReg(rm), (reg = 5), aimm)
                    NVRegIsAddr(rm) = False: NVRegIsMe(rm) = False: NVRegIsFormVt(rm) = False
                    NVRegObjType(rm) = "": NVRegObjVt(rm) = "": NVRegObjGuid(rm) = "": NVRegObjVtGuid(rm) = ""
                End If
            End If
        Case &H69, &H6B                 'imul reg, r/m, imm -> (r/m * imm) into reg
            modrm = NativeDumpByte(dump, i + 1)
            md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
            If md = 3 And rm <> 4 And rm <> 5 And reg <> 4 And reg <> 5 Then
                Dim mimm As Long, msrc As String
                If op = &H6B Then
                    mimm = NativeDumpInt8(dump, n - 1)
                ElseIf NativeHas66(dump) Then
                    mimm = NativeDumpInt16(dump, n - 2)
                Else
                    mimm = NativeDumpInt32(dump, n - 4)
                End If
                If NativeHas66(dump) Then
                    'A 16-bit `imul cx,cx,1000` building an Integer expression (mvarID *
                    '1000): multiply the 16-bit shadow value, leaving NVReg cleared.
                    msrc = NVR16Val(rm)
                    If NativeIs16Foldable(msrc) Then NVR16Val(reg) = "(" & msrc & " * " & CStr(mimm) & ")"
                ElseIf NativeIsFoldableArith(NVReg(rm)) Then
                    NVReg(reg) = "(" & NVReg(rm) & " * " & CStr(mimm) & ")"
                    NVRegIsAddr(reg) = False: NVRegIsMe(reg) = False: NVRegIsFormVt(reg) = False
                    NVRegObjType(reg) = "": NVRegObjVt(reg) = "": NVRegObjGuid(reg) = "": NVRegObjVtGuid(reg) = ""
                End If
            End If
        Case &H2D, &H5                  'sub/add eax, imm32
            If NativeIsFoldableArith(NVReg(0)) Then
                NVReg(0) = NativeFoldArith(NVReg(0), (op = &H2D), NativeDumpInt32(dump, i + 1))
                NVRegIsAddr(0) = False: NVRegIsMe(0) = False: NVRegIsFormVt(0) = False
                NVRegObjType(0) = "": NVRegObjVt(0) = "": NVRegObjGuid(0) = "": NVRegObjVtGuid(0) = ""
            End If
        Case &H03, &H2B                 'add / sub  r32, r/m32  -> fold into reg dest
            modrm = NativeDumpByte(dump, i + 1)
            md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
            If NativeHas66(dump) And Len(NVR16Val(reg)) > 0 Then
                '16-bit add/sub onto a value tracked in the WORD shadow (e.g. arg_C from
                '`mov cx,word[arg_C]`): the 32-bit value is unknown, so fold ONLY NVR16Val,
                'never NVReg.  R = the source operand (a local/field/element, not a bare
                'register).  Lets `FocusCarryIndex = Index + vsCarry.Value` reconstruct: a
                'following `movsx edx,cx` carries the folded word forward and the field
                'store renders it.  When the shadow is empty the value lives in NVReg, so
                'fall through to the 32-bit fold below (unchanged) - e.g. `mov ax,1;
                'add ax,[var_20]` still builds `(1 + var_20)` for a loop counter.
                Dim w16R As String, w16op As String
                w16R = NativeRmVal(dump, md, rm)
                If Len(w16R) > 0 And NativeIsCleanIndex(w16R) _
                   And (Len(NVR16Val(reg)) + Len(w16R)) <= 60 Then
                    If op = &H2B Then w16op = " - " Else w16op = " + "
                    NVR16Val(reg) = "(" & NVR16Val(reg) & w16op & w16R & ")"
                Else
                    NVR16Val(reg) = ""
                End If
                Exit Function
            End If
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

Private Function NativeUnkVCall(ByVal disp As Long) As String
    'Render an unresolved object vtable call as `Call [<obj>.]UnkVCall_<8hex>h(args)`
    'and drain the push stack.  COM pushes the implicit `this` last (topmost), so it
    'is the receiver; the pushes below it are the arguments in source order.  Read
    'NVPushImm directly (not NativeArgList) so an UNTRACKED `this` push - an empty
    'string - is preserved as the receiver slot (-> a leading-dot `.UnkVCall`) instead
    'of being collapsed into the first argument.
    Dim recv As String, args As String, i As Long
    If NVPushTop >= 1 Then recv = NVPushImm(NVPushTop - 1)
    For i = NVPushTop - 2 To 0 Step -1
        If Len(args) > 0 Then args = args & ", "
        args = args & NVPushImm(i)
    Next
    NVPushTop = 0
    'Only name the receiver when it is a STABLE object identity (a parameter, a
    'module global, Me, or a dotted Form.Control / App chain).  A bare local var_X
    'or a raw register is stale-prone - its tracked object can be wrong - so those
    'render as a leading-dot `.UnkVCall`, exactly as the commercial decompiler does.
    If Not NativeStableRecv(recv) Then recv = ""
    Dim nm As String
    nm = "UnkVCall_" & Right$("00000000" & Hex$(disp), 8) & "h"
    If Len(recv) > 0 Then nm = recv & "." & nm Else nm = "." & nm
    NativeUnkVCall = "Call " & nm & "(" & args & ")"
End Function

Private Function NativeStableRecv(ByVal r As String) As Boolean
    'A receiver object expression that is a STABLE identity, safe to name as the
    'method receiver: a parameter (arg_X), a module global (global_X), Me / Me.x, or
    'a dotted chain (Form.Control / App.Prop).  A bare local (var_X), a raw register,
    'or a number is rejected (its object identity is stale-prone).
    If Len(r) = 0 Then Exit Function
    If Left$(r, 4) = "arg_" Then NativeStableRecv = True: Exit Function
    If Left$(r, 7) = "global_" Then NativeStableRecv = True: Exit Function
    If r = "Me" Or Left$(r, 3) = "Me." Then NativeStableRecv = True: Exit Function
    If InStr(r, ".") > 0 Then NativeStableRecv = True
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
    'MsgBox's second argument is the Buttons bitfield (vbYesNoCancel Or vbInformation
    'etc.) - decompose a plain numeric value into its vb* constants.  trimZero marks
    'the MsgBox form (InputBox's 2nd arg is the title string, not buttons).
    If trimZero And idx >= 2 Then
        If IsNumeric(vals(1)) Then vals(1) = NativeMsgBoxButtons(CLng(vals(1)))
    End If
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

Private Function NativeMsgBoxButtons(ByVal v As Long) As String
    'Decompose a MsgBox Buttons bitfield into its vb* constants joined by " Or ".
    'Falls back to the raw number when the value does not decode cleanly (so no bit
    'is ever silently dropped).  vbOKOnly (0) contributes nothing and is omitted.
    Dim parts(8) As String, np As Long, recon As Long, grp As Long, icon As Long, dft As Long, i As Long, s As String
    grp = v And &HF
    Select Case grp
        Case 1: parts(np) = "vbOKCancel": np = np + 1: recon = recon Or grp
        Case 2: parts(np) = "vbAbortRetryIgnore": np = np + 1: recon = recon Or grp
        Case 3: parts(np) = "vbYesNoCancel": np = np + 1: recon = recon Or grp
        Case 4: parts(np) = "vbYesNo": np = np + 1: recon = recon Or grp
        Case 5: parts(np) = "vbRetryCancel": np = np + 1: recon = recon Or grp
        Case 0:                                  'vbOKOnly - the default, omit
    End Select
    icon = v And &H70
    Select Case icon
        Case &H10: parts(np) = "vbCritical": np = np + 1: recon = recon Or icon
        Case &H20: parts(np) = "vbQuestion": np = np + 1: recon = recon Or icon
        Case &H30: parts(np) = "vbExclamation": np = np + 1: recon = recon Or icon
        Case &H40: parts(np) = "vbInformation": np = np + 1: recon = recon Or icon
    End Select
    dft = v And &H300
    Select Case dft
        Case &H100: parts(np) = "vbDefaultButton2": np = np + 1: recon = recon Or dft
        Case &H200: parts(np) = "vbDefaultButton3": np = np + 1: recon = recon Or dft
    End Select
    If (v And &H1000) <> 0 Then parts(np) = "vbSystemModal": np = np + 1: recon = recon Or &H1000
    'Unrecognised bits remain -> keep the raw number rather than lose them.
    If recon <> v Or np = 0 Then NativeMsgBoxButtons = CStr(v): Exit Function
    For i = 0 To np - 1
        If Len(s) > 0 Then s = s & " Or "
        s = s & parts(i)
    Next
    NativeMsgBoxButtons = s
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

Private Sub NativeScanRetbufFuncs()
    'Find module Functions that return a Variant/String/UDT through a hidden retbuf: the
    'epilogue returns the retbuf POINTER (the first param) in eax - `mov eax,[ebp+8]`
    '(8B 45 08) - right before the SEH-frame restore + `ret N`.  The first param (arg_8)
    'IS the retbuf return slot, so a caller renders `<dest> = proc(<rest>)`.  Distinct from
    'the accumulator return (NativeDetectAccumReturn loads a NEGATIVE-disp local into ax/eax).
    'Byte-scans the raw proc bytes (the epilogue lives in the SEH tail past the body `ret`).
    Set gRetbufFunc = New Collection
    On Error Resume Next
    Dim pp As Long, addr As Long, b() As Byte, fp As Integer, hi As Long, qq As Long, dd As Long
    Dim j As Long, k As Long, restoreAt As Long, retN As Long, lo As Long, foundLoad As Boolean
    For pp = 0 To UBound(gNativeProcArray) - 1
        addr = gNativeProcArray(pp).offset
        If addr = 0 Then GoTo nextpp
        'ONLY .bas module procedures: a retbuf is the first param (ebp+8).  A class/form
        'method has Me/this at ebp+8 instead, and its epilogue `mov eax,[ebp+8]` is Me
        'cleanup - NOT a retbuf return (this wrongly promoted event handlers / methods).
        If NativeProcHasMe(addr) Then GoTo nextpp
        addr = NativeSnapEntry(addr)
        hi = 8190
        For qq = 0 To UBound(gNativeProcArray) - 1
            dd = gNativeProcArray(qq).offset - addr
            If dd > 0 And dd < hi Then hi = dd
        Next
        ReDim b(8191)
        fp = FreeFile
        Open SFilePath For Binary Access Read As #fp
            Get #fp, addr + 1 - OptHeader.ImageBase, b
        Close #fp
        'Locate the SEH-frame restore `mov fs:[0],reg` (64 89 <m> 00000000, reg<>esp).
        restoreAt = -1
        For j = 8 To hi - 7
            If b(j) = &H64 And b(j + 1) = &H89 And (b(j + 2) And &HC7) = 5 _
               And ((b(j + 2) \ 8) And 7) <> 4 _
               And b(j + 3) = 0 And b(j + 4) = 0 And b(j + 5) = 0 And b(j + 6) = 0 Then
                restoreAt = j: Exit For
            End If
        Next j
        If restoreAt < 0 Then GoTo nextpp
        'The retbuf-ptr load `mov eax,[ebp+8]` (8B 45 08) must sit in the epilogue, within
        '~64 bytes before the restore (the result is copied to [retbuf] right after it).
        foundLoad = False
        lo = restoreAt - 64: If lo < 8 Then lo = 8
        For k = lo To restoreAt - 1
            If b(k) = &H8B And b(k + 1) = &H45 And b(k + 2) = 8 Then foundLoad = True: Exit For
        Next k
        If Not foundLoad Then GoTo nextpp
        'Epilogue ends with `ret N` (C2 lo hi); a plain `ret` (C3) before it means this is
        'not the stdcall arg-cleanup epilogue we want.
        retN = -1
        For j = restoreAt + 7 To hi - 2
            If b(j) = &HC2 Then retN = b(j + 1) + b(j + 2) * &H100&: Exit For
            If b(j) = &HC3 Then Exit For
        Next j
        If retN >= 4 Then NativeColPut gRetbufFunc, "V" & addr, CStr(retN)
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

Private Function NativeIsVBGlobalDesc(ByVal descVA As Long) As Boolean
    'True when a __vbaNew2 object descriptor creates the VB6 _Global intrinsic-objects
    'holder (the thing whose vtable carries .App/.Screen/.Clipboard at 0x14/0x18/0x1C).
    'The descriptor's CLSID/IID pointers (at +4 / +8) are the VB OLB family
    '{FCFB3D2x-A0FA-1068-A738-08002B3371B5}: FCFB3D23 = Global coclass, FCFB3D22 =
    '_Global interface.  Check the GUID Data1 at either pointer.
    Dim p As Long, d1 As Long, off As Long
    On Error Resume Next
    If descVA < OptHeader.ImageBase Then Exit Function
    For off = 4 To 8 Step 4
        p = NativeFileDword(descVA + off)
        If p >= OptHeader.ImageBase Then
            d1 = NativeFileDword(p)
            If d1 = &HFCFB3D22 Or d1 = &HFCFB3D23 Then NativeIsVBGlobalDesc = True: Exit Function
        End If
    Next
End Function

Private Function NativeGlobalObjByOffset(ByVal disp As Long) As String
    'VB6 form-interface vtable slots for the intrinsic global objects.
    Select Case disp
        Case &H14: NativeGlobalObjByOffset = "App"
        Case &H18: NativeGlobalObjByOffset = "Screen"
        Case &H1C: NativeGlobalObjByOffset = "Clipboard"
    End Select
End Function

Private Function NativeGlobalMethodByOffset(ByVal disp As Long) As String
    'VB6 global STATEMENTS routed through the standalone _Global object's vtable
    '(the lazily-New'd intrinsic-objects holder).  Verified: `Unload <form>` compiles
    'to `call [_GlobalVt + 0x10]` with the form as the argument.  Extend as more global
    'routines are confirmed by tracing.  Verified against Dungeon Form_Load:
    'Load frmMainMenu = [+0xC], Unload <form> = [+0x10] (adjacent _Global slots).
    Select Case disp
        Case &HC: NativeGlobalMethodByOffset = "Load"
        Case &H10: NativeGlobalMethodByOffset = "Unload"
    End Select
End Function

Private Function NativeFormMethodByOffset(ByVal disp As Long) As String
    'Built-in VB6 _Form-interface methods at fixed vtable offsets, below the
    'control-accessor block (0x2F8) and the user-method block (0x6F8).  These are
    'part of the form runtime interface and stable across programs.  Verified by
    'tracing real forms (Dungeon frmMainMenu cmdExit/cmdLoad/cmdNew all
    '`call [Me_vt + 0x2B4]`) and cross-checking the commercial decompiler's Me.Hide.
    'Argument-less methods (Hide) render as a bare `Me.<method>`; an arg-taking
    'method (PopupMenu <menu>, see NativeFormMethodHasArg) renders `<method> <arg>`.
    'Extend as more offsets are confirmed by tracing.
    Select Case disp
        Case &H2B0: NativeFormMethodByOffset = "Show"
        Case &H2B4: NativeFormMethodByOffset = "Hide"
        Case &H2BC: NativeFormMethodByOffset = "PopupMenu"
    End Select
End Function

Private Function NativeFormMethodHasArg(ByVal disp As Long) As Boolean
    'A _Form method that takes a leading argument (pushed below the implicit `this`),
    'so the handler renders `<method> <arg>` (e.g. `PopupMenu mnuCarry`) rather than a
    'bare `Me.<method>`.  PopupMenu's extra optional args (flags/x/y) are not modelled.
    Select Case disp
        Case &H2BC: NativeFormMethodHasArg = True   'PopupMenu <menu>
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

Private Function NativeIntrinsicMethodByOffset(ByVal obj As String, ByVal disp As Long) As String
    'Value-returning METHOD vtable offsets on a VB6 intrinsic object - distinct from the
    'read-only property getters in NativeIntrinsicPropByOffset.  The Clipboard object
    '(reached via the standalone _Global object's .Clipboard accessor at [_Global+0x1C])
    'exposes GetData/GetText/GetFormat at fixed offsets in its runtime interface; verified
    'on the class-example EXE (Clipboard.GetData = call [clipboardVt + 0x54]).  Extend as
    'more offsets are observed.
    Select Case obj
        Case "Clipboard"
            Select Case disp
                Case &H54: NativeIntrinsicMethodByOffset = "GetData"
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
            Case &H8D                          'lea r, [mem]
                'lea into eax CLOBBERS the call result (the address overwrites it).
                'Without this a deferred call before `lea eax,[ebp-X]; push eax` was
                'wrongly kept until the push folded it as if push consumed the result
                '(dropping the call - e.g. the user call in Timer2_Timer).
                If reg = 0 Then NativeEaxUse = 2: Exit Function
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
            Dim rmIsMe As Boolean
            rmIsMe = False
            If bse >= 0 And bse <= 7 Then rmIsMe = (NVHasMe And NVRegIsMe(bse))
            If rmIsMe And disp > 0 And disp < &H2000 Then
                'Read of an instance FIELD of Me (this/self) -> field_<off> (or its real
                'typeinfo name), MIRRORING the field STORE side (NativeFieldStoreLHS) so a
                'read and write of the same field match: an Integer Property Get/Function
                'reads `ID = field_34` / `field_34 * 1000` to pair with its Let's
                '`field_34 = vData` and the `Private field_34` declaration, not the bare
                'Me-deref arg_8(52).  Scoped to THIS operand/16-bit-shadow path; the 32-bit
                'NVReg deref path is unchanged, so struct/array deref chains and the
                'string-concat operands that depend on the arg_8(N) form are unaffected.
                NativeRmVal = NativeFieldName(disp)
            ElseIf bse >= 0 And bse <= 7 And NativeIsDerefBase(NVReg(bse)) Then
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
                ElseIf Len(pv) > 0 Then
                    'Element index register is untracked (a computed/scaled byte offset
                    'we did not model, so NVReg(idx) is empty).  If the pre-pass
                    'recovered the LOGICAL index for this access, render
                    'global_X(12)(i)(off); otherwise mirror the STORE side, which drops
                    'the index (NativeFieldStoreLHS -> base(off)).  Either way beats a
                    'blank <cond> and is consistent with our own field stores.
                    Dim eix As String
                    eix = NativeColGet(NVElemIdx, "E" & NVCurVa)
                    If Len(eix) > 0 Then
                        If disp = 0 Then NativeRmVal = pv & "(" & eix & ")" Else NativeRmVal = pv & "(" & eix & ")(" & CStr(disp) & ")"
                    Else
                        If disp = 0 Then NativeRmVal = pv Else NativeRmVal = pv & "(" & CStr(disp) & ")"
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
    'A direct-test StrCmp (`call __vbaStrCmp; test eax,eax; jcc`): the preceding StrCmp
    'handler stashed the operands.  `test strcmp,strcmp` has the SAME Jcc polarity as
    '`cmp p1,p2` (je=equal, jl=p1<p2, ...), so feed the operands as a real compare and
    'the Jcc renders `If var_124 = "!"`.  Consumed by the immediately-following
    'test eax,eax; cleared otherwise so a later unrelated compare never reuses it.
    If NVStrCmpPending Then
        NVStrCmpPending = False
        If op = &H85 And md = 3 And reg = 0 And rm = 0 Then
            NVCmpL = NVStrCmpP1: NVCmpR = NVStrCmpP2
            NVCmpIsTest = False: NVCmpIsBool = False: NVCmpSet = True
            Exit Sub
        End If
    End If
    'test reg,reg where the register holds a recovered EXPRESSION (a relational
    'Boolean from an fp comparison, or a folded value): test it for non-zero.
    'This is the standard "If <fp compare> Then" tail (test ax,ax / test eax,eax)
    'and resolves what the 0x66-register guard below would otherwise drop.
    If op = &H85 And md = 3 And reg = rm Then
        Dim bv As String
        bv = NVReg(rm)
        If Len(bv) > 0 And (Left$(bv, 1) = "(" Or NativeIsCallExpr(bv)) Then
            If NativeLooksRelational(bv) Then
                'Already a relational Boolean (recovered fp compare) - render it
                'directly, without the redundant "<> 0".
                NVCmpL = bv: NVCmpIsBool = True: NVCmpSet = True
            Else
                'A folded value or predicate call (e.g. an arithmetic sum, or
                'IsNumeric(x)/EOF(f) word-tested as a VARIANT_BOOL) - test for non-zero.
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
            Case &H3D                                     'cmp ax, imm16 (16-bit accumulator)
                'A `Select Case <IntegerField> ... Case lo To hi` range check and other
                'Integer/Boolean compares load the value via `mov ax,word[mem]` then
                '`cmp ax,imm16`.  The generic decode bails on the 16-bit accumulator, so
                'use the word captured from the preceding memory load (NVR16Val) -> the
                'operand resolves (e.g. global_X(12)(2) >= 1) instead of leaving <cond>.
                Dim axw As String
                If Len(NVR16Val(0)) > 0 Then
                    axw = NVR16Val(0)
                ElseIf NativeIsCallExpr(NVReg(0)) Then
                    'ax holds an Integer-returning function result tested against a
                    'literal (`If modMap_Dist(...) <= 1`); resolve like the `test ax,ax`
                    'of a folded predicate call (9ab06f5) instead of dropping to <cond>.
                    axw = NVReg(0)
                End If
                If Len(axw) > 0 Then
                    NVCmpL = axw: NVCmpR = CStr(NativeDumpInt16(dump, i + 1))
                    NVCmpIsTest = False: NVCmpIsBool = False: NVCmpSet = True
                End If
                GoTo done3
            Case &HA9: GoTo done3                         'test ax, imm - 16-bit accumulator
            Case Else
                If md = 3 Then
                    'A 16-bit reg-reg compare is normally VB's Boolean juggling (low-word
                    'partials our 32-bit model can't follow) - resolving it yields garbage,
                    'so bail.  EXCEPT a `cmp r16,r16` whose BOTH registers hold a clean
                    'tracked named value or a numeric constant: such a value was loaded by
                    'a full 32-bit mov / xor-to-zero (a 16-bit write clears the tracked
                    'value), so the low word IS the Integer value, e.g.
                    '`mov ecx,[var_58]; xor eax,eax; cmp cx,ax` = `XXrun = 0`.  Same safety
                    'basis as the `test si,si` resolution (89c7762); falls through to decode.
                    Dim okRR As Boolean
                    If op = &H3B Or op = &H39 Then
                        okRR = (NativeIsCleanNamedVal(NVReg(reg)) Or NativeIsNumLit(NVReg(reg))) _
                           And (NativeIsCleanNamedVal(NVReg(rm)) Or NativeIsNumLit(NVReg(rm)))
                    End If
                    If Not okRR Then GoTo done3               'r/m is a register
                End If
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
    Dim rva As Long, k As Long, secEnd As Long
    If va < OptHeader.ImageBase Then Exit Function
    rva = va - OptHeader.ImageBase
    For k = 0 To MAXSECTIONS
        If SecHeader(k).SizeRawData > 0 And SecHeader(k).Address > 0 Then
            'Bound by the section's VIRTUAL size (Misc), not SizeRawData: VB6 module
            'globals/form instances live in BSS - the uninitialised tail of .data that
            'is allocated in memory (VirtualSize) but NOT stored on disk (SizeRawData),
            'so a SizeRawData bound wrongly rejects them (e.g. global_0055D5A0), leaving
            'their method calls untyped (global.UnkVCall_<off> instead of global.Member).
            secEnd = SecHeader(k).SizeRawData
            If SecHeader(k).Misc > secEnd Then secEnd = SecHeader(k).Misc
            If rva >= SecHeader(k).Address And rva < SecHeader(k).Address + secEnd Then
                NativeIsGlobalAddr = ((Int(SecHeader(k).Properties / &H20000000) Mod 2) = 0)
                Exit Function
            End If
        End If
    Next
End Function

Private Function NativeCallTargetName(ByVal tgt As Long) As String
    'Resolve a call target address to a procedure name, qualified by module
    'unless it is in the current form/module (then unqualified).
    Dim nm As String
    nm = NativeLookupName(tgt)
    If Len(nm) = 0 Then nm = NativeLookupName(NativeSnapEntry(tgt))
    If Len(nm) = 0 Then
        'Not a real (linked) name - prefer a declared-DLL (Win32 API) call thunk.
        Dim apinm As String
        apinm = NativeApiStubName(tgt)
        If Len(apinm) > 0 Then NativeCallTargetName = apinm: Exit Function
        'A discovered user procedure with no linked name (e.g. a frameless private
        'module Function): use its owner-qualified synthetic name from
        'gNativeProcArray (Module1.proc_403100) so a cross-module call is traceable
        'to the file the procedure lives in.
        nm = NativeProcArrayName(tgt)
        If Len(nm) = 0 Then nm = NativeProcArrayName(NativeSnapEntry(tgt))
        If Len(nm) = 0 Then
            NativeCallTargetName = "proc_" & Hex$(tgt)
            Exit Function
        End If
    End If
    If Len(NVForm) > 0 And Left$(nm, Len(NVForm) + 1) = NVForm & "." Then nm = Mid$(nm, Len(NVForm) + 2)
    NativeCallTargetName = nm
End Function

Private Function NativeProcArrayName(ByVal tgt As Long) As String
    'Owner-qualified synthetic name (e.g. "Module1.proc_403100") of the discovered
    'procedure whose entry is at or just before tgt, from gNativeProcArray - which
    'records the owning module for every procedure including those that have no
    'linked real name in SubNamelist.  Empty when tgt is not a known proc entry.
    Dim i As Long, d As Long, bestDelta As Long
    On Error Resume Next
    bestDelta = 99999
    For i = 0 To UBound(gNativeProcArray) - 1
        If gNativeProcArray(i).offset <> 0 Then
            d = tgt - gNativeProcArray(i).offset
            If d >= 0 And d <= 24 And d < bestDelta Then
                bestDelta = d
                NativeProcArrayName = gNativeProcArray(i).sName
            End If
        End If
    Next
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

'---------------------------------------------------------------------------
' UDT (record) recovery from VB6 record-layout descriptors
'
' __vbaRecAssign / __vbaRecDestruct / __vbaRec*ToUni take the address of a
' record-layout descriptor (emitted for any UDT containing reference-type fields -
' String/Variant/Object - that the runtime must deep-copy or free).  The descriptor
' encodes the struct SIZE and the OFFSET+TYPE of each reference field.  Field NAMES
' are stripped, so fields render field_<hexOffset> with their recovered type and the
' numeric gaps are filled with Long/Integer/Byte so the byte layout is exact.  This
' BEATS the commercial decompiler, which punts to a single `bStruc(N) As Byte`.
'
' Descriptor format (verified across 5 records: Dungeon MapType/MessageType + the
' UDT sample TOneString/TMixed/TVariant):
'   +2  WORD cbStruct   (total struct size)
'   +6  WORD cFields    (count of reference-type fields)
'   +11 BYTE kind       (0x2C = simple BSTR/Variant field table)
'   +12.. cFields * { WORD fieldOffset, WORD typecode }   (typecode 0x0001 = String)
' Object/array/nested records use other kinds (0x0C/0x3C) with longer, variable-size
' entries; those are not fully decoded (we bail rather than emit a wrong layout).
'---------------------------------------------------------------------------

Private Sub NativeHarvestUDTArgs()
    'Scan the pending call arguments for a record-layout descriptor address and
    'register any that decode.  Safe to call on any record-helper call; non-numeric
    'or out-of-range args are ignored, and the decode itself validates the structure.
    Dim i As Long, s As String, v As Long
    On Error Resume Next
    For i = 0 To NVPushTop - 1
        s = Trim$(NVPushImm(i))
        If Len(s) > 0 And IsNumeric(s) Then
            v = CLng(s)
            If v > OptHeader.ImageBase Then NativeRegisterUDT v
        End If
    Next
End Sub

Public Sub NativeRegisterUDT(ByVal va As Long)
    'Decode the descriptor at va once and store its full `Public Type ... End Type`
    'block in gUDTDesc, keyed by VA for dedup.  No-op if va is not a clean descriptor.
    Dim key As String, body As String, tmp As String, typeName As String
    If va < OptHeader.ImageBase Then Exit Sub
    If gUDTDesc Is Nothing Then Set gUDTDesc = New Collection
    key = "H" & Right$("00000000" & Hex$(va), 8)
    On Error Resume Next
    tmp = "": tmp = gUDTDesc.Item(key)
    On Error GoTo 0
    If Len(tmp) > 0 Then Exit Sub                     'already registered
    body = NativeDecodeRecordDescriptor(va)
    If Len(body) = 0 Then Exit Sub
    typeName = "UDT_" & Right$("00000000" & Hex$(va), 8)
    'Stored WITHOUT the scope keyword - GetUDTBlock prepends Public (module) or
    'Private (form) at emit time.
    body = "Type " & typeName & vbCrLf & body & "End Type" & vbCrLf & vbCrLf
    On Error Resume Next
    gUDTDesc.Add body, key
    On Error GoTo 0
End Sub

Public Sub NativeRegisterUDTBySize(ByVal sizeStr As String, ByVal arrayPtr As String, ByVal callVa As Long)
    'Fallback for a DESCRIPTOR-LESS UDT (all fixed-string/numeric fields - nothing for
    'the runtime to deep-copy, so no __vbaRec* descriptor): recover only the SIZE (a
    'dynamic UDT array's element size from __vbaRedim) and emit a byte-buffer Type
    '`bStruc(1 To N) As Byte` - the commercial decompiler's ceiling (no field types/names
    'without a descriptor).  Gated to UDT-plausible sizes (>16 excludes every primitive
    'incl. Variant).  Named/keyed by the array variable's ADDRESS (so it reads like the
    'descriptor UDTs UDT_<va> and dedups per array), falling back to the ReDim call site
    'when the array is a form field (no module-level VA).  The size is kept in the name.
    Dim n As Long, key As String, typeName As String, body As String, tmp As String
    Dim gva As String, foff As String, arrId As String, addrTok As String
    If Not IsNumeric(Trim$(sizeStr)) Then Exit Sub
    n = CLng(Trim$(sizeStr))
    If n <= 16 Or n > 65535 Then Exit Sub
    If gUDTDesc Is Nothing Then Set gUDTDesc = New Collection
    'Array identity (matches NativeDetectUDTStringFields): "G"&globalVA for a module
    'global, "F"&MeFieldOffset for a form field.  The display address in the name is the
    'global VA, else the ReDim call site (a real code address).
    gva = NativeExtractGlobalHex(arrayPtr)
    If Len(gva) > 0 Then
        arrId = "G" & gva: addrTok = gva
    Else
        foff = NativeExtractFieldOffset(arrayPtr)
        addrTok = Right$("00000000" & Hex$(callVa), 8)
        If Len(foff) > 0 Then arrId = "F" & foff Else arrId = "C" & Hex$(callVa)
    End If
    key = "U" & arrId
    On Error Resume Next
    tmp = "": tmp = gUDTDesc.Item(key)
    On Error GoTo 0
    If Len(tmp) > 0 Then Exit Sub
    typeName = "UDT_" & addrTok & "_" & n & "Bytes"
    body = "Type " & typeName & vbCrLf & _
           NativeRenderUDTBody(n, NativeColGet(gUDTStrFields, arrId)) & _
           "End Type" & vbCrLf & vbCrLf
    On Error Resume Next
    gUDTDesc.Add body, key
    On Error GoTo 0
End Sub

Private Function NativeRenderUDTBody(ByVal n As Long, ByVal fieldsStr As String) As String
    'Render a byte-buffer UDT body, placing each recovered fixed-string field
    '(`field_<off> As String * <len>`) at its offset and byte-padding the gaps.  Gaps are
    'computed from the NEXT detected offset, and a string field advances by 2*len (its
    'in-memory Unicode size), so the total stays exactly n even when a field is really an
    'array (its remainder is absorbed into the following byte gap).  Falls back to a single
    'bStruc(1 To n) when there are no string fields or the layout is inconsistent.
    Dim parts() As String, offs() As Long, kinds() As String, prm() As Long, c As Long, k As Long, m As Long
    Dim res As String, pos As Long, p1 As Long, p2 As Long, decl As String, adv As Long
    If Len(fieldsStr) = 0 Then NativeRenderUDTBody = "    bStruc(1 To " & n & ") As Byte" & vbCrLf: Exit Function
    parts = Split(fieldsStr, ";")
    ReDim offs(UBound(parts)): ReDim kinds(UBound(parts)): ReDim prm(UBound(parts)): c = 0
    For k = 0 To UBound(parts)
        p1 = InStr(parts(k), ":")                                'off : kind : param
        If p1 > 0 Then
            p2 = InStr(p1 + 1, parts(k), ":")
            If p2 > 0 Then
                offs(c) = CLng(Left$(parts(k), p1 - 1))
                kinds(c) = Mid$(parts(k), p1 + 1, p2 - p1 - 1)
                prm(c) = CLng(Mid$(parts(k), p2 + 1))
                c = c + 1
            End If
        End If
    Next
    If c = 0 Then NativeRenderUDTBody = "    bStruc(1 To " & n & ") As Byte" & vbCrLf: Exit Function
    For k = 0 To c - 2                                           'sort by offset
        For m = 0 To c - 2 - k
            If offs(m) > offs(m + 1) Then
                Dim t1 As Long, ts As String
                t1 = offs(m): offs(m) = offs(m + 1): offs(m + 1) = t1
                ts = kinds(m): kinds(m) = kinds(m + 1): kinds(m + 1) = ts
                t1 = prm(m): prm(m) = prm(m + 1): prm(m + 1) = t1
            End If
        Next
    Next
    pos = 0
    For k = 0 To c - 1
        If offs(k) < pos Then GoTo inconsistent
        If offs(k) > pos Then
            res = res & "    field_" & Hex$(pos) & "(1 To " & (offs(k) - pos) & ") As Byte" & vbCrLf
            pos = offs(k)
        End If
        Select Case kinds(k)
            Case "S": decl = "String * " & prm(k): adv = prm(k) * 2     'in-memory Unicode
            Case "I": decl = "Integer": adv = 2
            Case "L": decl = "Long": adv = 4
            Case "G": decl = "Single": adv = 4
            Case "D": decl = "Double": adv = 8
            Case Else: GoTo inconsistent
        End Select
        res = res & "    field_" & Hex$(pos) & " As " & decl & vbCrLf
        pos = pos + adv
        If pos > n Then GoTo inconsistent
    Next
    If pos < n Then res = res & "    field_" & Hex$(pos) & "(1 To " & (n - pos) & ") As Byte" & vbCrLf
    NativeRenderUDTBody = res
    Exit Function
inconsistent:
    NativeRenderUDTBody = "    bStruc(1 To " & n & ") As Byte" & vbCrLf
End Function

Private Function NativeExtractFieldOffset(ByVal s As String) As String
    'A form-field array's __vbaRedim arrayPtr renders "(arg_8 + NN)" (NN decimal = the Me
    'field offset); return NN as hex (matching the [esi+disp] decode), else "".
    Dim p As Long, i As Long, ch As String, num As String
    p = InStr(s, "arg_8 + ")
    If p = 0 Then Exit Function
    i = p + 8
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then num = num & ch Else Exit Do
        i = i + 1
    Loop
    If Len(num) > 0 Then NativeExtractFieldOffset = Hex$(CLng(num))
End Function

Private Function NativeExtractGlobalHex(ByVal s As String) As String
    'Extract the 8-hex address from a `global_XXXXXXXX` token (the array variable's VA),
    'or "" if s holds no global reference (e.g. a form field `(arg_8 + 52)`).
    Dim p As Long, i As Long, ch As Long, h As String
    p = InStr(s, "global_")
    If p = 0 Then Exit Function
    For i = p + 7 To p + 7 + 7
        If i > Len(s) Then Exit Function
        ch = Asc(UCase$(Mid$(s, i, 1)))
        If Not ((ch >= 48 And ch <= 57) Or (ch >= 65 And ch <= 70)) Then Exit Function
        h = h & Mid$(s, i, 1)
    Next
    If Len(h) = 8 Then NativeExtractGlobalHex = UCase$(h)
End Function

Public Function GetUDTBlock(Optional ByVal scope As String = "Public") As String
    'Concatenate every recovered UDT Type block (insertion order), prefixing each with
    'the requested scope - Public in a standard module, Private in a form (forms cannot
    'host Public Types).  Emitted once (the EXE does not record the owning module).
    Dim s As String, v As Variant
    On Error Resume Next
    If gUDTDesc Is Nothing Then Exit Function
    For Each v In gUDTDesc
        s = s & scope & " " & CStr(v)
    Next
    GetUDTBlock = s
End Function

Private Function NativeDecodeRecordDescriptor(ByVal va As Long) As String
    'Parse a simple-table (kind 0x2C) record descriptor at va and render its field
    'body (the lines between `Type` and `End Type`).  Returns "" if va is not a clean
    'all-String simple-table descriptor (so unknown/exotic layouts are left alone).
    Dim fp As Integer, rva As Long
    Dim cbStruct As Long, cFields As Long, kind As Long
    Dim i As Long, off As Long, tc As Long
    Dim offs() As Long, nf As Long
    On Error GoTo bad
    rva = va - OptHeader.ImageBase
    If rva < 5 Then Exit Function
    fp = FreeFile
    Open SFilePath For Binary Access Read As #fp
    cbStruct = NativeGetWord(fp, rva + 2)
    cFields = NativeGetWord(fp, rva + 6)
    kind = NativeGetByte(fp, rva + 11)
    If cbStruct < 1 Or cbStruct > 8192 Then GoTo bad
    If cFields < 1 Or cFields > 256 Then GoTo bad      'must have >=1 reference field
    If kind <> &H2C Then GoTo bad                       'only the simple BSTR/Variant table
    ReDim offs(cFields - 1)
    nf = 0
    For i = 0 To cFields - 1
        off = NativeGetWord(fp, rva + 12 + i * 4)
        tc = NativeGetWord(fp, rva + 12 + i * 4 + 2)
        If off < 0 Or off >= cbStruct Then GoTo bad
        If tc <> 1 Then GoTo bad                         'String only; Variant entries are longer - bail
        offs(nf) = off: nf = nf + 1
    Next
    Close fp
    'Walk 0..cbStruct emitting each String at its offset and Long/Integer/Byte filler
    'for the numeric gaps, so every field offset is byte-accurate.
    Dim res As String, pos As Long, j As Long, isStr As Boolean, nextStr As Long, gap As Long
    pos = 0
    Do While pos < cbStruct
        isStr = False
        For j = 0 To nf - 1
            If offs(j) = pos Then isStr = True: Exit For
        Next
        If isStr Then
            res = res & "    field_" & Hex$(pos) & " As String" & vbCrLf
            pos = pos + 4
        Else
            nextStr = cbStruct
            For j = 0 To nf - 1
                If offs(j) > pos And offs(j) < nextStr Then nextStr = offs(j)
            Next
            gap = nextStr - pos
            Do While gap >= 4
                res = res & "    field_" & Hex$(pos) & " As Long" & vbCrLf
                pos = pos + 4: gap = gap - 4
            Loop
            If gap = 2 Then
                res = res & "    field_" & Hex$(pos) & " As Integer" & vbCrLf
                pos = pos + 2
            ElseIf gap = 1 Then
                res = res & "    field_" & Hex$(pos) & " As Byte" & vbCrLf
                pos = pos + 1
            ElseIf gap = 3 Then
                res = res & "    field_" & Hex$(pos) & "(2) As Byte" & vbCrLf
                pos = pos + 3
            End If
        End If
    Loop
    NativeDecodeRecordDescriptor = res
    Exit Function
bad:
    On Error Resume Next
    Close fp
End Function

Private Function NativeGetWord(ByVal fp As Integer, ByVal rva As Long) As Long
    'Read a little-endian unsigned WORD at file offset rva (0-based RVA; raw==RVA for
    'VB6 EXEs, as NativeStringAt relies on).  Binary Get positions are 1-based.
    Dim w As Integer, v As Long
    Get #fp, rva + 1, w
    v = w: If v < 0 Then v = v + 65536
    NativeGetWord = v
End Function

Private Function NativeGetByte(ByVal fp As Integer, ByVal rva As Long) As Long
    Dim b As Byte
    Get #fp, rva + 1, b
    NativeGetByte = b
End Function

Public Sub NativeDetectUDTStringFields()
    'Recover fixed-length string FIELDS of descriptor-less UDTs by scanning the code for
    '__vbaStrFixstr / __vbaLsetFixstr calls.  Each compiles to a fixed idiom:
    '   mov  R1, [arrayBase]        ; the UDT array (Me+fieldOff, or a global)
    '   mov  R2, [R1 + 0x0C]        ; SAFEARRAY pvData (struct offset 0x0C)
    '   lea  R3, [R2 + idx + disp]  ; field address - disp = the field's byte offset
    '   push <len>                  ; the fixed-string length -> As String * len
    '   call __vbaStrFixstr/Lset
    'so the field offset + length + owning array are all recoverable.  Populates
    'gUDTStrFields (arrayId -> "off:len;...") keyed the same way NativeRegisterUDTBySize
    'keys its byte-buffer UDTs, so the typed fields drop into the right Type at emit.
    Dim fp As Integer, sz As Long, i As Long, tgt As Long, nm As String
    Dim buf() As Byte, cache As Collection
    Set gUDTStrFields = New Collection
    Set cache = New Collection
    'This pre-pass runs before the first proc decompile lazily creates dsmNative.
    If dsmNative Is Nothing Then Set dsmNative = New CDisassembler
    On Error GoTo done
    fp = FreeFile
    Open SFilePath For Binary Access Read As #fp
    sz = LOF(fp)
    If sz < &H1008 Then Close fp: Exit Sub
    ReDim buf(sz - 1)
    Get #fp, 1, buf
    Close fp
    For i = &H1000 To sz - 8
        If buf(i) = &HFF And buf(i + 1) = &H15 Then
            tgt = NativeBufDword(buf, i + 2)
            Dim ck As String, isFix As Long
            ck = "T" & tgt: isFix = -1
            On Error Resume Next
            isFix = cache.Item(ck)
            On Error GoTo done
            If isFix = -1 Then
                nm = ""
                If tgt >= OptHeader.ImageBase Then nm = dsmNative.GetApiByIatVa(tgt)
                isFix = IIf(InStr(nm, "Fixstr") > 0 Or InStr(nm, "FixStr") > 0, 1, 0)
                cache.Add isFix, ck
            End If
            If isFix = 1 Then NativeDecodeFixstrCall buf, i
        End If
    Next
    'Second pass: NUMERIC fields.  Anchor on the array+pvData idiom
    '(mov R1,[arrayBase]; mov R2,[R1+0x0C]) and classify the following element access by
    'width: word (0x66)->Integer, dword->Long, fld dword->Single, fld qword->Double.
    Dim r1 As Long, r2 As Long, aLen As Long, j As Long, aId As String
    For i = &H1000 To sz - 20
        aId = NativeArrBaseAt(buf, i, r1, aLen)
        If Len(aId) > 0 Then
            For j = i + aLen To i + aLen + 10
                If buf(j) = &H8B And (buf(j + 1) And &HC0) = &H40 And (buf(j + 1) And 7) = r1 And buf(j + 2) = &HC And ((buf(j + 1) \ 8) And 7) <> 4 Then
                    r2 = (buf(j + 1) \ 8) And 7        'mov R2,[R1+0x0C] (pvData) - R2 != esp
                    NativeDecodeNumField buf, j + 3, r2, aId
                    Exit For
                End If
                If buf(j) = &HFF Or buf(j) = &HE8 Or buf(j) = &HC3 Then Exit For   'a call/ret ends the idiom
            Next
        End If
    Next
    Exit Sub
done:
    On Error Resume Next
    Close fp
End Sub

Private Function NativeArrBaseAt(buf() As Byte, ByVal i As Long, ByRef destReg As Long, ByRef instLen As Long) As String
    'If buf(i) starts a "mov R1,[arrayBase]" - [Me+off] or [global] - return the arrayId
    '("F"&MeFieldOffset / "G"&globalVA), the dest register and the instruction length.
    Dim md As Long
    If buf(i) = &H8B Then
        md = buf(i + 1)
        destReg = (md \ 8) And 7
        If (md And &HC7) = &H46 Then NativeArrBaseAt = "F" & Hex$(buf(i + 2)): instLen = 3: Exit Function          'mov r,[esi+disp8]
        If (md And &HC7) = &H47 Then NativeArrBaseAt = "F" & Hex$(buf(i + 2)): instLen = 3: Exit Function          'mov r,[edi+disp8]
        If (md And &HC7) = &H86 Then NativeArrBaseAt = "F" & Hex$(NativeBufDword(buf, i + 2)): instLen = 6: Exit Function   'mov r,[esi+disp32]
        If (md And &HC7) = &H5 Then NativeArrBaseAt = "G" & Right$("00000000" & Hex$(NativeBufDword(buf, i + 2)), 8): instLen = 6: Exit Function   'mov r,[abs]
    ElseIf buf(i) = &HA1 Then
        destReg = 0: NativeArrBaseAt = "G" & Right$("00000000" & Hex$(NativeBufDword(buf, i + 1)), 8): instLen = 5   'mov eax,[moffs32]
    End If
End Function

Private Sub NativeDecodeNumField(buf() As Byte, ByVal startOff As Long, ByVal r2 As Long, ByVal arrId As String)
    'From just after the pvData load, find the first SIB-indexed element access whose base
    'is r2 (or r3 after a lea r3,[r2+idx+disp]) and record its offset+type.  Requires a
    'real typed access (mov/cmp/fld/grp1), so a string field's lea-then-push-Fixstr is not
    'misread as numeric (no typed access follows on the lea'd register).
    Dim p As Long, tgt As Long, off As Long, haveOff As Boolean, w16 As Boolean, op As Long, md As Long, sib As Long
    tgt = r2: off = -1: haveOff = False
    For p = startOff To startOff + 22
        w16 = False
        op = buf(p)
        If op = &H66 Then w16 = True: p = p + 1: op = buf(p)         'operand-size prefix -> word
        If op = &H8D Then                                            'lea r3,[r2+idx+disp] - follow r3, take disp
            md = buf(p + 1)
            If (md And 7) = 4 Then
                sib = buf(p + 2)
                If (sib And 7) = tgt Then
                    tgt = (md \ 8) And 7
                    If (md And &HC0) = &H40 Then off = buf(p + 3): haveOff = True
                    If (md And &HC0) = &H80 Then off = NativeBufDword(buf, p + 3): haveOff = True
                End If
            End If
        ElseIf op = &H8B Or op = &H89 Or op = &H3B Or op = &H39 Or op = &H83 Then  'mov/cmp/grp1 r/m
            md = buf(p + 1)
            If (md And 7) = 4 Then
                sib = buf(p + 2)
                If (sib And 7) = tgt Then
                    If Not haveOff Then
                        If (md And &HC0) = &H40 Then off = buf(p + 3): haveOff = True
                        If (md And &HC0) = &H80 Then off = NativeBufDword(buf, p + 3): haveOff = True
                        If (md And &HC0) = &H0 Then off = 0: haveOff = True
                    End If
                    If haveOff And off >= 0 And off < 65536 Then NativeAddUDTField arrId, off, IIf(w16, "I:2", "L:4")
                    Exit Sub
                End If
            End If
        ElseIf op = &HD9 Or op = &HDD Then                          'fld/fstp dword(D9)/qword(DD)
            md = buf(p + 1)
            If (md And 7) = 4 Then
                sib = buf(p + 2)
                If (sib And 7) = tgt Then
                    If Not haveOff Then
                        If (md And &HC0) = &H40 Then off = buf(p + 3): haveOff = True
                        If (md And &HC0) = &H80 Then off = NativeBufDword(buf, p + 3): haveOff = True
                        If (md And &HC0) = &H0 Then off = 0: haveOff = True
                    End If
                    If haveOff And off >= 0 And off < 65536 Then NativeAddUDTField arrId, off, IIf(op = &HD9, "G:4", "D:8")
                    Exit Sub
                End If
            End If
        ElseIf op = &HFF Or op = &HE8 Or op = &HC3 Then
            Exit Sub                                                 'call/ret - give up
        End If
    Next
End Sub

Private Sub NativeDecodeFixstrCall(ByRef buf() As Byte, ByVal callOff As Long)
    'Back-scan the <=32 bytes before a Fixstr call for its length (push), field offset
    '(lea disp) and owning array (the mov from [Me+off] / [global]).  All three required.
    Dim j As Long, ln As Long, off As Long, arrId As String
    Dim gotLn As Boolean, gotOff As Boolean, gotArr As Boolean, md As Long, pv As Long
    ln = -1: off = -1: arrId = ""
    For j = callOff - 1 To callOff - 32 Step -1
        If j < 1 Then Exit For
        If Not gotLn And buf(j) = &H6A Then ln = buf(j + 1): gotLn = True
        If Not gotLn And buf(j) = &H68 Then
            pv = NativeBufDword(buf, j + 1)
            If pv > 0 And pv < 32768 Then ln = pv: gotLn = True   'a real length (not a string-literal ptr)
        End If
        If Not gotOff And buf(j) = &H8D Then                      'lea reg,[base+idx+disp]
            md = buf(j + 1)
            If (md And &H7) = 4 Then                              'SIB form
                If (md And &HC0) = &H40 Then off = buf(j + 3): gotOff = True
                If (md And &HC0) = &H80 Then off = NativeBufDword(buf, j + 3): gotOff = True
            End If
        End If
        If Not gotArr And buf(j) = &H8B Then                      'mov reg,[Me+off] / [global]
            md = buf(j + 1)
            If (md And &HC7) = &H46 Then arrId = "F" & Hex$(buf(j + 2)): gotArr = True            'mod01 rm110(esi) disp8
            If (md And &HC7) = &H47 Then arrId = "F" & Hex$(buf(j + 2)): gotArr = True            'mod01 rm111(edi) disp8
            If (md And &HC7) = &H86 Then arrId = "F" & Hex$(NativeBufDword(buf, j + 2)): gotArr = True   'mod10 rm110 disp32
            If (md And &HC7) = &H5 Then arrId = "G" & Right$("00000000" & Hex$(NativeBufDword(buf, j + 2)), 8): gotArr = True  'mod00 rm101 [abs]
        End If
        If Not gotArr And buf(j) = &HA1 Then                      'mov eax,[moffs32] (global)
            arrId = "G" & Right$("00000000" & Hex$(NativeBufDword(buf, j + 1)), 8): gotArr = True
        End If
    Next
    If gotLn And gotOff And gotArr And ln > 0 And ln < 32768 And off >= 0 And off < 65536 Then
        NativeAddUDTField arrId, off, "S:" & ln
    End If
End Sub

Private Sub NativeAddUDTField(ByVal arrId As String, ByVal off As Long, ByVal tail As String)
    'Append "off:tail" to gUDTStrFields(arrId), skipping a duplicate offset.  tail encodes
    'the kind+param: "S:<len>" (String * len), "I:2" (Integer), "L:4" (Long), "G:4"
    '(Single), "D:8" (Double).  First detection of an offset wins (strings are scanned
    'before numerics, so a fixed string is never overwritten by a numeric guess).
    Dim cur As String
    If gUDTStrFields Is Nothing Then Set gUDTStrFields = New Collection
    cur = NativeColGet(gUDTStrFields, arrId)
    If InStr(";" & cur & ";", ";" & off & ":") > 0 Then Exit Sub      'offset already recorded
    If Len(cur) > 0 Then cur = cur & ";"
    NativeColPut gUDTStrFields, arrId, cur & off & ":" & tail
End Sub

Private Function NativeBufDword(ByRef buf() As Byte, ByVal i As Long) As Long
    'Little-endian DWORD from a byte buffer, returned as a (possibly negative) Long.
    'Computed in Double then wrapped to signed Long so a top byte >= 0x80 (which would
    'overflow the Long multiply) does not raise - the whole-file scan hits such bytes.
    If i < 0 Or i + 3 > UBound(buf) Then Exit Function
    Dim v As Double
    v = buf(i) + buf(i + 1) * 256# + buf(i + 2) * 65536# + buf(i + 3) * 16777216#
    If v >= 2147483648# Then v = v - 4294967296#
    NativeBufDword = CLng(v)
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

'---------------------------------------------------------------------------
' Frameless-function body recovery
'
' VB6 compiles a tiny routine (e.g. ReturnNumber = a*b) WITHOUT a stack frame:
' no `push ebp`, arguments read at [esp+4]/[esp+8], result returned in eax/ax.
' The general decoder assumes an ebp frame so it renders an empty body.  This is
' a small, self-contained interpreter for the common simple shapes (load params,
' deref, one or two arithmetic ops, return) - it BAILS (returns "") on anything
' it does not fully understand, so a complex frameless proc just keeps its empty
' body and a framed proc is never touched.
'---------------------------------------------------------------------------
Private Function NativeFramelessBody(col As Collection, ByRef isFunc As Boolean) As String
    On Error GoTo bail
    isFunc = False
    Dim regv(7) As String, regPtr(7) As Boolean      '0=eax 1=ecx 2=edx 3=ebx 4=esp 5=ebp 6=esi 7=edi
    Dim inst As CInstruction, dump As String, n As Long, i As Long, op As Long, op2 As Long
    For Each inst In col
        dump = UCase$(Replace(inst.dump, " ", ""))
        n = Len(dump) \ 2
        If n = 0 Then GoTo nextI
        i = NativeOpStart(dump, n)                    'past 0x66 etc.
        op = NativeDumpByte(dump, i)
        Select Case op
            Case &H90, &HCC, &H98, &H99               'nop / int3 / cwde / cdq  -> ignore
            Case &H70 To &H7F                         'short Jcc (jo/jno overflow guards) -> ignore
            Case &HEB                                 'short jmp -> ignore (skip over a guard)
            Case &H8B                                 'mov reg, r/m
                If Not NativeFLMov(dump, i + 1, regv, regPtr) Then GoTo bail
            Case &HB8 To &HBF                         'mov reg, imm32
                Dim rr As Long: rr = op - &HB8
                regv(rr) = NativeFLNum(NativeDumpInt32(dump, i + 1)): regPtr(rr) = False
            Case &H3                                  'add reg, r/m
                If Not NativeFLArith(dump, i + 1, regv, regPtr, " + ") Then GoTo bail
            Case &H2B                                 'sub reg, r/m
                If Not NativeFLArith(dump, i + 1, regv, regPtr, " - ") Then GoTo bail
            Case &H69                                 'imul reg, r/m, imm32
                If Not NativeFLImul3(dump, i + 1, regv, regPtr, False) Then GoTo bail
            Case &H6B                                 'imul reg, r/m, imm8
                If Not NativeFLImul3(dump, i + 1, regv, regPtr, True) Then GoTo bail
            Case &HF                                  'two-byte opcode
                op2 = NativeDumpByte(dump, i + 1)
                Select Case op2
                    Case &H80 To &H8F                 'near Jcc -> ignore
                    Case &HAF                         'imul reg, r/m
                        If Not NativeFLArith(dump, i + 2, regv, regPtr, " * ") Then GoTo bail
                    Case &HBE, &HBF, &HB6, &HB7       'movsx/movzx reg, r/m -> mov+deref
                        If Not NativeFLMov(dump, i + 2, regv, regPtr) Then GoTo bail
                    Case Else: GoTo bail
                End Select
            Case &HC3, &HC2                           'ret / ret N -> done
                Exit For
            Case Else
                GoTo bail                             'anything else: not a simple body
        End Select
nextI:
    Next
    'The return value lives in eax (reg 0).
    If Len(regv(0)) > 0 Then isFunc = True: NativeFramelessBody = regv(0)
    Exit Function
bail:
    NativeFramelessBody = ""
End Function

'Decode a ModRM (+SIB+disp) operand at byte `idx`.  Fills regField and the r/m.
'Returns the byte index after the operand, or -1 for a form we do not handle.
Private Function NativeFLOperand(ByVal dump As String, ByVal idx As Long, ByRef regField As Long, _
        ByRef rmIsReg As Boolean, ByRef rmReg As Long, ByRef rmDisp As Long, ByRef rmDeref As Boolean) As Long
    Dim modrm As Long, md As Long, rm As Long, e As Long
    modrm = NativeDumpByte(dump, idx)
    md = (modrm \ &H40) And 3
    regField = (modrm \ 8) And 7
    rm = modrm And 7
    e = idx + 1
    rmDisp = 0: rmDeref = False: rmIsReg = False: rmReg = -1
    If md = 3 Then rmIsReg = True: rmReg = rm: NativeFLOperand = e: Exit Function
    rmDeref = True
    If rm = 4 Then                                    'SIB
        Dim sib As Long, base As Long, idx2 As Long
        sib = NativeDumpByte(dump, e): e = e + 1
        base = sib And 7: idx2 = (sib \ 8) And 7
        If idx2 <> 4 Then NativeFLOperand = -1: Exit Function   'a scaled index -> not handled
        rmReg = base
        If md = 0 And base = 5 Then
            rmReg = -1: rmDisp = NativeDumpInt32(dump, e): e = e + 4
        ElseIf md = 1 Then
            rmDisp = NativeDumpInt8(dump, e): e = e + 1
        ElseIf md = 2 Then
            rmDisp = NativeDumpInt32(dump, e): e = e + 4
        End If
        NativeFLOperand = e: Exit Function
    End If
    If rm = 5 And md = 0 Then rmReg = -1: rmDisp = NativeDumpInt32(dump, e): e = e + 4: NativeFLOperand = e: Exit Function
    rmReg = rm
    If md = 1 Then rmDisp = NativeDumpInt8(dump, e): e = e + 1
    If md = 2 Then rmDisp = NativeDumpInt32(dump, e): e = e + 4
    NativeFLOperand = e
End Function

'Symbolic value of an r/m operand (a register's value, a parameter at [esp+N], or
'the deref of a register that holds a parameter pointer).  "" if not resolvable.
Private Function NativeFLRmVal(ByRef regv() As String, ByRef regPtr() As Boolean, _
        ByVal rmIsReg As Boolean, ByVal rmReg As Long, ByVal rmDisp As Long, ByVal rmDeref As Boolean) As String
    If rmIsReg Then
        NativeFLRmVal = regv(rmReg)
    ElseIf rmDeref Then
        If rmReg = 4 Then                             'esp-relative -> a parameter
            If rmDisp >= 4 Then NativeFLRmVal = "arg_" & Hex$(rmDisp + 4)
        ElseIf rmReg >= 0 And rmReg <= 7 Then         'deref [reg] of a parameter pointer
            If regPtr(rmReg) Then NativeFLRmVal = regv(rmReg)
        End If
    End If
End Function

Private Function NativeFLMov(ByVal dump As String, ByVal idx As Long, ByRef regv() As String, ByRef regPtr() As Boolean) As Boolean
    Dim regF As Long, rmIsReg As Boolean, rmReg As Long, rmDisp As Long, rmDeref As Boolean
    If NativeFLOperand(dump, idx, regF, rmIsReg, rmReg, rmDisp, rmDeref) = -1 Then Exit Function
    If rmIsReg Then
        regv(regF) = regv(rmReg): regPtr(regF) = regPtr(rmReg): NativeFLMov = True: Exit Function
    End If
    If rmReg = 4 Then                                 'mov reg, [esp+disp] -> POINTER to the parameter
        If rmDisp >= 4 Then regv(regF) = "arg_" & Hex$(rmDisp + 4): regPtr(regF) = True: NativeFLMov = True
        Exit Function
    ElseIf rmReg >= 0 And rmReg <= 7 Then             'mov reg, [reg] -> deref a parameter pointer
        If regPtr(rmReg) Then regv(regF) = regv(rmReg): regPtr(regF) = False: NativeFLMov = True
        Exit Function
    End If
End Function

Private Function NativeFLArith(ByVal dump As String, ByVal idx As Long, ByRef regv() As String, ByRef regPtr() As Boolean, ByVal opStr As String) As Boolean
    Dim regF As Long, rmIsReg As Boolean, rmReg As Long, rmDisp As Long, rmDeref As Boolean, rv As String
    If NativeFLOperand(dump, idx, regF, rmIsReg, rmReg, rmDisp, rmDeref) = -1 Then Exit Function
    rv = NativeFLRmVal(regv, regPtr, rmIsReg, rmReg, rmDisp, rmDeref)
    If Len(rv) = 0 Or Len(regv(regF)) = 0 Then Exit Function
    regv(regF) = "(" & regv(regF) & opStr & rv & ")": regPtr(regF) = False
    NativeFLArith = True
End Function

Private Function NativeFLImul3(ByVal dump As String, ByVal idx As Long, ByRef regv() As String, ByRef regPtr() As Boolean, ByVal imm8 As Boolean) As Boolean
    Dim regF As Long, rmIsReg As Boolean, rmReg As Long, rmDisp As Long, rmDeref As Boolean, e As Long, rv As String, imm As Long
    e = NativeFLOperand(dump, idx, regF, rmIsReg, rmReg, rmDisp, rmDeref)
    If e = -1 Then Exit Function
    rv = NativeFLRmVal(regv, regPtr, rmIsReg, rmReg, rmDisp, rmDeref)
    If Len(rv) = 0 Then Exit Function
    If imm8 Then imm = NativeDumpInt8(dump, e) Else imm = NativeDumpInt32(dump, e)
    regv(regF) = "(" & rv & " * " & NativeFLNum(imm) & ")": regPtr(regF) = False
    NativeFLImul3 = True
End Function

Private Function NativeFLNum(ByVal v As Long) As String
    NativeFLNum = CStr(v)
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
    'A public class method's true kind (Function / Property Get / Property Let / Sub)
    'comes from its typeinfo FuncDesc and is AUTHORITATIVE - it identifies Get vs Let
    'directly (via the return-slot bit), so it overrides the positional Get/Let adjacency
    'guess in LinkNativeProcNames (which assumes first-occurrence = Get, wrong when VB6
    'lays the Let accessor's name before the Get).  Only fall back to the adjacency kind
    'when the FuncDesc has no entry for this address (e.g. an event handler).
    Dim mkind As String
    If NativeTryMethodKind(addr, mkind) Then
        kindStr = mkind
        If InStr(kindStr, "Property") > 0 Then NVProcEndWord = "Property" Else NVProcEndWord = kindStr
    End If
    'An unclassified module Function recovered from its accumulator-return epilogue
    '(NativeDetectAccumReturn): promote Sub -> Function and remember the return type.
    If kindStr = "Sub" And NVAccumRet Then kindStr = "Function": NVProcEndWord = "Function"
    'A module Function returning a Variant/String/UDT via a hidden retbuf (first param).
    'Promote Sub -> Function; the param list drops the retbuf slot (NativeProcParams).
    If kindStr = "Sub" And NVRetbuf Then kindStr = "Function": NVProcEndWord = "Function"
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
    Else
        'The name already carries a typed signature (an event handler, from
        'getEventComplete - e.g. Winsock1_Error (ByVal Number As Integer, ...)).
        'Map those declared parameter names into the body too, so it reads `Number`
        'instead of arg_C.
        Dim ep As Long, eq As Long, eParams As String
        ep = InStr(nm, "(")
        eq = InStrRev(nm, ")")
        If eq > ep + 1 Then
            eParams = Trim$(Mid$(nm, ep + 1, eq - ep - 1))
            If Len(eParams) > 0 Then NativeBuildArgNameMap addr, eParams
        End If
    End If
    'Append the recovered return type for an accumulator-return module Function.
    Dim asType As String
    If NVAccumRet And Len(NVAccumRetType) > 0 And InStr(kindStr, "Function") > 0 Then asType = " As " & NVAccumRetType
    'A retbuf return is a Variant/String/UDT - indistinguishable from the binary, so As Variant.
    If NVRetbuf And InStr(kindStr, "Function") > 0 Then asType = " As Variant"
    NativeProcHeader = vis & " " & kindStr & " " & nm & asType & "   '" & Hex$(addr)
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
    'A module Function returning via a hidden retbuf has that retbuf as its FIRST param
    '(ebp+8), so user params start at ebp+0xC - exactly like a method's Me slot.  Treat it
    'that way and skip the class-style (last-slot) retbuf drop below.
    Dim retbufFirst As Boolean
    retbufFirst = (NVRetbuf And Not hasMe)
    If hasMe Or retbufFirst Then nParams = slots - 1 Else nParams = slots   'drop the hidden Me/this or retbuf slot
    'A CLASS Function returns via a hidden [out,retval] retbuf param - drop it.  An
    'accumulator-return module Function (NVAccumRet) returns in ax/eax with NO stack
    'retslot, so keep every parameter (else a real argument is wrongly dropped).
    If InStr(kindStr, "Function") > 0 And Not NVAccumRet And Not retbufFirst Then nParams = nParams - 1
    If nParams < 0 Then nParams = 0
    base = IIf(hasMe Or retbufFirst, &HC, &H8)   'first user parameter's ebp offset
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

Private Function NativeOwnerIsClass(ByVal owner As String) As Boolean
    'True when the owning object is a CLASS module (ObjectType bit 0x100000), as
    'opposed to a form, UserControl or standard module.  Used to gate private-field
    'type harvesting to classes (where a Get/Let/Set property's backing variable is a
    'private instance field we surface as `Private field_<off> As <type>`).
    Dim i As Long
    If Len(owner) = 0 Then Exit Function
    On Error Resume Next
    For i = 0 To UBound(gObjectNameArray)
        If gObjectNameArray(i) = owner Then
            NativeOwnerIsClass = ((gObject(i).ObjectType And &H100000) <> 0)
            Exit Function
        End If
    Next
End Function

Private Sub NativeRecordFieldType(ByVal off As Long, ByVal typ As String)
    'Remember the inferred VB type of a private instance field of the current class,
    'so a `Private field_<off> As <type>` declaration block can be emitted at the class
    'top.  Only harvested for classes (NVIsClass) and for plausible field offsets.  A
    'specific type (String/Object) wins over a vaguer numeric guess (Long/Integer)
    'recorded by another access, so a later word-store never downgrades a String field.
    If Not NVIsClass Then Exit Sub
    If Len(NVForm) = 0 Or off <= 0 Or off >= &H2000 Then Exit Sub
    Dim key As String, cur As String
    key = NVForm & ":" & off
    If gClassFieldType Is Nothing Then Set gClassFieldType = New Collection
    cur = NativeColGet(gClassFieldType, key)
    If Len(cur) > 0 Then
        'Keep the stronger evidence: a reference type (String/Object/<class>) is a
        'definite signal; a numeric width is a weak default that must not overwrite it.
        If cur <> "Long" And cur <> "Integer" And cur <> "Byte" Then Exit Sub
        If typ = "Long" Or typ = "Integer" Or typ = "Byte" Then Exit Sub  'don't churn numeric guesses
    End If
    On Error Resume Next
    gClassFieldType.Remove key
    On Error GoTo 0
    gClassFieldType.Add typ, key
End Sub

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
