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
Private NVLoopHdr As Collection    'addresses that are loop headers (back-edge targets)
Private NVCallHandled As Boolean   'set by NativeRuntimeCall: True when the call was recognised
Private NVErrHandler As Long       'address of this procedure's On Error handler block (0 = none)
Private NVProcEndWord As String    'closing keyword for this proc: "Sub" / "Function" / "Property"
Private NVRetN As Long             'the proc's `ret imm16` operand (callee-popped arg bytes), -1 if none
Private NVApiStubCache As Collection 'declared-DLL stub address -> resolved API name (global, "" = not a stub)
Private NVCmpL As String           'pending condition: left operand (symbolic)
Private NVCmpR As String           'pending condition: right operand (symbolic)
Private NVCmpIsTest As Boolean     'the pending compare came from TEST (zero-compare)
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
    NVProcEndWord = "Sub"
    NVLastControl = "": NVLastGuid = "": NVLastImm = "": NVPendingArg = ""
    NVLastLea = 0: NVLastLeaSet = False: NVLastCmp = ""
    NVCmpSet = False: NVCmpL = "": NVCmpR = "": NVCmpIsTest = False: NVFpuChk = False
    NVPendingCall = "": NVErrObjPending = False
    ReDim NVFpu(31): NVFpuTop = 0
    ReDim NVPushImm(31): ReDim NVPushDisp(31): NVPushTop = 0: NVLastPushDisp = 0
    ReDim NVIfTarget(31): NVIfTop = 0: NVIndent = 0
    Dim r As Long
    For r = 0 To 7: NVReg(r) = "": NVRegIsAddr(r) = False: NVRegAddr(r) = "": NVRegAddrDisp(r) = 0: NVRegIsMe(r) = False: NVRegIsFormVt(r) = False: NVRegObjType(r) = "": NVRegObjVt(r) = "": NVRegObjGuid(r) = "": NVRegObjVtGuid(r) = "": Next
    Set NVLocal = New Collection
    Set NVLocalGuid = New Collection
    Set NVStrLits = New Collection
    Set NVVSlot = New Collection
    Set NVSkipLabels = New Collection
    Set NVLoopHdr = New Collection
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

    output = NativeProcHeader(addr) & vbCrLf

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
        'Open a Do loop when this address is the target of a back-edge
        If NativeHas(NVLoopHdr, inst.va) Then
            output = output & NativeIndentStr() & "Do" & vbCrLf
            NVIndent = NVIndent + 1
        End If
        'A label is only needed when it is a real jump target (not an If close
        'point, a loop header, or a dropped error-check guard target)
        If NativeHas(labels, inst.va) And Not NativeIsIfTarget(inst.va) _
           And Not NativeHas(NVSkipLabels, inst.va) And Not NativeHas(NVLoopHdr, inst.va) Then
            output = output & "loc_" & Right$("00000000" & Hex$(inst.va), 8) & ":" & vbCrLf
        End If
        output = output & NativeProcessInst(inst)
        'A call breaks object-identity tracking: a control/object pointer can be
        'reloaded across a call by a path this lightweight tracker does not model,
        'leaving a STALE identity on a callee-saved register (which then mis-names a
        'later property access).  Drop every register's object/control identity
        'except eax (which carries a getter's result) after each call, so a property
        'resolves only from a tight load->deref->call chain with no intervening call
        '- the reliable case - never from a stale tag.
        If (inst.cmdType And C_TYPEMASK) = C_CAL Then
            Dim cr As Long
            For cr = 1 To 7
                NVRegObjType(cr) = "": NVRegObjVt(cr) = "": NVRegObjGuid(cr) = "": NVRegObjVtGuid(cr) = ""
            Next
        End If
    Next

    'Any call still deferred is the proc's last statement - emit it.
    If Len(NVPendingCall) > 0 Then output = output & NativeIndentStr() & NVPendingCall & vbCrLf: NVPendingCall = ""
    NativeCloseIfs output, &H7FFFFFFF
    output = output & "End " & NVProcEndWord & vbCrLf
    DecompileNativeProcToVB = output
    Exit Function
fail:
    DecompileNativeProcToVB = "' Error decompiling " & Hex$(addr) & ": " & Err.Description & vbCrLf
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
                'A form calling its own method: call [vtable + 0x6F8 + slot*4].
                'Checked first: requiring a real gFormVtable slot is a stronger
                'signal than the NVBase control heuristic (which the same form-method
                'calls can otherwise mis-solve into a bogus control base).
                Dim ftgt As Long
                ftgt = NativeFormVtableTarget(disp)
                If ftgt <> 0 Then
                    'The implicit Me/this (the last push, ebx) is normally
                    'untracked, so it never reaches the argument stack - the
                    'tracked pushes are exactly the explicit arguments.
                    Dim fpname As String, fargs As String
                    fargs = NativeArgList()
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
                        Dim cpname As String, cargs As String, csig As String
                        cargs = NativeArgList()
                        cpname = NativeCallTargetName(ctgt)
                        'Drop the trailing hidden return-value buffer (and any extra)
                        'once the method's real parameter count is known.
                        If NativeTryMethodSig(ctgt, csig) Then cargs = NativeTakeArgs(cargs, NativeArgCount(csig))
                        NVPushTop = 0: NativeResetValue
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
            NativeControlProp = "' " & ctlName & "." & propName & "()"
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
            'Object creation - result is a new object; flows into the Set/store
            NVPushTop = 0: NVReg(0) = "New (object)": NativeRuntimeCall = "": Exit Function
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
            If gtN >= 3 Then NativeRuntimeCall = "Get #" & gtA(2) & ", , " & gtA(1)
            Exit Function
        Case InStr(nm, "__vbaLineInputStr") > 0
            'Line Input #<filenum>, <var>.  Args: [var, filenum].
            Dim liA() As String, liN As Long
            NativeArgsSnapshot liA, liN
            If liN >= 2 Then NativeRuntimeCall = "Line Input #" & liA(1) & ", " & liA(0)
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
            If md = 3 Then
                NVReg(reg) = NVReg(rm)
                NVRegIsAddr(reg) = NVRegIsAddr(rm): NVRegAddr(reg) = NVRegAddr(rm): NVRegAddrDisp(reg) = NVRegAddrDisp(rm)   'address propagates on reg->reg
                NVRegIsMe(reg) = NVRegIsMe(rm): NVRegIsFormVt(reg) = NVRegIsFormVt(rm)
                NVRegObjType(reg) = NVRegObjType(rm): NVRegObjVt(reg) = NVRegObjVt(rm)
                NVRegObjGuid(reg) = NVRegObjGuid(rm): NVRegObjVtGuid(reg) = NVRegObjVtGuid(rm)
            ElseIf NativeDecodeDisp(dump, disp, isAbs) Then
                Dim bse As Long, baseObj As Boolean
                bse = NativeMemBase(dump)
                If bse >= 0 And bse <= 7 Then baseObj = NVRegIsMe(bse)
                If Not isAbs And disp < 0 Then
                    NVReg(reg) = NativeGetLocalExpr(disp)
                ElseIf isAbs And NativeIsGlobalAddr(disp) Then
                    NVReg(reg) = NativeGlobalName(disp)      'load of a module-level global
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
                NVRegObjType(reg) = "": NVRegObjVt(reg) = "": NVRegObjGuid(reg) = "": NVRegObjVtGuid(reg) = ""
                If Not isAbs And disp < 0 Then
                    If NativeIsIntrinsicObj(NVReg(reg)) Then NVRegObjType(reg) = NVReg(reg)
                    'Loading a local that holds a control object tags this register
                    'with the control's identity + GUID (for a later property call).
                    Dim lguid As String
                    lguid = NativeGetLocalGuid(disp)
                    If Len(lguid) > 0 Then NVRegObjGuid(reg) = lguid: NVRegObjType(reg) = NVReg(reg)
                ElseIf Not isAbs And disp = 0 And bse >= 0 And bse <= 7 Then
                    NVRegObjVt(reg) = NVRegObjType(bse)
                    NVRegObjVtGuid(reg) = NVRegObjGuid(bse)   'deref of a control pointer -> its vtable carries the GUID
                End If
            End If
        Case &H8D                       'lea r32, [mem]  (address-of)
            modrm = NativeDumpByte(dump, i + 1)
            reg = (modrm \ 8) And 7
            If NativeDecodeDisp(dump, disp, isAbs) Then
                'LEA takes the ADDRESS of a local.  Keep its value in NVReg (a
                'read of the register wants the value), but ALSO remember the
                'local's name so that PUSHing the register - passing the local by
                'reference - shows the local, not its (often 0/stale) value.
                If Not isAbs And disp < 0 Then
                    NVReg(reg) = NativeGetLocalExpr(disp)
                    NVRegIsAddr(reg) = True: NVRegAddr(reg) = "var_" & Hex$(Abs(disp)): NVRegAddrDisp(reg) = disp
                ElseIf isAbs And NativeIsGlobalAddr(disp) Then
                    'address-of a module-level global (e.g. an array passed by ref)
                    NVReg(reg) = NativeGlobalName(disp): NVRegIsAddr(reg) = False
                Else
                    NVReg(reg) = "": NVRegIsAddr(reg) = False
                End If
                NVRegIsMe(reg) = False: NVRegIsFormVt(reg) = False
                NVRegObjType(reg) = "": NVRegObjVt(reg) = "": NVRegObjGuid(reg) = "": NVRegObjVtGuid(reg) = ""
            End If
        Case &H89                       'mov r/m32, r32 (store)
            modrm = NativeDumpByte(dump, i + 1)
            md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
            If md = 3 Then
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
                    'later uses reference the variable rather than re-expanding.
                    If InStr(NVReg(reg), "(") > 0 And Left$(NVReg(reg), 1) <> Chr$(34) Then
                        lname = "var_" & Hex$(Abs(disp))
                        NativeTrackReg = lname & " = " & NVReg(reg)
                        NativeSetLocalExpr disp, lname
                    Else
                        NativeSetLocalExpr disp, NVReg(reg)
                    End If
                    'Storing a tracked control object to a local (plain mov, not
                    '__vbaObjSet) - remember its GUID so a later property access
                    'through that local resolves (e.g. the LET target temp).
                    If Len(NVRegObjGuid(reg)) > 0 Then NativeSetLocalGuid disp, NVRegObjGuid(reg)
                ElseIf isAbs And disp >= OptHeader.ImageBase Then
                    'Store to a module-level global: mov [abs], reg.  Surface a
                    'call / concat / string value as `global_X = ...` (without
                    'this a deferred call folded into the store would be lost).
                    If NativeIsExprValue(NVReg(reg)) Then
                        NativeTrackReg = NativeGlobalName(disp) & " = " & NVReg(reg)
                    End If
                End If
            End If
        Case &H33                       'xor r32, r/m32 (xor reg,reg -> 0)
            modrm = NativeDumpByte(dump, i + 1)
            md = (modrm \ &H40) And 3: reg = (modrm \ 8) And 7: rm = modrm And 7
            If md = 3 And reg = rm Then NVReg(reg) = "0": NVRegIsAddr(reg) = False: NVRegIsMe(reg) = False: NVRegIsFormVt(reg) = False: NVRegObjType(reg) = "": NVRegObjVt(reg) = "": NVRegObjGuid(reg) = "": NVRegObjVtGuid(reg) = ""
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
    Dim s As Single
    On Error Resume Next
    CopyMemory s, bits, 4
    If Abs(s) >= 0.0001 And Abs(s) < 1E+18 And s = Int(s) Then
        NativeNumFromBits = CStr(CLng(s))
    ElseIf Abs(s) >= 0.0001 And Abs(s) < 1E+18 Then
        NativeNumFromBits = Format$(s)
    Else
        NativeNumFromBits = CStr(bits)
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
    If NVCmpSet Then
        op = NativeJccOp(jmpMnem)
        If blockGuard Then op = NativeNegOp(op)
        L = NVCmpL
        If NVCmpIsTest Then R = "0" Else R = NVCmpR
        NVCmpSet = False: NVLastCmp = ""
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

Private Function NativeGlobalObjByOffset(ByVal disp As Long) As String
    'VB6 form-interface vtable slots for the intrinsic global objects.
    Select Case disp
        Case &H14: NativeGlobalObjByOffset = "App"
        Case &H18: NativeGlobalObjByOffset = "Screen"
        Case &H1C: NativeGlobalObjByOffset = "Clipboard"
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
    'Symbolic value of a ModR/M r/m operand: a register's tracked value, or a
    'local stack slot.  "" when it is a memory operand we do not model.
    Dim disp As Long, isAbs As Boolean
    If md = 3 Then
        NativeRmVal = NativeRegVal(rm)
    ElseIf NativeDecodeDisp(dump, disp, isAbs) Then
        If Not isAbs And disp < 0 Then
            NativeRmVal = NativeGetLocalExpr(disp)
        ElseIf isAbs And NativeIsGlobalAddr(disp) Then
            NativeRmVal = NativeGlobalName(disp)
        End If
    End If
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
End Function

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

Private Function NativeStringAt(ByVal va As Long) As String
    'Read a Unicode (BSTR) string constant from the image; returns it quoted,
    'or "" if the address does not hold clean printable text.
    Dim fp As Integer, ch As Integer, s As String, pos As Long, cnt As Long
    On Error GoTo done
    If va < OptHeader.ImageBase Then Exit Function
    fp = FreeFile
    Open SFilePath For Binary Access Read As #fp
    pos = va + 1 - OptHeader.ImageBase
    If pos < 1 Or pos > LOF(fp) Then Close #fp: Exit Function
    Do
        Get #fp, pos, ch
        If ch = 0 Then Exit Do
        If ch < 32 Or ch > 126 Then Exit Do
        s = s & Chr$(ch)
        pos = pos + 2
        cnt = cnt + 1
        If cnt > 256 Then Exit Do
    Loop
    Close #fp
done:
    If Len(s) >= 1 Then NativeStringAt = Chr$(34) & s & Chr$(34)
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
    'Add the parameter list when the name carries no signature yet (event handlers
    'already get a typed one from getEventComplete).  Parameters are named generically
    'arg_<ebp offset> with the count from the proc's `ret N`.
    If InStr(nm, "(") = 0 Then
        Dim psig As String
        If NativeTryMethodSig(addr, psig) Then
            'Real parameter names from the class typeinfo (FuncDesc table).
            nm = nm & "(" & psig & ")"
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
