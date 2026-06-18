Attribute VB_Name = "modNative"
'*********************************************
'modNative
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************
'Module for Processing Native Code
Option Explicit
Private Type API_VBDEF
    rva As Long
    Ordinal As Long
    uName As String
    uDescr As String
End Type
Public exeVB6_APIDEF() As API_VBDEF

Private Type NativeProcType
    offset As Long
    sName As String
End Type
Global gNativeProcArray() As NativeProcType
Global NativeShowOffsets As Boolean
Global NativeShowHexInformation As Boolean

'Event-handler name resolution (see LinkNativeEventNames): each event-link slot's
'raw link value + resolved proc address, and each control event's name + the
'tEventPointer struct address.  An event maps to the slot whose link value equals
'the event's tEventPointer + 8 (the E9 thunk sits at +8 inside the struct).
Private Type tEvSlot
    obj As String
    lnRaw As Long      'raw event-link array value (points to the slot's E9 thunk)
    addr As Long       'resolved procedure VA
End Type
Private Type tEvName
    obj As String
    nm As String
    taPtr As Long      'tEventPointer struct address (taPtr + 8 = the slot's lnRaw)
End Type
Private gEvSlot() As tEvSlot
Private gEvSlotN As Long
Private gEvName() As tEvName
Private gEvNameN As Long

'Decompiled code, grouped by owning object (form/module/class), built once at
'load so the main tree can show it inline.  Keyed "OBJ_<UPPER object name>".
Public gNativeCodeCache As Collection
'Raw native disassembly per object, keyed "OBJ_<UPPER name>", built in the same
'pass as gNativeCodeCache so the Dism tab never disassembles a second time.
Public gNativeDismCache As Collection

'*****************************
'ScanNativeProcsByPrologue
'Find procedures that have no event-link table entry - most importantly the
'Sub/Function procedures in .bas modules, but also any private class procedure.
'
'VB6 lays out every object's native code contiguously, in object-table order,
'and every procedure starts with the standard prologue:
'        55          push ebp
'        8B EC       mov  ebp, esp
'So we read the whole native code range (gProjectInfo.aStartOfCode ..
'aEndOfCode) and add any prologue address we did not already discover through
'the event tables.
'
'Ownership: the form/class objects give exact code-start anchors (their lowest
'event-proc address).  Because code is in object-table order, the modules that
'follow an anchored object - up to the next anchored object - own the non-event
'procedures found in that address range.  Within such a run of modules the procs
'are split in proportion to each module's declared ProcCount (native code keeps
'no per-module proc table, so this is the best available boundary).  See
'AssignRegionProcs.
'*****************************
Public Sub ScanNativeProcsByPrologue(ByVal F As Integer)
    On Error GoTo done

    Dim startVA As Long, endVA As Long, codeLen As Long
    startVA = gProjectInfo.aStartOfCode
    endVA = gProjectInfo.aEndOfCode
    If startVA = 0 Or endVA <= startVA Then Exit Sub
    codeLen = endVA - startVA
    If codeLen < 3 Or codeLen > 16000000 Then Exit Sub

    Dim nObj As Long
    nObj = UBound(gObjectNameArray)        'objects 0..nObj (table order = code order)

    'Per object: anchor (lowest known event-proc addr - 0 if none, i.e. a module),
    'whether it is a module, and its declared procedure count.
    Dim objAnchor() As Long, objIsMod() As Boolean, objPC() As Long
    ReDim objAnchor(nObj): ReDim objIsMod(nObj): ReDim objPC(nObj)
    Dim oi As Long
    For oi = 0 To nObj
        objIsMod(oi) = ((gObject(oi).ObjectType And &H2) = 0)
        objPC(oi) = gObject(oi).ProcCount
    Next oi

    Dim p As Long, dotPos As Long, objName As String
    For p = 0 To UBound(gNativeProcArray) - 1
        If gNativeProcArray(p).offset <> 0 Then
            dotPos = InStr(gNativeProcArray(p).sName, ".")
            If dotPos > 0 Then
                objName = Left$(gNativeProcArray(p).sName, dotPos - 1)
                For oi = 0 To nObj
                    If gObjectNameArray(oi) = objName Then
                        If objAnchor(oi) = 0 Or gNativeProcArray(p).offset < objAnchor(oi) Then _
                            objAnchor(oi) = gNativeProcArray(p).offset
                        Exit For
                    End If
                Next oi
            End If
        End If
    Next p

    'Addresses we already know (event procs + SubMain) so we never duplicate.
    Dim seen As Collection
    Set seen = New Collection
    On Error Resume Next
    For p = 0 To UBound(gNativeProcArray) - 1
        If gNativeProcArray(p).offset <> 0 Then seen.Add 1, "k" & gNativeProcArray(p).offset
    Next p
    If gVBHeader.aSubMain <> 0 Then seen.Add 1, "k" & gVBHeader.aSubMain
    On Error GoTo done

    'Read the whole native code blob once and collect every new prologue address
    '(55 8B EC = push ebp / mov ebp,esp) in ascending order.
    Dim b() As Byte
    ReDim b(codeLen - 1)
    Seek F, startVA + 1 - OptHeader.ImageBase
    Get F, , b

    Dim newAddr() As Long, nNew As Long
    ReDim newAddr(1024): nNew = 0
    Dim j As Long, va As Long
    For j = 0 To codeLen - 3
        If b(j) = &H55 And b(j + 1) = &H8B And b(j + 2) = &HEC Then
            va = startVA + j
            If Not KeyExists(seen, "k" & va) Then
                seen.Add 1, "k" & va
                If nNew > UBound(newAddr) Then ReDim Preserve newAddr(nNew + 1024)
                newAddr(nNew) = va: nNew = nNew + 1
            End If
        End If
    Next j

    'FRAMELESS procedures (e.g. `ReturnNumber = a*b` -> mov eax,[esp+4]; ...; ret N)
    'have no 55 8B EC prologue, so they are invisible to the scan above and a whole
    'module function goes missing.  Recover them from their CALL sites: every E8
    'rel32 whose target is in the code, is preceded by a function boundary (a ret /
    'NOP / int3 - so it is a genuine entry, not a jump into the middle of a proc),
    'and begins with a plausible prologue opcode, is a new procedure.
    For j = 0 To codeLen - 5
        If b(j) = &HE8 Then
            Dim relC As Currency
            relC = CCur(b(j + 1)) + CCur(b(j + 2)) * 256@ + CCur(b(j + 3)) * 65536@ + CCur(b(j + 4)) * 16777216@
            If relC > 2147483647@ Then relC = relC - 4294967296@
            Dim tgtC As Currency
            tgtC = CCur(startVA) + j + 5 + relC
            If tgtC > CCur(startVA) And tgtC < CCur(endVA) Then
                va = CLng(tgtC)
                Dim ti As Long
                ti = va - startVA
                If ti >= 4 And Not KeyExists(seen, "k" & va) Then
                    Dim pb As Byte, sb As Byte, boundary As Boolean
                    pb = b(ti - 1)
                    boundary = (pb = &H90) Or (pb = &HC3) Or (pb = &HCC) Or (b(ti - 3) = &HC2)
                    sb = b(ti)
                    If boundary And NativeIsProcStartByte(sb, b(ti + 1)) Then
                        seen.Add 1, "k" & va
                        If nNew > UBound(newAddr) Then ReDim Preserve newAddr(nNew + 1024)
                        newAddr(nNew) = va: nNew = nNew + 1
                    End If
                End If
            End If
        End If
    Next j

    If nNew = 0 Then GoTo done

    'Keep all discovered addresses in ascending (code) order - the prologue pass is
    'already sorted but the call-site pass appended frameless targets out of order,
    'and the region/proportional assignment below relies on address order.
    Dim sa As Long, sc As Long, st As Long
    For sa = 0 To nNew - 2
        For sc = sa + 1 To nNew - 1
            If newAddr(sc) < newAddr(sa) Then st = newAddr(sa): newAddr(sa) = newAddr(sc): newAddr(sc) = st
        Next sc
    Next sa

    'Build segment boundaries from the anchored (form/class) objects.  Code is
    'laid out in object-table order, so each anchored object starts a region and
    'the module objects that follow it (before the next anchored object) own the
    'non-event procedures found in that region.  Segment 0 covers any leading
    'modules before the first anchored object.
    Dim segStart() As Long, segLead() As Long, nSeg As Long
    ReDim segStart(nObj + 2): ReDim segLead(nObj + 2)
    segStart(0) = startVA: segLead(0) = -1: nSeg = 1
    For oi = 0 To nObj
        If objAnchor(oi) <> 0 Then
            segStart(nSeg) = objAnchor(oi): segLead(nSeg) = oi: nSeg = nSeg + 1
        End If
    Next oi

    'Sort segments by start address (object-table order should already be
    'ascending, but guard against the odd out-of-order layout).
    Dim a As Long, c As Long, ts As Long, tl As Long
    For a = 0 To nSeg - 2
        For c = a + 1 To nSeg - 1
            If segStart(c) < segStart(a) Then
                ts = segStart(a): segStart(a) = segStart(c): segStart(c) = ts
                tl = segLead(a): segLead(a) = segLead(c): segLead(c) = tl
            End If
        Next c
    Next a

    Dim si As Long, regEnd As Long, nextLead As Long
    For si = 0 To nSeg - 1
        If si < nSeg - 1 Then regEnd = segStart(si + 1) Else regEnd = endVA
        If si < nSeg - 1 Then nextLead = segLead(si + 1) Else nextLead = nObj + 1
        Call AssignRegionProcs(newAddr(), nNew, segStart(si), regEnd, segLead(si), nextLead, objIsMod(), objPC())
    Next si

done:
End Sub

'Assign the new (non-event) procedures found in [regStart, regEnd) to the module
'objects that fall between the leading anchored object and the next one.
'Procedures are kept in address (code) order and split between the modules in
'proportion to their declared ProcCount.  If the region has no modules the procs
'are the leading form/class's own private procedures.
Private Sub AssignRegionProcs(ByRef newAddr() As Long, ByVal nNew As Long, _
        ByVal regStart As Long, ByVal regEnd As Long, ByVal leadIdx As Long, _
        ByVal nextLead As Long, ByRef objIsMod() As Boolean, ByRef objPC() As Long)

    Dim regProc() As Long, nReg As Long, i As Long
    ReDim regProc(nNew): nReg = 0
    For i = 0 To nNew - 1
        If newAddr(i) >= regStart And newAddr(i) < regEnd Then
            regProc(nReg) = newAddr(i): nReg = nReg + 1
        End If
    Next i
    If nReg = 0 Then Exit Sub

    'Module objects in (leadIdx, nextLead)
    Dim mods() As Long, nMods As Long, oi As Long
    ReDim mods(UBound(objIsMod) + 1)
    nMods = 0
    For oi = leadIdx + 1 To nextLead - 1
        If oi >= 0 And oi <= UBound(objIsMod) Then
            If objIsMod(oi) Then mods(nMods) = oi: nMods = nMods + 1
        End If
    Next oi

    'No modules -> the leading form/class owns these private procedures.
    If nMods = 0 Then
        If leadIdx >= 0 Then
            For i = 0 To nReg - 1
                AddNativeProc gObjectNameArray(leadIdx), regProc(i)
            Next i
        End If
        Exit Sub
    End If

    'Shares proportional to ProcCount (even split if all counts are zero).
    Dim share() As Long, totPC As Long, m As Long, used As Long
    ReDim share(nMods - 1)
    totPC = 0
    For m = 0 To nMods - 1: totPC = totPC + objPC(mods(m)): Next m
    used = 0
    For m = 0 To nMods - 1
        If totPC > 0 Then
            share(m) = (CLng(nReg) * objPC(mods(m))) \ totPC
        Else
            share(m) = nReg \ nMods
        End If
        used = used + share(m)
    Next m
    'Hand out any rounding leftover round-robin so nothing is dropped.
    Dim rr As Long, leftover As Long
    leftover = nReg - used: rr = 0
    Do While leftover > 0
        share(rr Mod nMods) = share(rr Mod nMods) + 1
        rr = rr + 1: leftover = leftover - 1
    Loop

    'Assign in address order, module by module.
    Dim pIdx As Long, cnt As Long
    pIdx = 0
    For m = 0 To nMods - 1
        For cnt = 1 To share(m)
            If pIdx >= nReg Then Exit For
            AddNativeProc gObjectNameArray(mods(m)), regProc(pIdx)
            pIdx = pIdx + 1
        Next cnt
    Next m
    Do While pIdx < nReg
        AddNativeProc gObjectNameArray(mods(nMods - 1)), regProc(pIdx)
        pIdx = pIdx + 1
    Loop
End Sub

Private Function NativeIsProcStartByte(ByVal b0 As Byte, ByVal b1 As Byte) As Boolean
    'A plausible first opcode of a (frameless) procedure entry.  Used with the
    'preceding-boundary check to validate a call target as a real function start.
    Select Case b0
        Case &H50 To &H57            'push eax..edi
            NativeIsProcStartByte = True
        Case &H6A, &H68              'push imm8 / imm32
            NativeIsProcStartByte = True
        Case &H8B, &H89, &H8D, &H33, &HA1, &H66, &H83, &HFF
            NativeIsProcStartByte = True       'mov/lea/xor/mov-eax-abs/word-prefix/sub-esp/grp5
        Case &HB8 To &HBF            'mov reg, imm32
            NativeIsProcStartByte = True
        Case Else
            NativeIsProcStartByte = False
    End Select
End Function

Private Sub AddNativeProc(ByVal owner As String, ByVal va As Long)
    gNativeProcArray(UBound(gNativeProcArray)).sName = owner & ".proc_" & Hex$(va)
    gNativeProcArray(UBound(gNativeProcArray)).offset = va
    ReDim Preserve gNativeProcArray(UBound(gNativeProcArray) + 1)
End Sub

'*****************************
'LinkNativeProcNames
'Attach the real procedure names (Greet, Name, Form_Load, ...) to the native
'procedures we decompile.
'
'VB6 stores, per object, a procedure-names array (tObject.aProcNamesArray): one
'pointer-to-name per procedure, indexed in the object's procedure order, which
'is also the order the procedures are laid out in memory (ascending address).
'Null entries are unnamed/private procedures.  We do NOT use tObject's proc
'address table here - in native EXEs it is a descriptor/thunk structure, not a
'flat code-address array - so instead we take the procedure entry points we
'already discovered for this object (gNativeProcArray, filled by the event-link
'tables and ScanNativeProcsByPrologue), sort them ascending, and pair name index
'i with the i-th address.
'
'Safety: we only map an object when the number of discovered procedures matches
'the object's declared ProcCount exactly.  If they differ, the absolute name
'indices can no longer be trusted to line up with address order, so we leave the
'object's procedures as proc_<addr> (no risk of attaching a wrong name).
'
'Must run AFTER ScanNativeProcsByPrologue (so gNativeProcArray is complete) and
'while the EXE file F is still open.  Adds entries to SubNamelist keyed by the
'exact code address, so NativeProcName / NativeLookupName resolve them directly.
'*****************************
Public Sub LinkNativeProcNames(ByVal F As Integer)
    On Error GoTo done
    Dim oi As Long, nObj As Long
    nObj = UBound(gObjectNameArray)

    Dim p As Long, dotPos As Long, owner As String
    Dim addrs() As Long, nA As Long, a As Long, b As Long, t As Long
    Dim namesVA() As Long, pc As Long, i As Long, nm As String

    For oi = 0 To nObj
        pc = gObject(oi).ProcCount
        If pc <= 0 Then GoTo nextObj
        If gObject(oi).aProcNamesArray = 0 Then GoTo nextObj

        'Collect this object's discovered procedure entry points.
        ReDim addrs(UBound(gNativeProcArray))
        nA = 0
        For p = 0 To UBound(gNativeProcArray) - 1
            If gNativeProcArray(p).offset <> 0 Then
                dotPos = InStr(gNativeProcArray(p).sName, ".")
                If dotPos > 0 Then owner = Left$(gNativeProcArray(p).sName, dotPos - 1) Else owner = ""
                If owner = gObjectNameArray(oi) Then
                    addrs(nA) = gNativeProcArray(p).offset
                    nA = nA + 1
                End If
            End If
        Next p

        'The names array is indexed by the object's FULL procedure table, which can
        'include leading non-code "phantom" slots (interface/property descriptors VB6
        'counts in ProcCount but emits no body for).  The real procedures we
        'discovered occupy the LAST nA slots, so name-array index i maps to discovered
        'address (i - base) with base = pc - nA.  When pc = nA (no phantoms) this is
        'the original 1:1 alignment.  We cannot align when more procs were discovered
        'than the array can hold.
        If nA > pc Then GoTo nextObj
        Dim base As Long
        base = pc - nA

        'Sort the addresses ascending (procedure / name index order).
        For a = 0 To nA - 2
            For b = a + 1 To nA - 1
                If addrs(b) < addrs(a) Then t = addrs(a): addrs(a) = addrs(b): addrs(b) = t
            Next b
        Next a

        'Visibility: named members of a public-interface object (class /
        'usercontrol - ObjectType bit 0x100000) are Public by VB6 default;
        'everything else (forms, modules) defaults to Private.
        Dim vis As String
        If (gObject(oi).ObjectType And &H100000) <> 0 Then vis = "Public" Else vis = "Private"

        'Read the names array (index -> name-string VA) and pair with addresses.
        ReDim namesVA(pc - 1)
        Seek F, gObject(oi).aProcNamesArray + 1 - OptHeader.ImageBase
        Get F, , namesVA
        'Safety: every NAMED slot must fall in the tail region [base, pc) that aligns
        'with the discovered addresses.  A name in the leading phantom region means
        'the index model is untrustworthy here - leave the object unnamed rather than
        'attach a wrong name.
        If base > 0 Then
            For i = 0 To base - 1
                If NativeValidNamePtr(namesVA(i)) Then GoTo nextObj
            Next i
        End If
        'A class/usercontrol exposes its methods through a COM vtable: IUnknown(3) +
        'IDispatch(4) = 7 slots, then the user methods at 0x1C + namedSeq*4 (in name
        'order).  Record that map so a "call [Me_vtable + off]" self-call resolves to
        'the method (parallel to the form 0x6F8 path).  Keyed "Owner:off<offset>".
        Dim isClass As Boolean, namedSeq As Long
        isClass = ((gObject(oi).ObjectType And &H100000) <> 0)
        namedSeq = 0
        Dim prevName As String, prevPos As Long, prevIdx As Long, pos As Long, ai As Long
        prevName = "": prevPos = -1: prevIdx = -2

        'For a FORM, locate its FuncDesc pointer array (parallel to the name array) so
        'each PUBLIC method's REAL vtable offset can be mapped below.  A class's methods
        'sit at the predictable 0x1C+seq*4 layout (handled inline); a form's own public
        'methods sit at high offsets (0x750.. on frmClient) that ONLY the FuncDesc
        'records carry - without this a cross-module `<formInstance>.UnkVCall_<off>`
        'call can't resolve to <form>.<method>.
        Dim frmArr As Long, frmHdr As Long, fj As Long
        frmArr = 0
        If Not isClass And gObject(oi).aObjectInfo <> 0 Then
            frmHdr = NativeFileDword(F, gObject(oi).aObjectInfo + &HC)
            If frmHdr >= OptHeader.ImageBase Then
                If NativeArrayHasFuncDesc(F, NativeFileDword(F, frmHdr + &H18)) Then
                    frmArr = NativeFileDword(F, frmHdr + &H18)
                Else
                    For fj = &H1C To &H140 Step 4
                        If NativeArrayHasFuncDesc(F, NativeFileDword(F, frmHdr + fj)) Then frmArr = NativeFileDword(F, frmHdr + fj): Exit For
                    Next fj
                End If
            End If
        End If
        For i = 0 To pc - 1
            nm = ""
            If NativeValidNamePtr(namesVA(i)) Then
                Seek F, namesVA(i) + 1 - OptHeader.ImageBase
                nm = GetUntilNull(F)
            End If
            ai = i - base                                   'discovered-address index
            If Len(nm) > 0 And ai >= 0 And ai <= nA - 1 Then
                pos = UBound(SubNamelist)
                SubNamelist(pos).strName = gObjectNameArray(oi) & "." & nm
                SubNamelist(pos).offset = addrs(ai)
                SubNamelist(pos).visibility = vis
                SubNamelist(pos).kind = ""                  'default -> Sub
                'A read/write property is stored as the SAME name at two adjacent
                'indices: the first is the Get accessor, the second the Let/Set.
                If i = prevIdx + 1 And nm = prevName And prevPos >= 0 Then
                    SubNamelist(prevPos).kind = "Property Get"
                    SubNamelist(pos).kind = "Property Let"
                End If
                ReDim Preserve SubNamelist(UBound(SubNamelist) + 1)
                If isClass Then
                    On Error Resume Next
                    gFormVtable.Add addrs(ai), gObjectNameArray(oi) & ":off" & (&H1C + namedSeq * 4)
                    On Error GoTo done
                ElseIf frmArr <> 0 And i <= 255 Then
                    'Form public method: take its REAL vtable offset from the parallel
                    'FuncDesc and map it to this method's address.  Only when that slot
                    'genuinely holds a FuncDesc (event handlers / private slots have a
                    'name but no public FuncDesc -> skipped, not mis-mapped).
                    Dim fdp As Long, fvoff As Long
                    fdp = NativeFileDword(F, frmArr + i * 4)
                    If fdp <> 0 And NativeIsFuncDesc(F, fdp) Then
                        fvoff = (NativeFileDword(F, fdp) \ &H10000) And &HFFFC
                        If fvoff >= &H1C Then
                            On Error Resume Next
                            gFormVtable.Add addrs(ai), gObjectNameArray(oi) & ":off" & fvoff
                            On Error GoTo done
                        End If
                    End If
                End If
                namedSeq = namedSeq + 1
                prevName = nm: prevPos = pos: prevIdx = i
            End If
        Next i
nextObj:
    Next oi
done:
End Sub

Private Function NativeValidNamePtr(ByVal va As Long) As Boolean
    'A real procedure-name pointer points inside this image.  Some EXEs fill unused
    'name-array slots with garbage (e.g. 0x02020202, ~33 MB) rather than 0, which a
    'bare ">= ImageBase" test wrongly accepts as a name - bound it to the image.
    NativeValidNamePtr = (va >= OptHeader.ImageBase And va < OptHeader.ImageBase + &H1000000)
End Function

Private Function KeyExists(ByRef col As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = col(key)
    KeyExists = (Err.Number = 0)
    Err.Clear
End Function

'---------------------------------------------------------------------------
' Native event-handler name resolution (Form_Load, Timer1_Timer, ...)
'
' Form/control event handlers are NOT in aProcNamesArray, so LinkNativeProcNames
' cannot name them.  Their real ADDRESSES are discovered by the native event-link
' block (E9 thunks -> proc VA, in vtable SLOT order); their NAMES come from the
' event tables (getEventComplete -> "Timer1_Timer"), keyed by an aEvent ordering
' value.  Sorting events by aEvent reproduces the slot order, so: among the
' event-link slots NOT already named by aProcNamesArray (the public methods), the
' first E (in slot order) are the events - assign the aEvent-sorted names to them.
'(tEvSlot/tEvName types and the gEv* arrays are declared at the top of the module.)
'---------------------------------------------------------------------------
Public Sub ResetEventLists()
    ReDim gEvSlot(63): gEvSlotN = 0
    ReDim gEvName(63): gEvNameN = 0
End Sub

Public Sub AddEventSlot(ByVal obj As String, ByVal lnRaw As Long, ByVal addr As Long)
    'Called from the native event-link block for each slot (lnRaw = the raw link
    'array value, addr = the resolved procedure VA the thunk jumps to).
    On Error Resume Next
    If gEvSlotN > UBound(gEvSlot) Then ReDim Preserve gEvSlot(gEvSlotN + 64)
    gEvSlot(gEvSlotN).obj = obj
    gEvSlot(gEvSlotN).lnRaw = lnRaw
    gEvSlot(gEvSlotN).addr = addr
    gEvSlotN = gEvSlotN + 1
End Sub

Public Sub AddEventName(ByVal obj As String, ByVal nm As String, ByVal taPtr As Long)
    'Called from the per-control event block (taPtr = the tEventPointer address).
    On Error Resume Next
    If gEvNameN > UBound(gEvName) Then ReDim Preserve gEvName(gEvNameN + 64)
    gEvName(gEvNameN).obj = obj
    gEvName(gEvNameN).nm = nm
    gEvName(gEvNameN).taPtr = taPtr
    gEvNameN = gEvNameN + 1
End Sub

Public Sub LinkNativeEventNames()
    'Run AFTER LinkNativeProcNames.  Each control event's tEventPointer struct holds
    'its own E9 thunk at offset +8, and the event-link array entry for that slot is
    'exactly that thunk address (lnRaw).  So an event handler maps to the native
    'slot whose link value equals the event's tEventPointer + 8 - a direct, exact
    'correlation (no heuristics).  Helper/private methods occupy the slots that no
    'event points at, and stay unnamed.
    On Error GoTo done
    Dim e As Long, s As Long, pos As Long
    For e = 0 To gEvNameN - 1
        For s = 0 To gEvSlotN - 1
            If gEvSlot(s).obj = gEvName(e).obj And gEvSlot(s).lnRaw = gEvName(e).taPtr + 8 Then
                pos = UBound(SubNamelist)
                SubNamelist(pos).strName = gEvName(e).obj & "." & gEvName(e).nm
                SubNamelist(pos).offset = gEvSlot(s).addr
                SubNamelist(pos).visibility = "Private"
                SubNamelist(pos).kind = ""
                ReDim Preserve SubNamelist(UBound(SubNamelist) + 1)
                Exit For
            End If
        Next s
    Next e
done:
End Sub

'---------------------------------------------------------------------------
' Public-method parameter NAMES (and Function/Sub-ness) for classes / usercontrols
'
' VB6 compiles a class's public interface to a typeinfo block whose FuncDesc
' records carry each method's parameter names.  Chain (verified on Dungeon,
' Strategy2 and the decompiler's own classes):
'   tObject.aObjectInfo -> tObjectInfo +0x0C (aSmallRecord) -> TYPEINFO_HDR
'   FuncDesc array: if dw(hdr+0x18) points to a FuncDesc use it, else the array
'     is inline in the header - scan hdr+0x18..+0x140 for the first FuncDesc ptr.
'   FuncDesc (0x24B): +0x04 = 0x0000FFFF sig ; +0x10 -> param-name char*[] ;
'     +0x00: (byte2)&0xFC = the method's VTABLE OFFSET (maps to the method via the
'     gFormVtable "Owner:off..." class map); param count = (byte0\4) - (byte1 AND 1)
'     - the byte1 bit is the hidden return slot (Function/Property), so a Sub's
'     count is byte0\4 and a Function's is byte0\4 - 1.
' Param names are read straight from the +0x10 list (exactly the computed count),
' so the project-wide shared name table needs no boundary guessing.
'---------------------------------------------------------------------------
Private Function NativeFileDword(ByVal F As Integer, ByVal va As Long) As Long
    On Error Resume Next
    Dim v As Long
    If va < OptHeader.ImageBase Then Exit Function
    Seek F, va + 1 - OptHeader.ImageBase
    Get #F, , v
    NativeFileDword = v
End Function

Private Function NativeIsFuncDesc(ByVal F As Integer, ByVal p As Long) As Boolean
    Dim voff As Long
    If p < OptHeader.ImageBase Then Exit Function
    If NativeFileDword(F, p + 4) <> &HFFFF& Then Exit Function   '&HFFFF& = Long 65535 (&HFFFF alone is Integer -1)
    voff = (NativeFileDword(F, p) \ &H10000) And &HFC
    NativeIsFuncDesc = (voff >= &H1C And voff <= &H2000)
End Function

Private Function NativeVarTypeName(ByVal tc As Long, ByVal byteSize As Long) As String
    'Map a typeinfo VarDesc type code (the +0x18 field) to a VB type name.  Verified
    'across Dungeon + RPGWOEdit: 3 = Boolean, 6 = Integer, 0x10 = String.  An unknown
    'code falls back to the declared byte size (the gap to the next field's instance
    'offset): 1 = Byte, 2 = Integer, 8 = Double, else Long - correct for the common
    'cases; extend the Select as more codes are confirmed.
    Select Case tc
        Case 3: NativeVarTypeName = "Boolean"
        Case 6: NativeVarTypeName = "Integer"
        Case &H10: NativeVarTypeName = "String"
        Case Else
            Select Case byteSize
                Case 1: NativeVarTypeName = "Byte"
                Case 2: NativeVarTypeName = "Integer"
                Case 8: NativeVarTypeName = "Double"
                Case Else: NativeVarTypeName = "Long"
            End Select
    End Select
End Function

Private Function NativeArrayHasFuncDesc(ByVal F As Integer, ByVal base As Long) As Boolean
    'True when `base` points to a method-FuncDesc pointer array - i.e. one of its first
    'few slots holds a valid FuncDesc.  Tolerates LEADING NULL slots (the array is
    'parallel to the name array, which begins with the object's private methods that
    'have no public FuncDesc), so element[0] being 0 is normal, not a miss.
    If base < OptHeader.ImageBase Then Exit Function
    Dim k As Long, elem As Long
    For k = 0 To 15
        elem = NativeFileDword(F, base + k * 4)
        If elem <> 0 And elem <> &HFFFFFFFF Then
            If NativeIsFuncDesc(F, elem) Then NativeArrayHasFuncDesc = True: Exit Function
        End If
    Next k
End Function

Public Sub LinkNativePublicParams(ByVal F As Integer)
    On Error GoTo done
    Dim oi As Long, aoi As Long, hdr As Long, owner As String
    Dim pc As Long, i As Long, nMethods As Long, namesVA() As Long
    Dim a18 As Long, arr As Long, j As Long
    Dim p As Long, f00 As Long, voff As Long, b0 As Long, b1 As Long, nP As Long
    Dim pnl As Long, k As Long, nameVA As Long, sig As String, nm As String
    Dim addr As Long, v As Variant, kind As String, fflags As Long
    Dim isClass As Boolean, maxJ As Long

    For oi = 0 To UBound(gObjectNameArray)
        aoi = gObject(oi).aObjectInfo
        If aoi = 0 Then GoTo nextO
        owner = gObjectNameArray(oi)
        hdr = NativeFileDword(F, aoi + &HC)            'tObjectInfo.aSmallRecord
        'aSmallRecord is -1 for modules (no public typeinfo) -> skipped; classes and
        'forms both have a real typeinfo block here.
        If hdr < OptHeader.ImageBase Then GoTo nextO

        'Public instance-VARIABLE names, recovered from the object's typeinfo VarDesc
        'array (verified on RPGWOEdit frmReSize/frmFileSource and Dungeon frmMainMenu).
        'In the typeinfo record (hdr): hdr+0x10 = variable count, hdr+0x20 -> a pointer
        'array of VarDesc records.  Each VarDesc (0x1C bytes): +0x00 = name char* (VA),
        '+0x14 = the instance byte offset (form public vars start at 0x34, in
        'declaration order).  Store name keyed "Owner:<decimal offset>" so a
        '`mov [Me+off], value` store renders the real name instead of field_<off>.
        Dim vCnt As Long, vArr As Long, vi As Long, vd As Long, vNameVA As Long, vInstOff As Long, vNm As String
        Dim vType As Long, vN() As String, vO() As Long, vT() As Long, vK As Long
        vCnt = NativeFileDword(F, hdr + &H10)
        vArr = NativeFileDword(F, hdr + &H20)
        If vCnt > 0 And vCnt <= 1000 And vArr >= OptHeader.ImageBase Then
            ReDim vN(vCnt): ReDim vO(vCnt): ReDim vT(vCnt): vK = 0
            For vi = 0 To vCnt - 1
                vd = NativeFileDword(F, vArr + vi * 4)
                If vd >= OptHeader.ImageBase Then
                    vNameVA = NativeFileDword(F, vd)
                    vInstOff = NativeFileDword(F, vd + &H14)        'instance byte offset
                    vType = NativeFileDword(F, vd + &H18)           'VB type code (3=Bool,6=Int,0x10=String)
                    'Validate: a real per-instance field name pointer and a plausible
                    'instance offset (forms start vars at 0x34; guard the range).
                    If vNameVA >= OptHeader.ImageBase And vInstOff >= &H30 And vInstOff < &H2000 Then
                        Seek F, vNameVA + 1 - OptHeader.ImageBase
                        vNm = GetUntilNull(F)
                        If Len(vNm) > 0 And NativeValidNamePtr(vNameVA) Then
                            On Error Resume Next
                            gFieldName.Remove owner & ":" & vInstOff
                            gFieldName.Add vNm, owner & ":" & vInstOff
                            On Error GoTo done
                            vN(vK) = vNm: vO(vK) = vInstOff: vT(vK) = vType: vK = vK + 1
                        End If
                    End If
                End If
            Next vi
            'Build the Public declaration block (declaration order = offset order),
            'sizing an unknown type code from the gap to the next field's offset.
            If vK > 0 Then
                Dim declB As String, vsz As Long, vq As Long
                declB = ""                              'reset per object (Dim is function-scoped)
                For vq = 0 To vK - 1
                    If vq < vK - 1 Then vsz = vO(vq + 1) - vO(vq) Else vsz = 4
                    declB = declB & "Public " & vN(vq) & " As " & NativeVarTypeName(vT(vq), vsz) & vbCrLf
                Next vq
                On Error Resume Next
                gFieldDecl.Remove owner
                gFieldDecl.Add declB, owner
                On Error GoTo done
            End If
        End If

        isClass = ((gObject(oi).ObjectType And &H100000) <> 0)
        'A class/usercontrol's public-method count (non-null name slots) bounds its
        'CONTIGUOUS FuncDesc array.  A FORM's array is SPARSE (0x0 gaps for the private
        'event-handler slots) and its method names live in the event tables, so it is
        'read with gap-skipping up to its end marker instead (nMethods stays 0).
        nMethods = 0
        If isClass Then
            pc = gObject(oi).ProcCount
            If pc <= 0 Or gObject(oi).aProcNamesArray = 0 Then GoTo nextO
            ReDim namesVA(pc - 1)
            Seek F, gObject(oi).aProcNamesArray + 1 - OptHeader.ImageBase
            Get F, , namesVA
            For i = 0 To pc - 1
                If NativeValidNamePtr(namesVA(i)) Then nMethods = nMethods + 1
            Next i
            If nMethods = 0 Then GoTo nextO
        End If

        'Locate the FuncDesc pointer array.  hdr+0x18 holds the array BASE pointer.
        'The array is parallel to the method-name array, so it can start with LEADING
        'NULL slots for the object's PRIVATE methods (e.g. Class_Initialize/Terminate)
        'that have no public FuncDesc - element[0] is then 0, not a FuncDesc.  Treat
        'a18 as the base when ANY of its first slots holds a real FuncDesc (skipping the
        'leading nulls); only if a18 fails entirely, scan the other header slots for a
        'pointer to such an array (double-deref, matching the a18 test).
        a18 = NativeFileDword(F, hdr + &H18)
        If NativeArrayHasFuncDesc(F, a18) Then
            arr = a18
        Else
            arr = 0
            For j = &H1C To &H140 Step 4
                If NativeArrayHasFuncDesc(F, NativeFileDword(F, hdr + j)) Then arr = NativeFileDword(F, hdr + j): Exit For
            Next j
        End If
        If arr = 0 Then GoTo nextO

        'A class array spans pc slots (parallel to the name array): leading nulls for
        'private methods, then the public-method FuncDescs, then a terminator.  A form
        'array is sparse (gaps for private event-handler slots), read to a fixed bound.
        Dim started As Boolean
        started = False
        If isClass Then maxJ = pc - 1 Else maxJ = 255
        For j = 0 To maxJ
            p = NativeFileDword(F, arr + j * 4)
            If p = 0 Then
                If isClass Then
                    If started Then Exit For Else GoTo contJ   'leading null phantom -> skip; trailing -> stop
                End If
                GoTo contJ                 'form array is sparse - skip the gap
            End If
            If Not NativeIsFuncDesc(F, p) Then
                If isClass Then
                    If started Then Exit For Else GoTo contJ   'leading private (non-FuncDesc) -> skip
                End If
                Exit For                   'non-FuncDesc -> end of (form) array
            End If
            started = True
            f00 = NativeFileDword(F, p)
            voff = (f00 \ &H10000) And &HFFFC             '&HFFFC: form offsets are 0x6F8+
            b0 = f00 And &HFF
            b1 = (f00 \ &H100) And &H1            'hidden return slot bit
            nP = (b0 \ 4) - b1
            If nP < 0 Then nP = 0
            'Method kind: the property bit (flags & 0x800) marks a Property Get;
            'else the hidden-return-slot bit (b1) distinguishes Function from Sub.
            fflags = NativeFileDword(F, p + &HC) \ &H10000
            If (fflags And &H800) <> 0 Then
                kind = "Property Get"
            ElseIf b1 = 1 Then
                kind = "Function"
            Else
                kind = "Sub"
            End If
            pnl = NativeFileDword(F, p + &H10)
            sig = ""
            For k = 0 To nP - 1
                nameVA = NativeFileDword(F, pnl + k * 4)
                nm = ""
                If nameVA >= OptHeader.ImageBase Then
                    Seek F, nameVA + 1 - OptHeader.ImageBase
                    nm = GetUntilNull(F)
                End If
                If Len(nm) = 0 Then nm = "arg" & (k + 1)
                If Len(sig) > 0 Then sig = sig & ", "
                sig = sig & nm
            Next k
            'Map the method's vtable offset to its code address (the class vtable
            'map filled by LinkNativeProcNames), then key the signature by address.
            addr = 0
            On Error Resume Next
            If voff >= &H6F8 Then
                v = gFormVtable(owner & ":" & ((voff - &H6F8) \ 4))   'form: event-link vtable slot
            Else
                v = gFormVtable(owner & ":off" & voff)                'class: 0x1C-based offset
            End If
            If Err.Number = 0 Then addr = CLng(v)
            Err.Clear
            On Error GoTo done
            If addr <> 0 Then
                On Error Resume Next
                gMethodSig.Remove "A" & addr
                gMethodSig.Add sig, "A" & addr
                gMethodKind.Remove "A" & addr
                gMethodKind.Add kind, "A" & addr
                On Error GoTo done
            End If
contJ:
        Next j
nextO:
    Next oi
done:
End Sub

Public Sub Decode(ByVal Filename As String)
'*****************************
'Purpose: To Get the procdures of a Native Exe and produce a report
'*****************************

Dim FileNum As Integer
    Dim F As Long
    F = FreeFile
    Open App.Path & "\dump\" & SFile & "\NativeOut.txt" For Output As #F
        Print #F, "Semi VB Decompiler - VisualBasicZone.com"
        Print #F, "Native Output for : " & Filename
        Print #F, "---------------------------------"
       
        Print #F, "Procedure Offsets:"
        If gProjectInfo.aNativeCode <> 0 Then
            If gVBHeader.aSubMain <> 0 Then
                 Print #F, gVBHeader.aSubMain
            End If
        End If
        Dim i As Integer
        For i = 0 To UBound(gNativeProcArray) - 1
             Print #F, gNativeProcArray(i).offset
        Next i
    Close #F

    'Decompile every procedure to readable VB, grouped by object, for the tree.
    Call BuildNativeCodeCache

End Sub

Public Sub BuildNativeCodeCache()
'*****************************
'Decompile each procedure (grouped by its owning object) and cache the result
'so the main tree can display per-object code without re-running the engine.
'Addresses come from gNativeProcArray - the same authoritative list the Native
'Procedure Decompile window uses - so they are always valid (SubNamelist's
'event-proc addresses are unreliable).  The procedure name in each header is
'resolved internally by the engine from the address.
'*****************************
    On Error Resume Next
    Dim objNames() As String, objCode() As String, objCount As Long
    Dim p As Long, oi As Long, found As Long, addr As Long, body As String, total As Long, done As Long
    Dim sn As String, objName As String, dotPos As Long

    Set gNativeCodeCache = New Collection
    Set gNativeDismCache = New Collection
    Set gUsedWin32Const = New Collection            'reset the recognised-constant set for this run
    Set gUDTDesc = New Collection                   'reset recovered UDT record descriptors (harvested during this decompile pass)
    modNativeToVB.NativeDetectUDTStringFields        'scan __vbaStrFixstr/Lset sites -> fixed-string fields of descriptor-less UDTs (must run before the per-proc __vbaRedim harvest renders the Type)
    Dim objDism() As String
    Dim ub As Long
    ub = -1
    ub = UBound(gNativeProcArray)                      '-1 stays if not dimensioned
    If ub < 0 Then Exit Sub
    total = ub
    ReDim objNames(64): ReDim objCode(64): ReDim objDism(64): objCount = 0

    For p = 0 To ub - 1
        If CancelDecompile = True Then Exit For
        addr = gNativeProcArray(p).offset
        If addr = 0 Then GoTo nextProc

        'Owning object is the prefix of the synthetic name "Object.proc:addr"
        sn = gNativeProcArray(p).sName
        dotPos = InStr(sn, ".")
        If dotPos > 0 Then objName = UCase$(Left$(sn, dotPos - 1)) Else objName = "MODULE1"

        found = -1
        For oi = 0 To objCount - 1
            If objNames(oi) = objName Then found = oi: Exit For
        Next
        If found = -1 Then
            If objCount > UBound(objNames) Then
                ReDim Preserve objNames(objCount + 64): ReDim Preserve objCode(objCount + 64): ReDim Preserve objDism(objCount + 64)
            End If
            objNames(objCount) = objName
            objCode(objCount) = ""
            objDism(objCount) = ""
            found = objCount: objCount = objCount + 1
        End If

        done = done + 1
        frmMain.txtStatus.Text = frmMain.txtStatus.Text & "Decompiling " & sn & " (" & done & "\" & total & ")" & vbCrLf
        frmMain.txtStatus.Refresh
        DoEvents

        body = modNativeToVB.DecompileNativeProcToVB(addr)
        objCode(found) = objCode(found) & body & vbCrLf
        'Raw disassembly captured during the decompile above (no re-disassembly).
        objDism(found) = objDism(found) & _
            "; ---------------------------------------------" & vbCrLf & _
            "; " & sn & "  (" & Hex$(addr) & "h)" & vbCrLf & _
            "; ---------------------------------------------" & vbCrLf & _
            modNativeToVB.NVLastDisasmText & vbCrLf
nextProc:
    Next p

    For oi = 0 To objCount - 1
        gNativeCodeCache.Add objCode(oi), "OBJ_" & objNames(oi)
        gNativeDismCache.Add objDism(oi), "OBJ_" & objNames(oi)
    Next
End Sub

Public Function GetNativeObjectCode(ByVal objName As String) As String
'*****************************
'Return the cached decompiled code for an object, or empty stubs built from the
'procedure list when nothing was cached (e.g. P-Code projects).
'*****************************
    Dim code As String, p As Long
    On Error Resume Next
    If Not gNativeCodeCache Is Nothing Then
        code = gNativeCodeCache("OBJ_" & UCase$(objName))
        If Len(code) > 0 Then GetNativeObjectCode = code: Exit Function
    End If
    'Fallback: signature-only stubs
    For p = 0 To UBound(gProcedureList)
        If UCase$(gProcedureList(p).strParent) = UCase$(objName) And gProcedureList(p).strProcedureName <> "" Then
            If Right$(gProcedureList(p).strProcedureName, 1) = ")" Then
                code = code & "Private Sub " & gProcedureList(p).strProcedureName & vbCrLf
            Else
                code = code & "Private Sub " & gProcedureList(p).strProcedureName & "()" & vbCrLf
            End If
            code = code & "End Sub" & vbCrLf
        End If
    Next
    GetNativeObjectCode = code
End Function

'*****************************
'GetNativeObjectDisassembly
'Return the raw native disassembly of every procedure that belongs to objName
'(form / module / class).  Served from gNativeDismCache, which is built in the
'same pass as the decompiled code (BuildNativeCodeCache) so the procedure is
'never disassembled a second time.  Falls back to a live disassembly only if
'the cache was never built.
'*****************************
Public Function GetNativeObjectDisassembly(ByVal objName As String) As String
    On Error GoTo done
    If gProjectInfo.aNativeCode = 0 Then
        GetNativeObjectDisassembly = "; This project is compiled to P-Code - no native assembly available."
        Exit Function
    End If

    'Fast path: cached during load.
    If Not gNativeDismCache Is Nothing Then
        Dim cached As String
        cached = gNativeDismCache("OBJ_" & UCase$(objName))
        If Len(cached) > 0 Then GetNativeObjectDisassembly = cached: Exit Function
    End If

    'Fallback: cache not built yet - disassemble live this once.
    Dim ub As Long
    ub = -1
    ub = UBound(gNativeProcArray)
    If ub < 1 Then Exit Function

    Dim p As Long, dotPos As Long, owner As String, va As Long, cnt As Long
    Dim sb As String, target As String, ignoreVB As String
    target = UCase$(objName)
    For p = 0 To ub - 1
        If gNativeProcArray(p).offset <> 0 Then
            dotPos = InStr(gNativeProcArray(p).sName, ".")
            If dotPos > 0 Then owner = UCase$(Left$(gNativeProcArray(p).sName, dotPos - 1)) Else owner = ""
            If owner = target Then
                va = gNativeProcArray(p).offset
                'Fills NVLastDisasmText as a side effect; the VB return is ignored.
                ignoreVB = modNativeToVB.DecompileNativeProcToVB(va)
                sb = sb & "; ---------------------------------------------" & vbCrLf
                sb = sb & "; " & gNativeProcArray(p).sName & "  (" & Hex$(va) & "h)" & vbCrLf
                sb = sb & "; ---------------------------------------------" & vbCrLf
                sb = sb & modNativeToVB.NVLastDisasmText & vbCrLf
                cnt = cnt + 1
            End If
        End If
    Next p

    If cnt = 0 Then sb = "; No native procedures found for " & objName
    GetNativeObjectDisassembly = sb
    Exit Function
done:
    GetNativeObjectDisassembly = sb & vbCrLf & "; (disassembly stopped: " & err.Description & ")"
End Function

Sub VBFunction_Description_Init(ByVal fRes As String)
'*****************************
'Purpose: To load the Msvbvm60.dll api list from a file
'*****************************
Dim lfp As Integer, i As Long
Dim sAdr As String, sOrd As String, sName As String, sDef As String
lfp = FreeFile
Erase exeVB6_APIDEF()

    Open fRes For Input Access Read As #lfp
        i = 0
        Do
        i = i + 1
            Input #lfp, sAdr, sOrd, sName, sDef
            If LCase$(sAdr) <> "eof" Then
                ReDim Preserve exeVB6_APIDEF(1 To i)
                exeVB6_APIDEF(i).rva = Val("&H" & sAdr)

                exeVB6_APIDEF(i).Ordinal = CLng(sOrd)
                exeVB6_APIDEF(i).uName = sName
                exeVB6_APIDEF(i).uDescr = sDef
            Else
                Exit Do
            End If
        Loop Until EOF(1)
    
    Close #lfp

End Sub
Public Function VBFunction_Description(ByVal inOrdinal As Long, ByVal inAPIname As String, ByRef outRName As String) As String
'*****************************
'Purpose: To return the description of a function
'*****************************
Dim i As Long


If inOrdinal > 0 And inAPIname = vbNullString Then
    'by ordinal :
    For i = 1 To UBound(exeVB6_APIDEF)
        If exeVB6_APIDEF(i).Ordinal = inOrdinal Then
            VBFunction_Description = exeVB6_APIDEF(i).uDescr
            outRName = exeVB6_APIDEF(i).uName
            Exit Function
        End If
    Next i

Else
    'by name:
   
    For i = 1 To UBound(exeVB6_APIDEF)
        If exeVB6_APIDEF(i).uName = inAPIname Then
            VBFunction_Description = exeVB6_APIDEF(i).uDescr
            
            Exit Function
        End If
    Next i
End If

VBFunction_Description = "Error API incorrect or not present in msvbvm60.dll"

End Function


