# Deferred Plans — Native Decompiler

Tasks investigated and intentionally deferred (too risky / too large for a quick fix),
written so they can be resumed cold. See `DEVELOPMENT.md` (build/test/git), `NEXT_TASKS.md`
(long history), and the auto-memory dir for the rest. Benchmarks: Dungeon
(`...\forummods\rpgwo\DungeonFateSource\Dungeon.exe`, byte-stable ground truth) and
Client2 (`...\websites\a-client2-decomp\Client2.exe`).

---

## 1. clsPktSkillDef.Desc — For-loop byte-fill reconstruction (Client2, @5361F0)

**Status: investigated 2026-06-19, NOT shipped. The `movsx` sub-fix was tried and REVERTED
(regressed Dungeon integer math — see below).**

### The real code (from the binary)
```vb
Public Property Let Desc(inString)
    For i = 1 To 100                                      ' di/edi counter, 1..0x64
        If i > Len(inString) Then                         ' cmp esi,eax(=__vbaLenBstr); jle else
            field_44(i + 33) = 0                          ' xor ecx; __vbaUI1I2 -> al=0; mov [pvData+esi],al
        Else
            field_44(i + 33) = Asc(CStr(Mid(inString, i, 1)))
            ' rtcMidCharVar(var_3C,var_4C,i,var_2C)=Mid$(inString,i,1);
            ' __vbaStrVarVal -> CStr; rtcAnsiValueBstr -> Asc; __vbaUI1I2 -> CByte; store
        End If
    Next i
End Property
```
`field_44` is the packet byte buffer = `pvData` of the inline array at `Me+0x38` (descriptor
+0x0C; same relationship as the GetData/ID-stamp work, commit 37752d0).

### What our output is missing (4 interlocking failures)
1. **Counter renders as `0`** (`If 0 > Len`, `Mid(..,0,..)`). `i` lives in `edi`; the body reads
   it via `movsx esi, di`, but the value model (`NativeTrackReg`, ~8156) has NO `Case &HF`, so
   `esi` keeps its stale `xor esi,esi` 0.
2. **Two-armed if/else not reconstructed.** The `jle` to the else branch isn't structured; the
   THEN body `field_44(i+33)=0` is dropped and rendered as `If 0 > Len Then GoTo loc_00536318`.
3. **Asc/Mid store dropped.** The else branch collapses to bare `Call Mid(...)`; the
   `Asc(CStr(Mid(...)))` value and the byte store are lost.
4. **Dangling label / `Next` placement** — a SYMPTOM of #2: the unreconstructed if/else leaves a
   `GoTo loc_00536318` + label (the merge point) before `Next`. Reconstruct the if/else and the
   label/GoTo vanish and `Next` is clean. NOT a separate "Next bug."

### The `movsx` attempt (reverted — DO NOT re-ship the broad form)
Added a `Case &HF` to `NativeTrackReg` to propagate `NVReg(rm)->NVReg(reg)` for `movsx/movzx`
reg-to-reg (md=3), gated to a "simple" source value (no `(`/space/`&`/quote). Result:
- Client2: FIXED cleanly — `If i > Len(inString)`, `Mid(.., i, ..)`; all counters flat.
- Dungeon: REGRESSED integer-math procs — `proc_413050` expressions mangled (two distinct operands
  both became `arg_C`), `field_70 = (Index + var_A8)` -> `field_70 = 0`, modMap/modPlayer/modItem
  perturbed. `movsx`-word propagation helps loop counters but corrupts 16-bit Integer arithmetic,
  and the two are **indistinguishable at the instruction level** (the documented "word reg-reg
  juggling" hard ceiling). Even the simple-val gate didn't separate them. Reverted to 82b457d.

### What a SAFE fix needs (the real task)
A **bespoke, tightly-gated pre-pass** that recognises this whole For-loop-byte-fill idiom end to
end and emits it as a unit (like the array-Variant / byte-stamp pre-passes), rather than touching
the general `movsx`/if-else paths:
- Anchor on the shape: `For`-counter (already detected) + `cmp esi,Len; jle` two-arm + both arms
  ending in `mov [pvData+esi],al` (one via `__vbaUI1I2` of 0, one via `__vbaUI1I2(Asc(StrVarVal(rtcMidCharVar(...))))`).
- Bind `esi` to the counter name `i` LOCALLY within the matched region (avoids the global movsx
  regression).
- Emit `field_<arrayOff>(i+33) = 0` / `field_<arrayOff>(i+33) = Asc(CStr(Mid(inString, i, 1)))`
  inside reconstructed `If i > Len(inString) ... Else ... End If`, suppressing the scaffolding.
- MANDATORY: Dungeon byte-identical + Client2 no new `<arg>`/`<cond>`/garbage.
High effort, niche payoff (packet `Desc`-style setters). Only worth it if a clean gate is found.

---

## 2. Function-result-via-out-param folding — `var_60` -> `Command` (frmClient, Client2)

From the 82b457d work: `If ("1024lUnAtIc1024" = var_60)` should be `If (Command = "1024lUnAtIc1024")`.
`Call Command(var_60)` leaves the result in the by-ref out-param `var_60`; we don't fold that into
the following compare. General pattern: a runtime/intrinsic call whose result is delivered through a
by-ref local, then that local is read. Folding `var_X -> <call>` after `Call F(var_X)` is the task —
broader than the string fix, needs care (the local may be read multiple times / reassigned).

---

## 3. Packet byte-buffer field stores — SetData / Randomizer Let / loaders (Client2)

From the GetData/ID-stamp work (8fa3737 / 37752d0). The byte buffer `field_44` (= `field_38.pvData`)
is written by `SetData`, the `Randomizer` Let helper (`proc_536860`), and `proc_536D5B` via
`mov [pvData+index], al`, but those stores are DROPPED (bodies render empty / partial). Recovering
them needs general field-pointer deref tracking: `mov edx,[Me+0x44]` should mark `edx` as a deref
base so `mov [edx+k], al` renders `field_44(k) = ...`. Two blockers in the 0x88 byte-store handler:
(a) it requires `disp > 0` (the stamp store is `mov [edx],al`, disp 0); (b) `edx` (a deref of the
Me field) isn't tracked as a deref base. General field-ptr tracking is broad/risky (touches every
`mov reg,[Me+disp]`) — do as a tightly-gated pre-pass, not a global model change.

---

## 3b. Late-bound control property get/put chain — frmClient.Form_Resize — DONE 1ccc876

**Status: DONE 2026-06-19 (commit 1ccc876).** All three pieces shipped:
(1) stock-extender DISPID table Left/Top/Width/Height = 0x80010003..6 (NativeStockExtenderMember,
consulted before the OCX typelib) - fixed 12 wrong gets (HideSelection->Height) AND 18
__vbaLateIdSt puts (Call LateIdSt -> obj.Member = value); (2) __vbaR4Var/R8Var -> CSng/CDbl
FPU fold (Call R4Var 16->0); (3) stock-extender Ld result tracked in eax so `push eax; R4Var`
inlines `CSng(var_X)` (gated to stock gets so a Dungeon CommonDialog StrCmp path stays
byte-identical). Client2: only frmClient+frmChannel changed, all counters/proc-counts flat;
Dungeon byte-identical. Form_Resize now matches commercial. Original analysis kept below.

**Original analysis (investigated 2026-06-19):**

The repeated block
```
var_40 = frmClient.txtMessage.HideSelection     ' WRONG member (DISPID mis-resolved)
Call R4Var(frmClient.txtMessage)                ' __vbaR4Var = CSng (Single coercion)
var_14(4) = var_4C ; var_14(8) = st0 ; var_14(12) = var_44   ' build a VT_R4 Variant
Call LateIdSt()                                 ' __vbaLateIdSt = late-bound property PUT (dropped)
```
should be (commercial):
```
frmClient.txtGlobalMsg.Height = CSng(frmClient.txtMessage.Height)
```
(Commercial writes `CSgn(...)` — that's its label for the **Single** coercion `__vbaR4Var`;
real VB is `CSng`. There is no `CSgn` keyword.)

**Decode (disasm @486CB0-486D66):** the txtMessage/txtGlobalMsg/... controls are accessed
LATE-BOUND. `mov eax,[esi=Me]; call [eax+0x644/0x648/0x64c/...]` are the FORM-vtable control
accessors (these controls sit at high vtable offsets, not the 0x2F8 block) — they return the
control object, cached via `__vbaObjSet` (edi) into var_28/var_2C. Then:
- GET: `__vbaLateIdCallLd` reads `txtMessage.Height` (a Variant) -> `__vbaR4Var` -> `fstp` Single.
- A 16-byte VT_R4 Variant (vt=4 at [esp], the Single at +8) is built inline (`sub esp,0x10`).
- PUT: `__vbaLateIdSt` sets `<targetCtrl>.Height = <that Variant>`. The `0x80010006` push is the
  late-bind flags/lcid word.

**Two bugs to fix (both in the existing late-bound path NativeLateIdCall/NativeDetectLateCalls):**
1. **DISPID mis-resolves to `HideSelection` instead of `Height`.** `Top`/`Left`/`Height`/`Width`
   are EXTENDER/stock properties (provided by the VB container, not the RichTextBox OCX typelib),
   so `modCOM.LateMemberName(RichTextBox, dispid, ...)` finds a coincidental same-memid member
   (HideSelection). FIX: a stock/extender-DISPID table (Top/Left/Height/Width/Visible/Enabled/...)
   consulted BEFORE the control's OCX typelib.
2. **`__vbaLateIdSt` PUT renders `Call LateIdSt()` (unresolved).** The object token and/or DISPID
   aren't recovered for these chained puts (object is cached in var_28 via __vbaObjSet; the value
   is the inline VT_R4 Variant). NativeLateIdCall's St branch exists but needs the obj token from
   the cached-control local and the value from the inline Variant; also fold `__vbaR4Var` -> CSng.

HIGH effort, HIGH risk (late-bound changes touch Winsock/RichTextBox across the program). Do as a
dedicated effort with heavy regression on the customocx + Client2 benches. Net payoff is large for
readability (Form_Resize is ~80 mangled lines that should be ~25 clean property assignments).

## 3c. Indirect predeclared-instance clsBitmap receiver — DONE 2026-06-19

**Status: DONE.** Pre-pass NativeDetectIndirectNew records the pointer-to-pointer As-New
global -> class (gated to user CLASSES, not forms); a narrowly-gated 0xA1 handler tags only
those globals' loads as a field-address whose deref is the object.  Resolved util0/util1
`var_10C.LoadBitmap(...)` + 30 `clsBitmap.ImageDC`/`.MaskDC` (inlined into BitBlt as the
source DC); receiver renders as `clsBitmap` (matches commercial).  Client2 UnkVCall 1649->1618,
counters/proc-counts unchanged, Dungeon byte-identical (the class-only gate + narrow 0xA1
avoid the modPlayer SIB regression that broad 0xA1 tracking causes).  Original analysis below.

**Original analysis (investigated 2026-06-19):**

`clsBitmap` is a PREDECLARED-instance class (commercial renders `clsBitmap.LoadBitmap(clsBitmap, ..)`).
Its instance is reached through a 3-LEVEL indirection: `mov eax,[0x55d148]` (slot ptr) ->
`mov ecx,[eax]` (object ptr) -> `mov [ebp-X],ecx`; later `mov eax,[ebp-X]; mov ecx,[eax]
(vtable); call [ecx+0x30]`. The standard As-New model is 2-level (global -> object -> vtable),
and __vbaNew2's @dest here is a REGISTER (`push eax` where eax=[0x55d148]), not a `push imm`
global address, so NVObjClass never records the class for 0x55d148. To fix: recognise the
predeclared-instance global (0x55d148 -> clsBitmap, e.g. from the __vbaNew2 ObjInfo 0x416264
+ the pushed global VALUE), thread the type through the extra indirection level, and render the
receiver as the class name `clsBitmap`. Niche; the string args + the 2-level cases already work.

## 4. Indexed-property parameter drop (Client2, medium value)

`Property Get OK(Index As Integer)` / packet `Size(Index As Integer)` render as `OK()` / `Size()` —
the indexed-property param is dropped. The FuncDesc param-count `nP = (b0\4) - b1` computes 0 for
these (b0=5, b1=1). Decode the VB6 FuncDesc indexed-property encoding (low 2 bits of b0 seem to flag
the invoke kind). See the FuncDesc field dump in SESSION_HANDOFF_2026-06-19.md task C.

---

## 5. Form `Property Get` that should be `Sub` (frmMainMenu.Recv_*, Client2)

`Recv_Player_List` etc. render `Private Property Get` but are `Public Sub`. The form FuncDesc is
UNRELIABLE PER-FORM (frmMainMenu's Recv_* decode with property+return bits wrongly, BUT frmClient's
Send_* are reliable). A blanket "gate forms out of FuncDesc kind" REGRESSED frmClient (already
reverted). Needs a per-method reliability signal, OR extend the method-link resolver to forms (forms
mix own methods at voff 0x1C with events at 0x6F8). Do NOT re-attempt the blanket gate.

---

## 6. Transitive parameter typing (both projects)

**OBJECT variant — DONE 2026-06-19 (commit 9a1bd58).** A helper/module proc calling vtable methods
on an object PARAMETER (`arg_X.UnkVCall_<off>h`, class stripped from typeinfo) is now typed from the
class its callers consistently pass. New pass `NativeBuildParamObjClasses` (runs after the first
render pass in `BuildNativeCodeCache`): scans rendered bodies for object params + their call sites,
resolves the position-matched arg's class (local `New`/`As`, or As-New field via `gFormFieldClass`),
and — on SINGLE-class consensus + `NativeClassHasMethodAt` validation — records `gParamObjClass`;
`BuildNativeCodeCache` re-renders only those helpers. The ByRef param register is tagged
`NVRegFieldCls` (pointer-to-object = same two-deref shape as an As-New field address), so the existing
user-class-method path resolves `arg_X.Method`; the `umRecv` visible-statement gate now fires for
`arg_` receivers too. Client2: `arg_*.UnkVCall` 405→352 (53 real packet-property names, e.g.
`arg_8.ImageType`/`.ItemID`/`.Quantity`; modTexture 21→0, modMain 122→90); Dungeon byte-identical.
Conservative gates (single-class, method-existence, proc_<hex> helpers, cls-prefixed classes) mean a
polymorphic helper / unresolved arg / struct-base `arg_X` stays untyped (verified 0 polymorphism).

**STILL OPEN (deferred — lower yield, see verification 2026-06-19):**
- Transitive/fixpoint chains (`arg_X` passed onward to another helper) and untyped-local args
  (`Set var = Func()`); the no-caller form helpers (`frmNPCTrade.Add_Player_Item` etc., 33 sites) have
  NO caller in the output so propagation can't help — need form method-link resolution instead.
- STRING/scalar variant: `FirstChar(s) = LeftN(s, 1)` — propagate a String/Long type across calls (an
  arg passed as the Nth arg to a proc whose Nth param is typed). Same propagation skeleton, different
  consumer (param TYPE not object class). Also: positional string args inside `InStr(1, s, needle)`.

---

## Hard ceilings — confirmed, do NOT attempt (see condition-resolution.md)
- Broad `movsx`/`movzx` counter propagation (regresses Integer math — task 1 above).
- `Abs()` compares needing CFG data-flow (Open_Sight_Line).
- 2D/UDT array-element compares — UDT field names stripped (commercial punts too).
- Boolean-materialization double-If collapse (Task A) — structural, regression-prone.
- Module global / user-proc / UDT-field / local / constant NAMES — stripped from native EXEs.
