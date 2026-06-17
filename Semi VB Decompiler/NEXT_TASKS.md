# Semi VB Decompiler — Native Decompiler Roadmap

Next tasks / ideas for improving the **native** decompiler output (modNativeToVB.bas),
benchmarked against the Dungeon test program. Written to survive a context reset.

> **See `DEVELOPMENT.md`** for the exact build/test/git workflow and the current
> prioritized next-task ideas (local `Dim` type inference, LoadImage-position
> constants, more loop coverage, SEH-tail return rename, recompile-ability triage).
> This file is the detailed history (what's done, what's unrecoverable).

## Test setup (how to build & verify)

- **Source under test:** `C:\Users\Owner\Desktop\forummods\rpgwo\DungeonFateSource\Dungeon.exe`
  (original `.frm`/`.bas`/`.cls` source is in that same folder — the ground truth).
- **Three-way compare dirs:**
  - Ours (latest run): `C:\Users\Owner\Desktop\websites\dungeondecomipler\ourdecompiler_v28`
  - Commercial decompiler: `C:\Users\Owner\Desktop\websites\dungeondecomipler\commercialdecompiler`
  - Original source: the DungeonFateSource folder above
- **Build** (VB6 CLI; exe lands in `Install Folder\SemiVBDecompiler.exe` per `Path32`):
  loop up to ~8x (TLB-load flakiness): kill VB6, `VB6.EXE /make "VBDecompiler.vbp" /out build_check2.log`,
  break when the exe's LastWriteTime changes. Success line: "Build of 'SemiVBDecompiler.exe' succeeded."
- **Headless decompile:** `& ".\Install Folder\SemiVBDecompiler.exe" "<target.exe>" /vbp /out <dir>`
- **Regression gate every change:** proc counts unchanged
  (`grep -cE "^(Private|Public) (Sub|Function)"` per file), known-good files
  (frmMainMenu.frm, clsBitmap.cls, modSound.bas, modMain.bas) byte-identical unless intended,
  no new `<arg>` placeholders, no new dangling GoTos.
- **Key structure reference:** `C:\Users\Owner\Downloads\Alex_Ionescu_vb_structures(2).pdf`
  (VB image internal format — Object Info §8, Public Object Descriptor §7, etc.). Use it instead of guessing struct layouts.

## Hard limits — UNRECOVERABLE (do not attempt; verified stripped from native EXEs)

- Module-level **global variable names** (rendered `global_XXXXXXXX`).
- **User Sub/Function names** in standard modules (rendered `proc_XXXXXX`).
- **UDT field names** and **local variable names** (rendered `var_XX` / `arg_XX`).
- Module **constant names AND values** (inlined + stripped; grep confirms 0 occurrences).
- Commercial decompiler hits the same walls. Form/control names, class method names,
  Declare names, and event-handler names ARE recoverable (and already are).

## Control arrays — `lblSkillName(i)` / `picCarry(i)` — DONE 2026-06-17

`Update_Skills` (40E730) and `Update_Carry` (40F2F0) reconstruct their control-array
element accesses now (`lblSkillName(i).Caption`, `lblSkillName(i).ToolTipText`,
`picCarry(i).Cls`, `picCarry(i).hDC`, `picCarry(i).Refresh`) - **beating commercial**,
which leaves these as `UnkVCall`/`.Default`.  Dungeon UnkVCall 62 -> 42; 18 element
`Set var_X = Form.ctrl(idx)` statements recovered.  No regressions (proc counts /
`<cond>` / `<arg>` / GoTo unchanged, only frmMain.frm changed, all sentinels identical).

Decoded control-array-object vtable offsets (verified): **0x40** = the array Item /
element accessor (`<arrayCtrl>.UnkVCall_0040h(i)` IS `<arrayCtrl>(i)`); the form
accessor (vtable 0x2F8+idx*4) returns the ARRAY object, 0x40 on it returns element i.
**0x54** = element DEFAULT property put (Label Caption); **0x19C** = ToolTipText put.
`lblSkillName.Count` (0x4C on the array object) is still an unresolved member.

**How it works (the deterministic pre-pass the prior live-state attempts needed):**
1. **is-array flag** persisted into `gControlNameArray.bIsArray` (set in
   frmMain.ProccessControls from `bCtlArray`).  But the robust gate is by NAME:
   `NativeNameIsControlArray` = the (form,name) appears MORE THAN ONCE (one entry per
   defined element - e.g. 8 `picCarry` entries) OR any entry has the IID flag.  This
   matters because each element has its own form-accessor offset and only SOME carry
   the array-member IID+1 (the first element's offset 0x314 did not, 0x310 did).
2. **`NativeDetectControlArrays(col)`** pre-pass: for each `call [reg+0x40]`, recover
   ENTIRELY from the disasm (not live NVPushImm/NVReg): (a) the array control - the
   nearest preceding indirect `call [reg+offN]` form accessor (the __vbaObjSet/__vbaI4
   between them are `call [abs]`, skipped) whose offset is an is-array control; (b) the
   element retbuf local - the `lea reg,[ebp-X]` feeding the bottom-of-3 push; (c) the
   index - `NativeArrIndexTok` finds the `mov ecx,SRC` feeding the __vbaI4 coerce, where
   SRC is `[ebp-Y]` (-> var_Y) or a register mirroring a spill local
   (`NativeRegSpillLocal`, a wide back-scan: the loop counter is spilled at the loop top,
   often >16 instrs back).  Records `K&callVA -> (elemExpr, guid, retbufDisp)`.  Also
   marks the redundant array-obj `__vbaObjSet` for suppression (NVSuppressObjSet).
3. **Render**: the 0x40 call emits `Set var_<rb> = Form.ctrl(idx)`, sets
   `NVLocal(rb)=var_<rb>` + `NVLocalGuid(rb)=arrayGuid`.  The following
   `mov reg,[ebp-rb]; mov vt,[reg]; call [vt+0x54/0x19C]` then resolves to
   `var_<rb>.Caption` / `.ToolTipText` via the existing NativeControlProp path.
4. **Cached-vtable threading** (`NVLocalVtName`/`NVLocalVtGuid`): the FIRST Caption put
   caches its vtable in a local (string helpers clobber the register in between); the
   0x89 store records the slot's vtable identity and the 0x8B reload restores it, so the
   deferred put still resolves.

Still unresolved (separate, pre-existing gaps): the first ToolTipText VALUE renders
`<value>` (a SIB array-element source push, not control-array-specific); `.Count`
(0x4C on the array object); the `If i > lblSkillName.Count` condition.

### Gap A — ToolTipText `<value>` (SIB element source) — ROOT CAUSE FOUND, fix deferred
`lblSkillName(i).ToolTipText = SkillDef(sIndex).Description` pushes the value via
`mov eax,[0x423020]; mov ecx,[eax+edi*8+4]` (the Description field).  The value is lost
because `mov eax,[abs]` uses the **short-form opcode 0xA1** (mov eax, moffs32, no ModR/M)
which `NativeTrackReg` does NOT handle - so eax never becomes the array pointer
`global_00423020` and the SIB read can't identify it (drops to `<value>`).  The Caption
case worked only because it loaded the same global via `mov ecx,[abs]` (0x8B, handled).
A `Case &HA1` that mirrors the 0x8B abs-global path DOES fix it
(`ToolTipText = global_00423020((var_24-1))(4)`) **and** improves the whole program
(`<cond>` 44->35, `<value>` 3->1, `<arg>` 16->15, fixes the modSound `App.Path = 0`
mangle -> `global_004230F4 = 0`, fixes `UBound((UBound(1)+10))` -> `UBound(global_X)`).
BUT it also perturbs the SAFEARRAY struct-copy in modPlayer `modPlayer_Add` (0x4177D0):
several field copies that rendered `global_00423088(12)(N) = global_004230C8(12)(M)`
degrade to `= edx`, and a LSet helper call loses an arg.  Traced to a cascade where the
now-tracked eax-global makes a stale register in the SIB index chain look `global_`,
so the array-pointer detection (which needs exactly ONE of base/index to be `global_`)
fails.  Tried scoping NVRegIsMe + clearing movsx/movzx dest - neither fully neutralised
it.  So 0xA1 is net-positive but NOT strictly-additive; do it as a dedicated effort
that also fixes the SIB pv-detection when base+index both look like data pointers.
REVERTED 2026-06-17 (kept the codebase clean; control-array work unaffected).

### Gap B — `.Count` (0x4C on the array object) — decoded, not attempted
`If i > lblSkillName.Count` = `call [arrayVt + 0x4C]` (Count getter, value via a retbuf
out-param) then `movsx edx,word[retbuf]; cmp i,edx; setg`.  Fixable by: (1) extend
NativeDetectControlArrays to also catch offset 0x4C on an is-array receiver and fold
`Form.ctrl.Count` into the retbuf local; (2) add `movsx reg,word[local]` tracking so the
compare reads the folded value -> `If i > frmMain.lblSkillName.Count`.  Shares the movsx
tracking dependency with Gap A.

## Dungeon Form_KeyDown / Select-Case-on-Integer (2026-06-17) — partial

`frmMain.Form_KeyDown` (408DE0) is `Select Case KeyCode` rendered as a deep
`If <cond>` cascade.  Each arm compiles to
`mov ecx,imm ; call __vbaI2I4 ; cmp di, ax ; je` where `di` = KeyCode (a ByRef
Integer param `mov di,[edx]`, loaded ONCE - edi is callee-saved - and reused).
- **DONE 3de82ae**: `__vbaI2I4` now threads the numeric-literal ECX even when eax
  is stale (fixed `picLife.ScaleMode = 1` = vbTwips elsewhere).  So `ax` = the case
  constant is now available.
- **REVERTED (too risky)**: tracking the 16-bit ByRef-Integer deref (`mov di,[KeyCode]`
  → param name) + relaxing the 0x66 `cmp r16,r16` bail.  These DID resolve
  `If KeyCode = 97` (case 1 only - later arms reuse edi for `Game.pIndex` so they
  bail), BUT the di-tracking changed register resolution in OTHER procs, rewriting
  already-correct conditions into plausible-but-wrong ones (proc_4177D0:
  `If (var_24 - global_X(20)) > global_X(16)` → `If arg_8 > 0`; modItem field
  `= ecx` → `= arg_8`).  proc_X names are stripped so correctness is unverifiable.
  Reverted per the no-plausible-but-wrong rule.  To finish safely: scope the
  16-bit-deref tracking to ONLY the Select-Case `cmp di,ax` shape (a pre-pass that
  recognises the `mov di,[param] ; {mov ecx,imm; call I2I4; cmp di,ax; jcc}+` chain
  and binds di to the param for that run only), then reconstruct the if-chain as a
  real `Select Case`.  Big, structural; do with a dedicated mini-test.

## mnuFileLoad CommonDialog1 → cmdSkillRaise (2026-06-17)

- **Wrong-naming FIXED 358bb7d**: every `CommonDialog1.X` rendered as
  `frmMain.cmdSkillRaise`.  Root cause: CommonDialog1 (an OCX) is accessed at
  vtable+0x37C = control index 33, but it is ABSENT from the parsed control array
  (frmMain ControlCount=33, indices -1,1..32 - the OCX tControl is not in
  `aControlArray`; reading struct[33] past the array is garbage, so it is genuinely
  not there - indices 0 AND 33 have vtable accessors but no tControl).  With 0x2F8
  mapping nothing in that proc, the per-proc base solver fabricated a base mapping
  0x37C onto a real control (cmdSkillRaise idx 10).  Fix: a standard Form always uses
  the verified 0x2F8 base (NativeOwnerIsStdForm), so the unparsed control stays
  unresolved (`arg_8.UnkVCall_0000037Ch(...)`) instead of mis-named.  customocx
  Winsock1/RichTextBox1 unaffected (they ARE in their array, map at 0x2F8).
- **OCX control naming — DONE 2a8ff0b**: CommonDialog1 now resolves to
  `frmMain.CommonDialog1` in code (was `arg_8.UnkVCall_037Ch`).  The control-hierarchy
  parse (ProccessControls Case 255) sees the external control and `cControlHeader.cId`
  IS its vtable index (CommonDialog1 cId=33 → accessor 0x37C); paired with the property
  GUID from the matching gOcxList entry (strLibname == external class).  Skips controls
  already in gControlNameArray (customocx Winsock has a tControl - no duplicate).
- **OCX late-bound property VALUES — DONE 249eccd**: `CommonDialog1.DialogTitle = "..."`
  etc. now resolve.  Root cause was NOT the receiver (that resolved once the control was
  named) but the late-call PRE-PASS: it identified the helper via NativeApiName, which
  only decodes a DIRECT `call [iat]`.  VB caches the helper IAT in a callee-saved reg
  and calls it indirectly (`mov edi,[__vbaLateIdSt]; call edi`), so the pre-pass never
  saw it and its DISPID was never collected → bare `Call LateIdSt()`.  New
  NativeResolveCallApi traces `call reg` back to the reg's IAT load.  mnuFileLoad now
  matches commercial (DialogTitle/InitDir/FileName/DefaultExt/Flags=4100/Filter/ShowOpen);
  unresolved `Call LateId*()` program-wide: several → 0.
- **InitDir App.Path — DONE aa58564**: turned out the two "minor leftovers" were ONE
  root cause - the value is `Global.App.Path & "\saved\"`, accessed via the standalone
  VB `_Global` object (lazily `New`'d into a module global), and the App/Screen/Clipboard
  accessor resolution was gated to the FORM's vtable only (NVRegIsFormVt).  (The "0x50
  put" was actually `.Path`, not `FileName=""`.)  Recognise the `_Global` object at
  __vbaNew2 (CLSID/IID family {FCFB3D2x-A0FA-1068-A738-08002B3371B5}) and accept it as a
  receiver → `InitDir = (App.Path & "\saved\")` resolves.  Tiny residual: a FileName=""
  condition reads `If App.Path = 0` (App.Path lingering in a reused local; the commercial
  mangles that condition too).

## New test bench: VB6LangTest (2026-06-16)

`C:\Users\Owner\frogger\vb6native\LangTest\VB6LangTest.exe` (full source alongside)
exercises ~every VB6 statement/function/feature — a broader bench than Dungeon for
classes, properties, events, file I/O, and the intrinsic-function set. Findings:

- Class method sig/kind — **DONE 574391e** (FuncDesc array leading-null fix:
  `Sub Greet(arg_C)` → `Function Greet()`, `Property Let Name(NewValue)`).
- `Exit Sub` inside Function/Property — **DONE c6d4215** (now `Exit Function`/`Exit Property`).
- Option Compare Text `__vbaStrTextCmp` fold — **DONE 509e226** (the case-insensitive
  twin of `__vbaStrCmp`, materialised form).
- Built-in value-fold — **DONE 52e4fbf** (value intrinsics Environ$/QBColor/Now/Rnd/
  Format/financial/date/TypeName now fold into `lhs = X(args)`; LangTest bare-Call
  lines 604→483). EXCLUDED (kept as Calls until the value model catches up): Boolean
  predicates Is*/EOF (consumed by `test eax/jcc` - condition renderer can't build the
  relational yet, would drop the call), and FreeFile (file number reused live in eax
  as the next Open's `As #<n>` without reload - folding leaks `#FreeFile(...)`).
- **TODO return + parameter TYPES** for class methods (no `As <type>` anywhere, even
  Dungeon clsBitmap). FuncDesc has the tdesc; the field isn't located yet and the
  Ionescu doc omits the method struct. Build a mini test with varied return/param
  types, dump FuncDesc fields, map. See memory `class-method-signatures.md`.
- **TODO File I/O (#5)**: `Open ... As #<n>` still shows the path/#0 because FreeFile
  isn't folded (see above - blocked on reused-result-register tracking: re-tag a
  register to its stored local AFTER `var = expr()` so the live-in-eax file number
  reads `var_X`; the naive global re-tag collapsed legit repeated `var = func()`
  stores, so it needs to be scoped). Also: Width # renders `Call undef(...)` (helper
  not in the API DB - add `__vbaWidthFile`-style name); Print # shows a raw address.
- Condition renderer for folded predicates — **DONE 9ab06f5** (`<cond>` 66→44).
  A 16-bit `test ax,ax` of a tracked call expression (NativeIsCallExpr) now resolves
  to `<expr> <> 0`, so `If SomeFunc(args) Then` / `If IsNumeric(x) Then` reconstruct
  instead of dropping to `<cond>`.  Is* predicates re-enabled for folding.
- **TODO RaiseEvent (#3) — investigated, deferred (deep + under-sampled)**:
  `RaiseEvent NameChanged(NewValue)` renders `Call RaiseEvent(arg_8, 1, 1)`.
  Findings (LangTest Class1.Name_let @40A600):
  - ABI: `__vbaRaiseEvent(Me, 1, 1, <inline 16-byte VARIANTARG built on the stack via
    `sub esp,0x10`>)`.  cdecl, `add esp,0x1C` after.  The two `1`s are presumably
    (eventIndex, cArgs) but WHICH is which can't be disambiguated from LangTest -
    every event there is single-event / single-arg.  Need a class with 2+ events and
    a 2-arg event to nail the encoding.
  - Event NAME "NameChanged" IS in the binary (~0x4047E0) but in the COM/TLB
    REGISTRATION name pool (packed flat with the class method names Name/Greet), NOT
    referenced by any aligned pointer the way method FuncDescs are - so it needs the
    dispinterface/registration parser (modCOM territory), not the FuncDesc path, to
    map eventIndex→name.  The event PARAM name ("NewName" @4095A8) IS pointer-referenced
    (4092C0/4092D0) near the param-name arrays.
  - Also needs extracting the inline Variant arg values (vt at [esp], BSTR/val at
    [esp+4]) to render the `(NewValue)` argument list.
  Do NOT ship a guessed event name (plausible-but-wrong); gather multi-event samples
  first.
- **TODO** Class2 `Implements` member `IGreet_Greet` sig (FuncDesc under the IGreet
  iface vtable - different voff mapping than the class default interface).

## Next tasks (prioritized — value vs effort/risk)

### SAFEARRAY bounds-check suppression — DONE (this cycle, f7e5c75)
Every array element access emitted a bounds-check guard (`test ARR / je ERR ; cmp word[ARR],1
/ jne ERR ; cmp idx,cElements / jb OK ; call __vbaGenerateBoundsError`) that rendered as bogus
nested `If arr <> 0 / If arr = 1 / If (idx-lb) >= cEls` blocks. Pre-pass NativeDetectBoundsChecks
(anchored on __vbaGenerateBoundsError) suppresses the guard jumps + feeding cmp/test + skip-jmp;
handles direct `call[iat]` and register-cached `mov ebx,[iat]; call ebx`. Dungeon: <cond>
426->153, GoTo 925->448, cDims checks 360->9, output nearly halved. clsDirectSound/modMonster/
modMessage massively cleaner. Two gotchas handled: don't suppress the error call itself (it
clears the push stack - else a stray arg leaks into the next call); emit labels for suppressed
jump-target instructions (an On Error handler can land on a bounds-error call -> else dangles).

### Loop induction variables — Phase 1 DONE (a0c792e); Phase 2 = For/Do reconstruction
PHASE 1 (done): register-resident loop counters showed their stale init constant (`If 1 <= 10` for
`For var_24 = 1 To 10`). New pre-pass NativeDetectCounterSlots marks ebp stack slots WRITTEN+READ
inside a backward-branch loop (induction vars, loop-type-agnostic); the store handler binds them +
their mirroring register to the NAME var_X (not a per-iteration value), surfacing init/increment.
Result: `var_24 = 1` / `If var_24 <= 10` / `var_24 = (1 + var_24)`; 35 named-counter headers,
garbage `0<>0` 10->7. NOTE: barely moved blank `<cond>` (93->92) - the loop BODY `<cond>` are SIB
array-element word compares (`cmp word [base+idx*scale+disp],0`), a separate gap (the deferred
"SIB-indexed element index" item), NOT counter naming.
PHASE 2 DONE (ecc2e8b): reconstruct top-tested loops as `Do While <cond> ... Loop` (chosen over
extending the strict For detector - Do While is the general safe form: the increment stays visible
in the body, covers register-counter loops For misses AND genuine Do While/While-Wend). New
pre-pass NativeDetectWhileLoops: unconditional back-edge jmp to header H (single ref), H's first
cmp+forward-exit-Jcc -> emit `Do While` at the Jcc / `Loop` at the back-edge / suppress the header
label. Builds on Phase 1 (counter reads var_X, so `Do While var_24 <= 10`). Dungeon: 35 loops
reconstructed, GoTo 448->413; verified Do==Loop + If/EndIf+Do/Loop nesting balanced in every proc,
proc counts identical, known-good byte-identical, no new dangling, 15 For loops intact.
NOT done (deferred - low value / unrecoverable): 2D/UDT-array element compares (the bulk of the
remaining ~92 `<cond>`) - base is computed pvData + scaled register index; UDT field names are
stripped (memory udt-array-deref-rendering) and the COMMERCIAL decompiler punts too, so the
ceiling is opaque `global_X(i)(264)`. Also deferred: bottom-tested Do...Loop While (conditional
back-edge), For/Do Until, raw-register conds (untracked call results).

### 16-bit register-test conditions — DONE (this cycle, 89c7762)
`mov esi,[var_44]; test si,si; je` left `<cond>` (the 0x66 word-register guard). Now resolved to
`var_44 <> 0` when the register holds a clean var_/arg_/global_ ref (NativeIsCleanNamedVal): a
16-bit mov already clears the register, so a still-tracked value came from a full 32-bit load, and
VB only word-tests Integer/Boolean vars (low word = value). Dungeon: <cond> 153->93, raw-register
conds 57->11; garbage `0<>0` held at 10. IMPORTANT FINDING: the modMap "dead code after GoTo" is
NOT dead - reachability analysis showed those regions are reachable; the only unreachable code is
On Error/bounds-error handler stubs (reached via exceptions). DO NOT add dead-code elimination -
it would delete error handlers. Remaining `<cond>` (~93): register/loop-induction conditions, a
few compound/Variant-compare residuals.

### Compound relational conditions (setcc + or/and) — DONE (this cycle, 8facc95)
`If x <= 0 And x >= -10000` compiles to `cmp/setl; test/setg; or; jcc` and rendered as
garbage `If 0 <> 0`. Now: SETcc binds `(L <op> R)` from the pending compare into its register;
`or`/`and` of two relational-Boolean registers combine with ` Or `/` And ` (guarded by
NativeLooksRelational) and arm the Boolean condition for the following Jcc. clsDirectSound.SetVolume
-> `If Not ((arg_10 < -10000) Or (arg_10 > 0))`. Dungeon: garbage `0 <> 0` 77->10, ~20 compound
conditions recovered. Note: it fixed the GARBAGE-condition category, not the blank `<cond>` (still
~153) - those are a different cause (see below). Some compound operands still imperfect (0/1/self-
compares) from register-tracking gaps.

### __vbaObjIs -> `Is Nothing` — DONE (this cycle, 916aac0)
`Call ObjIs(a, 0)` + `If <cond>` now renders `If a Is Nothing` (no materialization - bind the
relational at the call; ` Is ` added to NativeLooksRelational). 7 sites.

### A. Boolean-materialization collapse (medium value, medium effort) — STILL PENDING
`If <fncall/relational> Then` compiles to a boolean materialization (neg/sbb/inc/neg or
mov1/jmp/xor) that our structurer renders as TWO nested Ifs re-expanding the same expression,
e.g. the double-`If (var_30 = "NAME=")` nesting in modItem (now that B resolves the operands,
this is the most visible remaining instance). NOTE: clsDirectSound's noise turned out to be the
bounds-check + ObjIs idioms (both done above), NOT this double-If. clsDirectSound residuals left:
array element renders `arg_8(56)(12)` not `Sounds(Index)` (private-array field name unrecoverable);
compound `NewVolume <= 0 And >= -10000` lost as `If arg_10 = 0` (setl/setg/or And-materialization);
method call on an array-element object dropped as `' call .50` (element not vtable-tracked).
Other bounds-check variants still leak as conditions (UBound-based checks, `If (i-1) >= 20`).

### B. String comparison relational `__vbaStrCmp` — DONE (this cycle)
`Call StrCmp(a, b)` + `If <cond>` now renders `If a = b`. Implemented as a pre-pass
(`NativeDetectStrCmpCompares`, modelled on `NativeDetectFpCompares`) that detects the
`__vbaStrCmp` + `neg/sbb/inc/neg` boolean materialization, suppresses the scaffolding, and
binds the equality relational into eax (the `mov REG,eax` propagates it to the tested reg).
Shipped together with (a) Line Input/Get slot zero-init invalidation and (b) string-transform
folding (Left$/Right$/Mid$/Trim$/UCase$/LCase$ now thread their result). Dungeon: 13 StrCmp
→ real relationals, 65 lost-string-args → 0, `<cond>` 426→409, garbage `0<>0` 77→58, no
regressions (known-good files identical, no new `<arg>`/dangling GoTo). Numeric operands are
rejected (a bare number is an unresolved string ptr → falls back to a visible Call).
Remaining: compound boolean (`or` of two materializations) still `<cond>` (overlaps A); the
boolean can leak into surrounding arithmetic at 4 sites (pre-existing garbage region).

### NEW. String→numeric conversion folding (`__vbaR8Str`/`R8IntI2`/`I4Str`…) — DONE (this cycle)
Config parsers showed `Call R8Str(var_30)` + `Call R8IntI2()` (13+6 sites) as dropped Calls,
losing the field STORE that consumed them. Now folded: `__vbaR8Str`/`R4Str` push `CDbl(s)`/
`CSng(s)` into the NVFpu model; `__vbaR8IntI2/IntI4` pop it and wrap `CInt(...)`/`CLng(...)`
(collapsing the redundant inner CDbl); eax-only `__vbaI2Str`/`I4Str` fold directly. The store
tracking then emits `<field> = CInt(var_30)`. Dungeon: `iNumber = Int(Trim$(Right$(...)))`
reconstructs as `var_24 = CInt(var_30)`; ItemDef array fields recover too. All bare conversion
Calls gone; no regressions (known-good identical, no new `<arg>`/`CInt(<arg>)`/dangling GoTo).

### C. Register-resident loop induction variables — DONE (ef139aa)
The simplest register-counter case is reconstructed. VB6 keeps a small For loop's counter AND limit
in registers and RELOADS the limit register at the loop top, so the back-edge targets a `mov
limitReg, imm` (not the cmp) and the compare is a 16-bit `cmp di,cx` (0x66 + 3B). NativeDetectForLoops
now: scans past leading limit-setup `mov reg,imm` to find the cmp (NativeIsMovRegImm gate);
NativeForCounterReg accepts `cmp r16,r16` (66+3B/39, md=3) + the 0x39 counter form; captures start
(NativeForStartIsConst returns the init const) and limit (new NativeForLimitVal: immediate, or the
const a limit register was loaded with via a backward mov-imm scan) into NVForStart/NVForLimit since
render-time decode can't read the cmp; suppresses the header..exit-Jcc run; binds the counter reg to
the loop var name so the body resolves `x` not its stale init. Result: customocx Winsock1_DataArrival
→ `For i = 0 To 1000 / ... CStr(i) ... / Next i` (was `Do While <cond>` / `CStr(0)`). Dungeon: For/Next
15→16, Do/Loop 35→34 (modMap `Do While var_20<=100` → `For i = 0 To 100`), garbage 25→24; <cond>/<arg>/
GoTo/proc-counts unchanged, sentinels identical. STILL DEFERRED below: the nested draw-loop case
(`For i = -radius To radius`) where the counter mirror + FPU coords couple with D; and memory/variable
register-limits (only constant register-limits resolve so far — variable limits fall back to Do While).

**Original concrete reproducer (now passing): customocx `Project1.exe` `Winsock1_DataArrival`.**
`C:\Users\Owner\Desktop\websites\customocx\Project1.exe` (source alongside; the proc decompiles now,
just the loop cond is blank). Source = `For x = 0 To 1000 / RichTextBox1.Text = RichTextBox1.Text & x & vbCrLf / Next x`.
We emit `Do While <cond>` with `x` shown as `CStr(0)` (stale init). Disasm of the header:
`xor edi,edi` (x=0, counter in **edi**); `mov ecx, 0x3E8` (limit 1000 in **ecx**);
`66 3b f9` = **`cmp di, cx`** (16-bit register-vs-register); `0f 8f .. jg exit`. So the condition is
`x <= 1000` but BOTH operands are registers (di=counter, cx=limit-imm) and the 16-bit `3B` reg-reg
`cmp` isn't decoded into a relational, so the While-loop reconstructor gets no cond. To finish:
(1) decode `cmp r16,r16` (0x66 + 0x3B) in NativeDecodeCompare; (2) recognise edi as a register loop
counter (extend NativeDetectCounterSlots / a new NativeDetectRegCounter: a reg `xor`'d to 0 (or `mov
reg,imm`), read+`inc`/`add reg,1`'d inside a back-edge loop, compared at the header) and bind it to a
NAME so the cond reads `<ctr> <= 1000` and the body's `x` resolves (not `CStr(0)`); (3) ideally emit
`For <ctr> = 0 To 1000 / Next` instead of `Do While`. GOTCHA: at decompile time (linear) the reg holds
its INIT value at the header compare, so naively reading NVReg gives `0 <= 1000` (the old Phase-1
stale-init bug) — the counter must be bound to a name BEFORE the cond renders.

Also (the harder, original case): nested draw loops (`For i = -radius To radius`) keep counters in
registers and show `If -((picView.ScaleHeight/32)/2) <= ((picView.ScaleHeight/32)/2)`. Same root;
gated on the same reg-counter tracking. The simple register->local mirror (committed earlier) named
some counters but not these. Couples with D (FPU integer-expression tracking) for the draw-loop coords.

### D. FPU integer-expression tracking (medium value, high effort — coupled with C)
`picView.Line (var_100, var_F8)-(31, 31), 255` shows float-temp coords where commercial got
`var_20*32`. Needs tracking `imul reg,reg,imm` (×const) and `fild`/`fstp` carrying the integer
expression through the FPU model. The base value is usually a register loop counter, so this is
gated on C. Note: the `neg` op is intentionally NOT tracked (an attempt mis-rendered VB's
boolean-idiom neg as `-(handle)` in 19 places — reverted).

### E. Line/Circle/PSet Step + flags (low value, low effort — beats commercial)
`Line (x1,y1)-(x2,y2), color` currently drops the `Step` (relative 2nd coord) and `B`/`F` flags
(commercial also drops them). The flag byte (e.g. 0x2E) encodes Step/B/F; reverse the bit mapping
to emit `-Step(...)`, `, B`, `, BF`. Would beat commercial. (Code: NativeControlProp `Line` case.)

### F. FPU compare `st0` operand recovery (low value, medium effort)
`If (st0 > (var_C + 5))` — the left operand is lost because a value-preserving fp helper `call`
between the `fld` and the `fcom` resets the NVFpu stack model (NVFpuTop=0). Let such helper calls
carry the operand through. ~13 sites.

### G. Sub Main detection (low value, low effort)
VB header `+0x2C lpSubMain` gives the entry point — label it. Quick win.

### H. Form-property fidelity (low-medium value, medium effort)
Reconstructed `.frm` form-layout properties differ from source (e.g. `MaxButton = -1` vs `0`;
missing `StartUpPosition`, `FontName`). Improve extraction from Optional Object Info §9 / Control Info §10.

### I. Countdown / bottom-tested loops (low value, medium effort)
For-loop detection misses `For i = N To 1 Step -1` (jl exit + dec) and bottom-tested loops.

### J. Loose ends (low value)
- 2 pre-existing `New clsBitmap` leaks into BitBlt args (As-New `__vbaNew` result leaking into a push).
- Functions called as STATEMENTS with return used elsewhere — edge cases of the value-fold (D done for the common case).

## What's already DONE this cycle (so you don't redo it)

Proc-overrun fix (jo/jno) · For/Next reconstruction · SAFEARRAY element conditions + lea fix ·
FPU comparison conditions · user-class method calls + `New <class>` (via Object Info chain) ·
module-level `Public global_X` declarations · API `Declare` block (once, Public, in first module) ·
Code-tab declaration display · control method calls as real statements (`Line (x1,y1)-(x2,y2),color`) ·
orphan `loc_` label strip (−1437 lines) · coercion-helper fold (I2I4/FpR4/FpI4/CastObj) ·
register-counter naming via stack mirror · As-New private form-field method/property resolution ·
value-returning Function folding (Property Get + Function via retbuf detection) ·
string-comparison reconstruction (B: __vbaStrCmp→relational + string-transform folding +
Line Input/Get zero-init invalidation).

### DONE 2026-06-16 cycle (OCX / customocx Project1.exe focus — see auto-memory)
OCX form-property decode via IPersistStreamInit::Load + TLI (modOcx.bas: _ExtentX/Y/Version,
real prop values, Font BeginProperty block, invisible Left/Top trailer, TextRTF→frx,
OleObjectBlob fallback) · OCX control EVENT names via the coclass DefaultEventInterface
(modCOM.OcxEventSig; Winsock1_Connect/DataArrival/Error) + the aEvent-gate fix (native event
names were dropped when the P-Code aEvent field read 0) · event-handler bodies use the declared
PARAM NAMES (Number/Description not arg_C) · form-self property puts (`Form1.Caption = "Hey"` via
the _Form GUID + a global→form-instance map so it works from a .bas module too) · late-bound
property PUT VALUE + late-bound METHOD ARGS recovered from the Variant DATA fields
(`RichTextBox1.Text = "Connected"`, `Winsock1.Connect "127.0.0.1", 535`) + Variant-build-line
suppression · FRAMELESS function discovery (E8 call sites + boundary filter, so module functions
aren't missing) AND simple frameless BODY decode (`proc = (arg_8 * arg_C)`, typed Function).
**Remaining from that cycle → task C above** (register-counter For-loop `<cond>` in DataArrival).

See the per-topic notes in the auto-memory dir
(`...\memory\*.md` — MEMORY.md is the index) for implementation details and gotchas.
