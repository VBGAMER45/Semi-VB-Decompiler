# Semi VB Decompiler — Development & Handoff Guide

How to build, test, and commit changes to the **native decompiler**
(`modNativeToVB.bas` and friends), plus the prioritized next-task ideas. Read this
first when resuming work. The prioritized task roadmap (what's done, what's left)
lives in **`NEXT_TASKS.md`** — this file is the *how*, that one is the *what*.

---

## 1. Build

VB6 is installed; the engine is itself a VB6 project (`VBDecompiler.vbp`).

**Command line build** (output goes to `Install Folder\SemiVBDecompiler.exe`, per
`Path32` in the `.vbp`):

```
& "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make "VBDecompiler.vbp" /out "build_check2.log"
```

**Gotchas (important):**
- VB6 intermittently fails with *"Missing or not registered VB6TMPL.TLB"* and does
  **not** update the exe. Retry: `Get-Process VB6 | Stop-Process -Force`, wait ~2s,
  re-run. A loop of up to ~12 attempts reliably gets a build.
- `build_check2.log` is **append-mode** — its trailing line can be a stale
  "succeeded" from a prior run. **Never trust the log alone.** Confirm success by the
  exe's `LastWriteTime` changing.
- A real compile error logs `Build of 'SemiVBDecompiler.exe' failed.` + the error.
- Close any open VB6 IDE instance first (a second `/make` while the IDE has the
  `.vbp` open errors with "invalid command line argument").

**Use the helper script** — `build_and_test.ps1` (repo root) wraps the retry loop,
verifies the exe timestamp actually changed, distinguishes a real compile error from
TLB flakiness, then runs a decompile:

```
powershell -ExecutionPolicy Bypass -File build_and_test.ps1 -OutDir "<output dir>"
# -NoBuild to skip the build and just re-run the decompile
```

---

## 2. Test (decompile + regression gate)

**Decompile headlessly** (no GUI):

```
& ".\Install Folder\SemiVBDecompiler.exe" "<target.exe>" /vbp /out "<output dir>"
```
Writes `Form1.frm`, `*.cls`, `*.bas`, a `Dungeon.vbp`, and `decompile.log` into the dir.

**The benchmark** is the Dungeon Fate game:
- Target EXE: `C:\Users\Owner\Desktop\forummods\rpgwo\DungeonFateSource\Dungeon.exe`
  (original `.frm`/`.bas`/`.cls` source is in that same folder — the **ground truth**).
- Three-way compare dirs under `C:\Users\Owner\Desktop\websites\dungeondecomipler\`:
  - `ourdecompiler_v29` — our latest run (the `build_and_test.ps1` default `-OutDir`).
  - `commercialdecompiler` — a commercial decompiler's output (a *reference*, not
    always correct — we have beaten it in several places).
  - the DungeonFateSource folder — the original source.

### Regression workflow (do this every change)
1. **Snapshot a baseline before editing:** copy the current `ourdecompiler_v29` to a
   baseline dir (e.g. `baseline_<feature>`) so you can diff *just your change*. (The
   session used a frozen `baseline_pre_strcmp`; for an isolated diff, snapshot right
   before the change instead.)
2. Make the change, build, decompile into `ourdecompiler_v29`.
3. **Gate — these must all hold (PowerShell):**
   - **Proc counts identical** per file: `^(Private|Public) (Sub|Function)` counts
     match the baseline. A changed count means a proc was lost/merged — a real bug.
   - **No new `<arg>` placeholders** and **no new dangling GoTos** (a `GoTo loc_X`
     with no matching `loc_X:`). Count unique `(file|target)` dangles vs baseline.
   - **Garbage `0 <> 0` / `0 = 0` must not rise.**
   - For structural changes: **`If`/`End If` + `Do`/`Loop` block nesting balanced
     in every proc**, and **Do == Loop / For == Next counts** per file.
   - **Known-good sentinels** — files that *shouldn't* change unless intended:
     `frmMainMenu.frm`, `modSound.bas` stayed byte-identical all session;
     `modMain.bas` only gained the Win32 `Public Const` block; `clsBitmap.cls`
     changed intentionally (param names, returns, raster-ops). When a feature *should*
     touch a sentinel, verify the diff is **only** the intended improvement.
4. **Spot-check the actual output** against the original source for the feature you
   changed — this is how the "plausible but wrong" bugs get caught (e.g. the greedy
   arithmetic fold that produced `(arg_8(96) + 4)` / `(var_38 + 1)`).

Useful counters to watch over time (whole-program): `<cond>` (blank conditions),
`<arg>`, `GoTo`, garbage `0 <> 0`. Trajectory: `<cond>` 426→92→**65** (feada43:
UDT/array-element compares resolve to `global_X(12)(244) <op> R`), GoTo 925→413,
garbage 77→7, output roughly halved. The remaining ~65 `<cond>` are the hard
ceiling: boolean-materialization (`test ax,ax`, Task-A-structural) + word reg-reg
juggling (resolving yields garbage).

---

## 3. Git workflow

- **Work directly on `master`. Do NOT create feature branches** for this repo (solo
  project; the owner asked for this explicitly).
- **Source-only commits:** stage only the engine source (`*.bas`, `*.cls`, `*.frm`).
  Do **not** commit the built `Install Folder\SemiVBDecompiler.exe` or
  `VBDecompiler.vbw` (build/IDE artifacts — they show as modified; leave them).
- **Verify before committing:** the regression gate above must pass. Commit one
  improvement at a time (each is bisectable).
- **Commit message:** explain the root cause + the fix + the verification. End with:
  ```
  Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>
  ```
- **Push** to `origin/master` once verified (confirm with the owner first unless told
  otherwise). Example one-liner used all session:
  ```
  git add "modNativeToVB.bas" && git commit -m "..." && git push origin master
  ```

---

## 4. How the engine is structured (orientation)

`DecompileNativeProcToVB(addr)` in `modNativeToVB.bas` decompiles one proc:
- **Pre-passes** over the instruction collection mark idioms by VA before rendering:
  `NativeDetectForLoops`, `NativeDetectWhileLoops`, `NativeDetectFpCompares`,
  `NativeDetectStrCmpCompares`, `NativeDetectCounterSlots`, `NativeDetectBoundsChecks`,
  `NativeDetectReturnSlot`. Each is *strict* — a miss leaves the raw form, a false
  match is the danger.
- **Per-instruction rendering** (`NativeProcessInst`) tracks a symbolic value model:
  `NVReg(0..7)` (register values), `NVLocal` (stack-slot expressions), `NVPushImm`
  (the pending argument stack), `NVCmp*` (the pending compare for the next Jcc).
- **Finalization** (line ~486): `NativeStripOrphanLabels(NativeSubstituteConstants(
  NativeSubstituteArgNames(output)))` — post-passes that substitute parameter names /
  return-value name / Win32 constants by whole-identifier token replacement.
- **Two-pass output:** `BuildNativeCodeCache` (modNative.bas) decompiles **all** procs
  into `gNativeCodeCache` first; then `modOutput.bas` writes the modules from the
  cache (so things discovered during decompile — e.g. used Win32 constants in
  `gUsedWin32Const` — are complete by write time, emitted once in the first standard
  module like the API `Declare` block).

Key reference: `C:\Users\Owner\Downloads\Alex_Ionescu_vb_structures(2).pdf` (VB image
internal format — Object Info, typeinfo FuncDesc/VarDesc). Use it instead of guessing
struct layouts. **UDT type/field names and user-defined constant names are stripped
from native EXEs — unrecoverable (verified by binary string search).**

---

## 5. Next ideas (prioritized)

See `NEXT_TASKS.md` for the full done-list and the hard/unrecoverable limits. The
concrete next candidates:

### A. Local `Dim` declarations with type inference — DONE (443472f, 2026-06-15)
`NativeInsertLocalDims` post-pass emits a sorted `Dim var_X As <type>` block after
each proc header, types inferred from usage: String (string funcs/`&`/literal),
Long/Integer (Len/UBound/CLng/CInt/arithmetic), control class (`Set v=frmMain.picView`
→ PictureBox/Label/... via gControlOffset), `clsX` (`Set v=New clsX`), else Variant.
49% of Dungeon's 507 locals concretely typed (beats commercial's blanket Variant).
Runs after arg/const substitution so the return-slot/param names aren't var_X.
See memory `local-dim-inference.md`. Follow-up: transitive `var=otherVar` propagation;
ties into the recompile-ability triage (E).

### B. LoadImage-position / general API flag constants (Tier-2 continuation)
Now that arg truncation is fixed, `LoadImage(…, 0, …, 16)` shows the flags. Add a
data-driven `(API, arg#) → constant family` table so `0`→`IMAGE_BITMAP`,
`16`→`LR_LOADFROMFILE` (+ their `Public Const`s via the existing emit path), scaling
to `ShowWindow` `SW_*`, file-open modes, window styles. Safe because the position
pins the family. Mostly filling a table; the substitution + Const-emit framework
already exists (`NativeSubstituteConstants` / `GetWin32ConstBlock`).

### C. More loop coverage
Top-tested `Do While` is done. Still missing: **bottom-tested `Do … Loop While`**
(conditional back-edge — the `Loop While <cond>` path exists but `NVLoopHdr` is never
populated for it), **`Do Until`**, and true **`For … Next`** for the counted loops
(cleaner than `Do While` + explicit increment). Builds on the Phase-1 counter naming.

### D. SEH-tail return rename (small polish)
`NativeDetectReturnSlot` finds the retbuf→local copy, but when that copy sits **past
the `ret`** in the SEH tail (e.g. `LoadBitmap`), it isn't matched, so the proc shows
`var_18` instead of `LoadBitmap`. Extend the scan to the post-`ret` epilogue region.

### E. Recompile-ability triage (high-leverage meta-task)
Try compiling the reconstructed `.vbp` in VB6 and let the **actual errors** drive
priorities (missing `Dim`s → A, `Property Let`/`Set` gaps, malformed expressions)
rather than guessing. The single best signal for what's worth fixing next.

### Known hard limits (don't chase)
- Remaining `~92 <cond>` are mostly **2D/UDT array element compares** (`pvData +
  scaled register index`) — UDT field names are stripped and the commercial punts
  too; ceiling is opaque `global_X(i)(264)`.
- The boolean-materialization double-`If` collapse (Task A) is structural and
  regression-prone.
