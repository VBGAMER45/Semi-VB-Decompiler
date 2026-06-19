# UnkVCall analysis — frmClient.frm (Client2.exe)

`obj.UnkVCall_<offset>h(args)` is the decompiler's render for an **unresolved indirect
vtable call** `call [vtable + offset]` — the receiver and/or the member at that vtable
offset could not be named. As of 2026-06-19 frmClient.frm has **519** of them. This
document categorises every type, what each is, why it's unresolved, and what would fix it.

> Background: a COM/VB object's vtable starts with IUnknown/IDispatch (offsets 0x00–0x18:
> QueryInterface/AddRef/Release/GetTypeInfoCount/GetTypeInfo/GetIDsOfNames/Invoke — these
> are never UnkVCall). Members begin at **0x1C**. A VB **form** vtable additionally has
> control accessors at **0x2F8 + index*4** and the form's own user methods at **0x6F8+**.
> A **user class** has its methods at 0x1C+ (and higher for many-membered classes).

## Summary by receiver category

| Category | Count | Resolvable? |
|---|---:|---|
| (empty receiver) `.UnkVCall` | 239 | Hard — receiver chain broke |
| object parameter `arg_*.UnkVCall` | 153 | Needs object-param type inference |
| user-class instance global `global_0055xxxx.UnkVCall` | 43 | Needs global typing + class vtable map |
| intrinsic control `frmClient.<ctl>.UnkVCall` | 37 | Needs VB6.OLB vtable-offset→member map |
| `_Global` object `global_00560E54.UnkVCall` | 26 | Extend the _Global offset map |
| Winsock OCX `frmClient.Winsock1.UnkVCall` | 21 | Needs OCX-typelib vtable-offset→member map |
| **Total** | **519** | |

---

## 1. Empty receiver — `.UnkVCall_<off>h` (239) — HARDEST

The leading `.` means the `this` register was **not tracked** at the call site, so neither
the receiver nor the member resolves. Top offsets: 0x24(37), 0x1C(17), 0x30(15), 0x20(14),
0x3C(13), 0x2C(13), 0x14(12), 0x54(10). These are mostly user-class methods / control
properties where the receiver object came from something the value model doesn't carry: a
reused/clobbered register, a deref we don't model, a control/object obtained via an
unhandled load (e.g. the 0xA1 short form — now handled only for the narrow predeclared-class
case), or a property-Get retbuf temp. **Fix:** per-pattern receiver tracking; this is the
long tail (each sub-pattern is its own small fix). No single lever.

## 2. Object parameter — `arg_8.UnkVCall` (151) + `arg_10` (2) = 153

`arg_8` is the first parameter of a helper/packet proc (or the masked Me of an event
handler) that is an **object whose class is not recovered** — so the vtable offset can't map
to a member. Example: `modMain_Add_Carry_Item` (4B3160) calls `arg_8.UnkVCall_24h(var_24)`.
Offsets cluster at 0x5AC/0x5A4/0x8C/0x7C/0x74/0x6C/0x54/0x24 (a control/class member block).
**Fix:** parameter-type inference for object params (these procs take a class/control
instance; once `arg_8` is typed, the class/control vtable map names the offset). Ties into
the deferred transitive-param-typing work (DEFERRED_PLANS #6).

## 3. User-class instance global — `global_0055xxxx.UnkVCall` (43)

Calls on module-level class-instance globals (clsX). Offsets 0x6F8/0x6FC/0x700/0x704/0x708/
0x710/0x718 are **user-class method offsets** (>0x6F8); lower ones (0x2E4) are class
properties. Receivers: global_0055DCB0/DCD8/DD28/DD3C/DB8C/E030/… **Fix:** type the global
(NVObjClass at its `__vbaNew`, or gNativeGlobalClass from a resolved call elsewhere, or the
indirect-predeclared path just added for clsBitmap) → the class vtable map resolves the
offset. The predeclared-clsBitmap case (49d91d1) is the template; these other globals need
their class identified.

## 4. Intrinsic control — `frmClient.<ctl>.UnkVCall` (37)

Receiver resolves (lblSkillValue/lblSkillName/lblStrength/lblLevel/vscrollSkill/…) but the
**member offset needs VB6.OLB**. Dominant offset **0x40 (29) = the control-array Item
accessor** — `frmClient.lblSkillName.UnkVCall_40h(i, …)` should be `lblSkillName(i).…`
(the control-array pre-pass that handles this in Dungeon didn't match these sites). 0x54(5)
= Caption (`_Default`). **Fix:** (a) extend NativeDetectControlArrays to these element-access
shapes; (b) a VB6.OLB vtable-offset→member map (cTypeInfo/GetProperty) for the remaining
Label/control members.

## 5. `_Global` object — `global_00560E54.UnkVCall` (26)

The standalone VB `_Global` object. Known slots already render as statements elsewhere:
**0xC = Load, 0x10 = Unload** (the 18× 0x10 + 7× 0xC here are Load/Unload sites whose
*argument* resolved but the call still printed UnkVCall — e.g. `Load frmClient.Winsock1`
renders, but a few variants print `global_00560E54.UnkVCall_10h(ctl)`). 0x788(1) = another
_Global member. **Fix:** route all 0xC/0x10 through the existing Load/Unload statement
renderer regardless of arg shape, and add the other _Global slots to
NativeGlobalMethodByOffset (App/Screen/Clipboard accessors live here too).

## 6. Winsock OCX — `frmClient.Winsock1.UnkVCall` (21)

Early-bound vtable calls on the Winsock OCX (offsets 0x28/0x30/0x38/0x40/0x48/0x50/0x68/0x70…).
The **late-bound** path (`__vbaLateIdCall` + DISPID → OCX typelib) is already handled, but
these are **early-bound** `call [vt+off]` — the OCX typelib is keyed by DISPID/memid, not
vtable offset, so the offset doesn't map directly. **Fix:** build a vtable-offset→member map
from the OCX typelib (the funcs are laid out in vtable order after IDispatch), then resolve
like the intrinsic-control case.

---

## Offset quick-reference (frmClient observations)
- `0x0C`/`0x10` on `_Global` → Load / Unload.
- `0x40` on a control → control-array Item accessor → `ctrl(i)`.
- `0x54` on a control → `_Default` (Label Caption).
- `0x1C`–`~0x2E0` → object member methods/properties (class/control/OCX, per receiver typelib).
- `0x2F8 + i*4` → form control accessors.
- `0x6F8+` → form / user-class user-method offsets.

## Priority to reduce the 519
1. **Object-param typing (#2, 153)** — biggest single lever; ties to param-type inference.
2. **User-class global typing (#3, 43)** — extend the just-added predeclared/indirect typing.
3. **Control-array Item 0x40 (#4, 29)** — extend NativeDetectControlArrays.
4. **VB6.OLB + OCX vtable-offset→member maps (#4/#6)** — name resolved-receiver members.
5. **_Global slot map (#5, 26)** — small, mechanical.
6. The 239 empty-receiver tail (#1) is per-pattern and slowest.
