============================================================
 VB Scanner  -  companion tool for Semi VB Decompiler
============================================================

WHAT IT DOES
------------
Recursively scans a folder (and every sub-folder) for EXE, DLL and
OCX files, identifies the ones built with Visual Basic 4, 5 or 6 or
with the .NET runtime, lists them with their detected runtime, and
lets you hand the selected file straight to Semi VB Decompiler.

Files that are not VB4/5/6 or .NET (native C/C++, packed stubs, etc.)
are simply skipped, so the list only shows things the decompiler can
actually work with.


HOW TO USE
----------
1. Open VBScanner.vbp in Visual Basic 6 and press F5 (or compile it
   to VBScanner.exe via File > Make VBScanner.exe).
2. Click "Browse..." and pick the folder to scan.
3. Tick the file types you care about (EXE / DLL / OCX) and the
   runtimes you want listed (VB4 / VB5 / VB6 / .NET).
4. Click "Scan". Progress is shown at the bottom; click "Stop" to
   cancel a long scan.
5. Select a row and click "Send to Semi VB Decompiler" (or just
   double-click the row). The file is opened in the decompiler.

The path to SemiVBDecompiler.exe is shown at the bottom. It is auto-
detected as ..\SemiVBDecompiler.exe (the parent folder, where the
main project lives). Use the "..." button to point at a different
copy; the choice is remembered in VBScanner.ini next to the program.


HOW DETECTION WORKS
-------------------
The scanner reads each file's PE header directly, mirroring the
detection in Semi VB Decompiler (modPeSkeleton.bas):

  .NET   -> the PE optional-header data directory #14 (the COM/CLR
            descriptor) has a non-zero address.
  VB4    -> the import table references VB40032.DLL
  VB5    -> the import table references MSVBVM50.DLL
  VB6    -> the import table references MSVBVM60.DLL

The file is passed to the decompiler exactly as the decompiler's own
command line expects:

  SemiVBDecompiler.exe "<full path to file>"

(see modCmdLine.bas in the main project for the full CLI, which also
supports /out <dir>, /vbp, /dism and /solution).


LIMITATIONS
-----------
* Only 32-bit PE files carry the VB4/5/6 runtimes. Genuine 16-bit
  VB1/2/3 and VB4-16 programs are NE-format, not PE, and are not
  detected here.
* .NET files are reported simply as ".NET"; the language (VB.NET vs
  C#) is not distinguished, matching the main decompiler's behaviour.


REQUIREMENTS
------------
* Visual Basic 6 (to open/compile the project).
* mscomctl.ocx (the ListView control) registered - it already ships
  with Semi VB Decompiler, so no new dependency is introduced.
============================================================
