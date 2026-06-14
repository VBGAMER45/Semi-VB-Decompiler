-------------------------------
Semi VB Decompiler - VisualBasicZone.com
Version: 2.0
Build: 2.0.7
Website: http://www.visualbasiczone.com/products/semivbdecompiler
-------------------------------
Contents
1. What's New?
2. Features
3. Command Line Options
4. Questions?
5. Bugs
6. Contact
7. Credits

1. What's New?

   Version 2.0 Build 2.0.7
   .Net decompiler rebuilt.  It now reconstructs classes, fields, properties and methods to C# and VB.NET (in addition to a full IL disassembly).  The reconstructed types are browsable under a new ".Net Classes" node in the project tree, with C#, VB.NET and IL views per class and the method list for navigation.  Method bodies are reconstructed including if/else, while/do-while, switch and try/catch/finally, falling back to labeled goto for control flow that cannot be structured.  Added File > Build .Net Solution to export a Visual Studio solution (.sln with C# and VB.NET projects, one source file per class).  Added a headless command-line mode for decompiling and generating projects/solutions without the UI - see "3. Command Line Options" below.  The .Net helper DLL is no longer required (its BitConverter helpers are now implemented in pure VB6).

   Version 2.0
   Native decompiler now finds procedure offsets for modules and classes, recovers .bas module procedures, shows them in the project tree, exports decompiled code into the generated project, and adds a raw disassembly (Dism) tab.

   Version 0.09
   Added a new tool. Api Add allows you to add Api's to the Semi VB Decompiler Api Database.

   Version 0.08 Build 1.0.64
   Updated Native Procedure Decompile dissembles faster and added some native dissemble options to the options screen. Also updated decompile from offset, it now verify the files has a VB5! signature.
   For .Net applications added the view console under the Tools menu.  Added data directories to the PE Optional Header list.

   Version 0.07 Build 1.0.63
   Minor GUI updates and fixes. Fixed VBP external component bug.

   Version 0.07 Build 1.0.61
   Redid the P-Code property name finder procedure for the standard toolbox controls.
   Now the VTables are now pulled from VB6.olb the typelib file instead of having them hardcoded.

   Version 0.07 Build 1.0.60
   Switched back to my old sytle control property editing. Added saving the old value, so when you switch controls, the changed value is shown.

   Version 0.07 Build 1.0.59
   Added over 100 more P-Code properties to the database

   Version 0.07 Build 1.0.58
   Fixed a problem loading with some VB Dll's
   Included support for Windows XP Styles

   Version 0.07 Build 1.0.57
   Fixed some VB Detection Problems.
   Added option to show offset off P-Code String in the P-Code String List

   Version 0.07 Build 1.0.53
   New icons to indicate the control type in the treeview.
   Using a property sheet control for property editing and viewing now.
   Memory Map Generated way faster
   Added View Report menu under the Tools Menu.

   Version 0.06C Build 1.0.52
   Improved P-Code decompiling.
   Api calls are shown as library name . Function Name
   For VCallHresult properties are recovered and the object type.
   Fixed Extra opcodes after end of procedure for P-Code.

   Version 0.06B Build 1.0.50
   Updated VB 1/2/3 Binary Form To Text supports all default controls properties.
   Redid the menus for all the applications.   The menu's now include Bitmaps for some items, and are now styled.
   Improved handleing of non vb files, in the treeview.
   Will detect if the file is protected by the UPX packer.
   Added Startup Form Patcher, you can choose which form appears first!
   Other improvements here and there.

   Version 0.06A Build 1.0.46
   Added support for User Documents.
   Better Control Processing for unknown opcodes, errors recorded to a file.
   Added P-Code String List.
   Added Type Library Explorer program.

   Verion 0.06A Build 1.0.44
   Added a new tool to convert VB 1/2/3 binary forms to text.
   VBP File output now includes the thread flags.
   P-Code output is now semi-colored and is in bold.
   P-Code To VB Code is now colored as well.
   Redid part of the Control Property editing functions works better.
   Took out the ComFix.txt file and just included the information in the decompiler.

   Version 0.06 Build 1.0.21
   Detection for all versions of vb from 1 to vb.net
   Ne Format exe's now shown in FileReport.txt
   Partial VB4 Support Added.
   Fixed Backcolor property on labels VB5/6
   Now correctly decompiles VB5 dll files.
   Fixed too many Ends for Menu's.
   Faster form processing.
   .Net Structures now shown under .Net Structures.
   Added more .Net processing now shows Strings, Blobs, Guids, and User Strings

   Version 0.05A Build 1.0.20
   Faster syntax coloring.
   More detailed filereports containing pe information.

   Version 0.05 Build 1.0.19
   Added VB.net Detection and shows the CLR header.
   Better handling of PE imports and VB version detection for other versions.

   Version 0.05
   Correct events identified for all common controls.
   Control Arrays now handled correctly.
   Minor fixes here and there.

   Last Version 0.04C Build 1.0.15
   Added Advanced Decompile which you can decompile a vb project
   by offset, could be used against packed exes, packed with upx and other compressors.
   Added Native Procedure Decompile a work in progress.
   Added Object= includes to forms and project file.
   Added Export Procedure List, Add Address to P-Code Procedure Decompile.
   Now shows PE information even if its not a valid vb exe.

   Version 0.04C
   Added update checker.
   Added vbcode to P-Code procedure decompile.
   Added memory offset to file offset.
   Improved Form and control decoding.
   Fixed Ocx loading bug.
   Fixed Empty .frx generation. Fixed null subs in form/class generation.
   Thinking about working on VB4 and VB3 support for the next version.

   Version 0.04
    Improved P-Code Decompiling a lot.
    Better ObjectType detection.
    Added control property editing.
    Added VB5 Support!

   Version 0.03
     P-Code decoding started and image extraction.
     Numerous bug fixes.
     Event detection added.
     Dll and OCX Support added.
     External Components added to vbp file.
     Begun work on a basic antidecompiler.
     Form property editor, complete with a patch report generator.
     Procedure names are recovered.
     Api's used by the program are recovered.
     Msvbvm60.dll imports are listed in the treeview.
     Syntax coloring for Forms.
     Fixed scrolling bug.

   Version 0.02
     Rebuilds the forms
     Gets most controls and their properties.

   Initial Release version 0.01

2. Features
     Decompiling the P-Code/native vb 4/5/6 exe's, dll's, and ocx's
     Decompiling .Net assemblies to C# and VB.NET, with a full IL disassembly
     Browse reconstructed .Net classes (C# / VB.NET / IL) in the project tree
     Build a .Net solution (.sln with C# and VB.NET projects) from an assembly
     Command line / batch mode for headless decompiling and project generation
     Form Generation
     Resource extraction wmf, ico, cur, gif, bmp, jpg, dib
     Control/Form Editor
     Startup Form Patcher for VB 5/6
     Address to File Offset converter.
     P-Code Event/Procedure Decompile
     Native Event Disassembly
     Shows offsets for controls and allows you to edit the control properties.
     Decompile a file from an offset useful against packed exe's using compression such as upx.
     Multilanguage support including Dutch, French, German, Italian, spanish and more
     Memory Map of the exe file, so you can see what's going on.
     Advanced decompiling using COM instead of hard coding property opcodes.

3. Command Line Options

   Semi VB Decompiler can run without its window (headless) so you can
   decompile files and generate projects from a script or batch file.

   Usage:
     SemiVBDecompiler.exe <inputfile> [/out <dir>] [/vbp] [/dism] [/solution]

   <inputfile>   The VB 4/5/6 or .Net exe, dll or ocx to decompile.  This is
                 the first argument that is not a switch.
   /out <dir>    Directory to write the generated VB project or .Net solution
                 into.  If omitted it defaults to the file's dump folder.
                 (The intermediate decompile output - IL, file report, P-Code,
                 images - is always written under the dump\<filename> folder
                 next to the program, the same as the GUI.)
   /vbp          Generate a VB project (.vbp plus all forms, modules, classes,
                 user controls, property pages, user documents and designers).
   /dism         Generate a VB project that uses the raw native disassembly for
                 each procedure instead of the decompiled VB code.
   /solution     For a .Net assembly, build a Visual Studio solution containing
                 a C# project and a VB.NET project, one source file per class.
   /?            Show usage.

   Notes:
   - The correct action is chosen automatically for the file type: a VB 4/5/6
     file always produces a .vbp project and a .Net assembly always produces a
     solution, regardless of which generate switch is passed.
   - With no generate switch the file is only decompiled (its dump folder is
     populated); no project or solution is written.
   - No message boxes or prompts appear in command line mode.  A decompile.log
     is written to the output directory, the program returns an exit code of 0
     on success or non-zero on error, and then closes automatically.
   - Quote any path that contains spaces.
   - Run the program from (or alongside) its install folder so it can find its
     data files (languages, API list, vb6.olb, required OCX/DLLs, etc).

   Examples:
     SemiVBDecompiler.exe "C:\Apps\MyApp.exe" /out "C:\Out" /vbp
     SemiVBDecompiler.exe "C:\Apps\MyApp.exe" /out "C:\Out" /dism
     SemiVBDecompiler.exe "C:\Libs\MyNet.dll" /out "C:\Out" /solution
     SemiVBDecompiler.exe "C:\Apps\MyApp.exe"

4. Questions?
   Q. What about Native Code Decompiling?
   A. It is in the works. Right now I have offsets for all the events and can do a disassembly of each event but I need to work on an assembly to VB engine still.
   Q. What the heck are the P-Code Tokens?
   A. P-Code tokens is the last step before turning the P-Code into readable VB Code.
      All you have to do now is link the imports of the exe with the functions in P-Code.
   Q. Why does it not show all the controls on my forms?
   A. If it is not a common control found in the toolbox then we can not get extra information it, in the future we maybe able to process these controls.
      Another reason can be because it is a property that is not detected by COM using vb6.olb.
   Q. Why doesn't it get my procedure names for Modules?
   A. Visual Basic only saves procedures names for Form's and Classes.  And it only saves them for forms if they are public.
   Q. How does this decompiler work?
   A. First it gets all the main vb structures from the exe.
      Next it gets all the controls properties via COM using vb6.olb
   Q. What files does this decompiler require?
   A. It requires the following files:
      TLBINF32.dll
      comdlg32.OCX
      RICHTX32.OCX
      MSCOMCTL.OCX
      TABCTL32.OCX
      MSFLXGRD.OCX
      MSINET32.OCX
      Msvbvm60.dll
      SSubTmr6.dll
      WinSubHook2.tlb
      pePropertySheet.ocx
      cPopMenu6.ocx
      And VB6.olb version 6.0.9
      All of the above files need to be registered(the installer should auto register the files.)
      If you are examining a .Net file then you need to have the .Net framework installed.
   Q. Where can I learn more about Visual Basic 5/6 Decompiling?
   A. Head over to https://sandsprite.com/vb-reversing/  tons of information on vb decompiling.

5. Bugs
     Some properties aren't handled yet such as dataformat
     P-Code decoding may hang use the disable P-Code option under options.
     Please include as much information as possible so we can try to fix it and even better send us the file if possible.

6. Contact/Support
     Semi VB Decompiler Website:
     http://www.visualbasiczone.com/products/semivbdecompiler/

7. Credits
     I would like to thank the following people for helping me with this project.
     Sarge, Mr. Unleaded, Moogman, _aLfa_, Alex Ionescu, Warning and many others.
