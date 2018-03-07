-------------------------------
Semi VB Decompiler - VisualBasicZone.com
Version: 0.09
Build: 1.0.64
Website: http://www.visualbasiczone.com/products/semivbdecompiler
-------------------------------
Contents
1. What's New?
2. Features
3. Questions?
4. Bugs
5. Contact
6. Credits

1. What's New?

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

3. Questions?
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
   A. Head over to http://www.vb-decompiler.com  tons of information on vb decompiling.

4. Bugs
     Some properties aren't handled yet such as dataformat
     P-Code decoding may hang use the disable P-Code option under options.
     If you would wish to report a bug email me at
     support@visualbasiczone.com
     Please include as much information as possible so we can try to fix it and even better send us the file if possible.

5. Contact/Support
     Email=support@visualbasiczone.com
     Semi VB Decompiler Website:
     http://www.visualbasiczone.com/products/semivbdecompiler/

6. Credits
     I would like to thank the following people for helping me with this project.
     Sarge, Mr. Unleaded, Moogman, _aLfa_, Alex Ionescu, Warning and many others.

