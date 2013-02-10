Attribute VB_Name = "modGlobals"
Option Explicit

' main.txt - global definitions
Type T1257
  M1265 As Integer
  M126E As Integer
  M1277 As Integer
  M127E As String * 9
  M1289 As Integer
  M1294 As Integer
End Type

Type T129E
  M12AE As Byte
  M12B8 As Byte
  M12BD As Integer
  M12C5 As Integer
  M12CF As Integer
  M12D9 As String * 1
  M12E2 As Integer
End Type

Type T1369
  M1374 As Integer
  M1380 As Integer
  M138C As Integer
  M1397 As Integer
End Type

Type T13C2
  M1374 As Integer
  M13CF As Integer
  M13DD As Integer
  M13E4 As Integer
End Type

Type T1415
  M1422 As Integer
  M142D As Integer
  M1435 As Integer
  M143F As Integer
  M1448 As Integer
  M13DD As Integer
End Type

Type T1467
  M1422 As Integer
  M1471 As Integer
End Type

Type T148E
  M149E As String * 3
  M13DD As String * 1
  M14A9 As String * 1
  M14B1 As Long
End Type

Type T14B9
  M14C8 As Long
  M14D1 As String * 1
End Type

Type T14DB
  M13DD As Integer
  M14E8 As Integer
  M1374 As Integer
  M14F2 As Integer
  M14FC As Integer
  M1506 As Integer
  M1511 As Integer
  M1519 As Integer
  M1521 As Integer
  M1527 As Integer
  M12E2 As Integer
  M152D As Integer
  M1538 As Integer
  M1541 As Integer
  M154C As Integer
  M1558(15 To 17) As Integer
  M1561 As Integer
  M156B As Integer
  M1577 As Integer
  M1582(21 To 22) As Integer
  M158D As Integer
  M1599 As Integer
  M15A4 As Integer
  M15AE As Long
End Type

Type T15B8
  M13DD As Integer
  M14E8 As Integer
  M1374 As Integer
  M14F2(3 To 6) As Integer
  M15C5 As Integer
  M15D0(8 To 12) As Integer
  M15D6 As Integer
  M15DF As Integer
  M15E5 As Integer
  M15F0(16 To 28) As Integer
  M12D9 As Integer
  M15F7 As Long
End Type

Global Const Version = "3.59" 'Version of DoDi's decompiler
Global Const Language = "e" 'Language of DoDi's decompiler
Global Const gc0322 = "Error Situation VB3 Decompiler OCX"
Global Const gc0326 = "For this program you'll need an upgraded VB3 Decompiler OCX"
Global Const gc032A = "Severe errors may cause the Decompiler to crash"
Global Const gc032E = "Internal problems, the code created may be buggy"
Global Const gc0332 = "Do you have the latest edition of VB Decompiler?"
Global Const gc0336 = "Runtime Error in VB Decompiler"
Global Const gc033A = "News from Semi VB Decompiler"
Global Const gc033E = "You may send this program to vbgamer45 to improve VB Decompiler"
Global Const gc0342 = "Error "
Global Const gc0346 = "This option is available only in the Professional version"
Global Const gc034A = "Found unknown data structures!"
Global Const gc034E = "Not a Visual Basic program"
Global Const gc0352 = " is not supported"
Global Const gc0356 = "Found an unknown resource!"
Global Const gc035A = "Found unknown fixups!"
Global Const gc035E = "Error in Decompiler logic!"
Global Const gc0362 = "The program is too big for this version of VB Decompiler"
Global Const gc0366 = "Missing description for "
Global Const gc036A = " contains unknown structure!"
Global Const gc036E = "Found an unknown token!"
Global Const gc0372 = "Unexpected variable reference!"
Global Const gc0376 = "Found incompatible scopes!"
Global Const gc037A = "Found incompatible types"
Global Const gc037E = "Found an unknown collection!"
Global Const gc0382 = "An already known problem occured"
Global Const gc0386 = "File not found or wrong version: "
Global Const gc038A = "Name too long"
Global Const gc038E = "Program without code - abort"
Global Const gc0392 = "Visual Basic not found"
Global Const gc0396 = "File not found"
Global Const gc039A = "Cheat!"
Global Const gc039E = "Already done"
Global Const gc03A2 = "VB Decompiler"
Global Const gc03A6 = "Initializing"
Global Const gc03AA = "Select EXE file"
Global Const gc03AE = "Scanning forms"
Global Const gc03B2 = "Scan finished"
Global Const gc03B6 = "Creating project files"
Global Const gc03BA = "Creating declarations"
Global Const gc03BE = "Now save *.FRM as text, please"
Global Const gc03C2 = "Project created"
Global Const gc03C6 = "Combining forms und code"
Global Const gc03CA = "Decompilation finished"
Global Const gc03CE = "Open "
Global Const gc03D2 = "Loading "
Global Const gc03D6 = "Scanning "
Global Const gc03DA = "Creating "
Global Const gc03DE = "Forms"
Global Const gc03E2 = "Modules"
Global Const gc03E6 = "Segments"
Global Const gc03EA = "Scopes"
Global Const gc03EE = "Types"
Global Const gc03F2 = "Tokens"
Global Const gc03F6 = "Fixups"
Global Const gc03FA = "Data"
Global Const gc03FE = "local "
Global Const gc0402 = "global "
Global Const gc0406 = "Subroutine calls"
Global Const gc040A = "Global Declarations"
Global Const gc040E = "Declarations in "
Global Const gc0412 = "Error module offset"
Global Const gc0416 = "Rename"
Global Const gc041A = "Startform not found"
Global Const gc041E = "No more errors"
Global Const gc0422 = "Specify type!"
Global Const gc0426 = "Unknown type"
Global Const gc042A = "Different modules selected"
Global Const gc042E = "Unexpected EXE fixup"
Global Const gc0432 = "Token uses no variable"
Global Const gc0436 = "Show Variable"
Global Const gc043A = "Source is not saved in binary format"
Global Const gc043E = "Form must be saved 'As Text'"
Type T19BA 'MS Dos Header
  M19CA As Integer
  M19D5 As Integer
  M19DE As Integer
  M19E7 As Integer
  M19F1 As Integer
  M19FC As Integer
  M1A07 As Integer
  M1A12 As Integer
  M1A1C As Integer
  M1A26 As Integer
  M1A30 As Integer
  M1A3A As Integer
  M1A44 As Integer
  M1A4E As Integer
  M1A58(15) As Integer
  M1A5F As Long
End Type

Type T1A6A 'NE Format
  M19CA As Integer
  M1A7A As Integer
  M1A84 As Integer
  M1A96 As Integer
  M1A26 As Long
  M12D9 As Integer
  M1AA8 As Integer
  M1AB2 As Integer
  M1ABE As Integer
  M1A30 As Integer
  M1A3A As Integer
  M1AC9 As Integer
  M1AD3 As Integer
  M1ADD As Integer
  M1AE6 As Integer
  M1AEF As Integer
  M1AF8 As Integer
  M1B01 As Integer
  M1B0A As Integer
  M1B12 As Integer
  M1B1A As Integer
  M1B22 As Long
  M1B2B As Integer
  M1B37 As Integer
  M1B40 As Integer
  M1B4D As Integer
  M1B59 As Integer
  M1B62 As Integer
  M1558 As Integer
  M1B6B As Integer
End Type

Global Const gc05D2 = 23117 ' &H5A4D% 'MS Dos Signiture
Global Const gc05D4 = 17742 ' &H454E% 'NE Format Signiture
Type T1BCE
  M12D9 As String * 1
  M1BDB As Integer
End Type

Type T1BE3
  M12D9 As String * 1
  M1BF0 As Integer
  M1BF9 As String * 1
  M1BDB As Integer
End Type

Type T1C1C
  M12D9 As Integer
  M1C2B As Integer
  M13E4 As Integer
End Type

Global gv064C() As T1C1C
Global gv067E As Integer
Global gv0680() As Integer
Global gv06B2 As Integer
Global gv06B4 As Long
Global gv06B8 As Long
Global gv06BC As Integer
Global gv06BE As Integer
Type T1C86
  M1C92 As String * 1
  M1C99 As String * 1
  M13E4 As Integer
  M1CA1 As Integer
  M1CA7 As Integer
End Type

Type T1CAD
  M13E4 As Integer
  M1CB7 As Integer
  M1C99 As Integer
  M1CA1 As Integer
  M1CA7 As Integer
End Type

Global gv0726() As T1CAD
Global gv0758 As Integer
Type T1CCE
  M1CDD As Integer
  M1CE7 As Integer
  M12D9 As Integer
  M1CF1 As Integer
End Type

Global gv0784() As T1CCE
Global Const gc07BE = 256 ' &H100%
Type T1D5D
  M1D68 As String
  M1D70 As Integer
End Type

Global gv07E4() As T1D5D
Global gv0816 As Integer
Global gv081A() As T1D5D
Global gv084C As Integer
Global gv084E() As String
Global gv0882 As String
Global gv0886 As String
Global gv088A As String
Global gv088E As Integer
Global iVBVersion As Integer 'VB Version
Global gv0894 As T19BA
Global gv08D6 As T1A6A
Global gv0916 As Integer
Global gv0918 As Integer
Global gv091A As Integer
Global gv091C As Integer
Global gv091E As Integer
Global gv0920 As Long
Global gv0924 As Long
Global gv0928 As Long
Type T1E54
  M1E3E As Long
  M1E62 As Integer
  M1E6C As String
End Type

Global gv0956() As T1E54
Global gv0988 As Integer
Global gv098A As Integer
Global gv098C() As Integer
Global gv09BE As Long
Global gv09C2 As Integer
Global gv09C4() As Integer
Global gv09F6 As Long
Global gv09FA As Integer
Global gv09FC() As Integer
Global gv0A30 As String * 1
Type T1EDA
  M1EE3 As String * 2
End Type

Global gv0A46 As T1EDA
Type T1EEE
  M1EF7 As Integer
End Type

Global gv0A5A As T1EEE
Type T1F02
  M1EE3 As String * 4
End Type

Type T1F0B
  M1F14 As Long
End Type

Type T1F19
  M1F22 As Integer
  M1F28 As Integer
End Type

Global gv0A98 As T1F02
Global gv0A9E As T1F0B
Type T1F40
  M1F4C As Variant
  M1F51 As Integer
  M1F59 As Integer
  M034F As String * 8
  M1F62 As String * 10
  M1F6A As Integer
  M1F72 As Integer
  M1F79 As Integer
  M1F80 As Integer
  M1F87 As Integer
  M1F8E As Integer
End Type

Global gv0B16 As T1F40
Global gv0B70 As Integer
Global gv0B72 As Integer
Type T20E0
  M12B8 As String * 1
  M1EF7 As Integer
End Type

Global gv0B8E As Long
Global gv0B92 As Long
Global gv0B96 As Long
Global gv0B9A
Global gv0B9E As Long
Global gv0BA2 As Long
Global gv0BA6 As String
Global gv0BAE As String
Global gv0BB2 As String
Global gv0BB6 As String
Global gv0BBA As String
Global gv0BC0 As Form
Global gv0BC4 As Integer
Global gv0BC6 As Integer
Global gv0BC8 As Integer
Global Const gc0BCC = 1 ' &H1%
Global gv0BCE As Integer
Global Const gc0BD2 = 1 ' &H1%
Global gv0BD8 As Integer
Global gv0BDA As Long
Global gv0BDE As String
Type T21FD
  M13DD As Integer
  M220A As Integer
End Type

Global gv0BFC As T21FD
Type T221B
  M2225 As Integer
  M12B8 As String * 1
  M222D As String * 1
End Type

Type T2234
  M223F As String
  M2246 As String
End Type

Global gv0C4A As T2234
Global gv0C52 As Integer
Global gv0C56 As Control
Global gv0C5A As String
Global gv0C5E As Integer
Type T227A
  M13DD As Integer
  M2283 As Integer
  M1558 As Long
End Type

Type T228A
  M13DD As Integer
  M2283 As Integer
  M2295 As Integer
  M1D68 As String
End Type

Global gv0CD0() As T228A
Global gv0D02 As Integer
Global gv0D04(16) As String
Type T238C
  M13E4 As Integer
  M2395 As Integer
  M12D9 As Integer
  M239C As Integer
  M1558 As Long
End Type

Type T23A2
  M13DD As Integer
  M2395 As Integer
  M12D9 As Integer
  M239C As Integer
  M13E4 As Integer
  M1D68 As String
End Type

Global gv0D8C() As T23A2
Global gv0DBE As Integer
Global gv0DC0 As Integer
Global gv0DC2 As Integer
Global gv0DC6 As T23A2
Global gv0DD4 As Integer
Global gv0DDA As Long
Global gv0DDE As Long
Global gv0DE2 As Integer
Global gv0DE4 As Integer
Global gv0DE6() As Integer
Global Const gc0E18 = 1 ' &H1%
Global Const gc0E1A = 2 ' &H2%
Global Const gc0E1C = 3 ' &H3%
Global Const gc0E20 = 1 ' &H1%
Global Const gc0E22 = 6 ' &H6%
Global Const gc0E24 = 1 ' &H1%
Global Const gc0E26 = 2 ' &H2%
Global Const gc0E28 = 3 ' &H3%
Global Const gc0E2A = 5 ' &H5%
Global Const gc0E2C = 6 ' &H6%
Type T249F
  M24AC As Integer
  M1374 As Integer
  M24B4 As Integer
  M24BE As Integer
  M12C5 As Integer
  M12BD As Integer
  M24C8 As Long
End Type

Type T24D3
  M24E0 As Integer
  M24AC As Integer
  M13E4 As Integer
  M1C2B As Integer
  M1374 As Integer
  M24E7 As Integer
  M24F0 As Integer
  M24F9 As Integer
  M1CB7 As String
  M2503 As Integer
  M250C As Integer
  M2517 As String
End Type

Global gv0EE2(7) As Integer
Type T2522
  M13DD As Integer
  M12BD As Integer
  M12C5 As Integer
  M12CF As Integer
  M12D9 As String * 1
  M12E2 As Integer
  M252C As Integer
  M2535 As Integer
End Type

Global gv0F42(256) As String
Global gv0F5A(256) As T2522
Global gv0F70 As Integer
Global gv0F72 As Integer
Global gv0F74 As Integer
Global gv0F76 As String
Global gv0F7A As String
Global gv0F7E As String
Global gv0F82 As Integer
Global gv0F84 As Integer
Global gv0F86 As Integer
Global gv0F88 As Integer
Type T25B7
  M25C6 As Integer
  M25D0 As Integer
  M25DA As Integer
  M1D68 As String
End Type

Global gv0FBC(255) As T25B7
Global gv0FD2 As Integer
Type T25FD
  M25C6 As Integer
  M1D68 As String
End Type

Global gv0FF6() As T25FD
Global gv1028 As Integer
Type T2627
  M2635 As Integer
  M263C As Integer
End Type

Global gv1042() As Integer
Global gv1076() As T2627
Type T265D
  M1E3E As Long
  M266C As Integer
  M2672 As Integer
  M2678 As Integer
  M267E As Integer
  M2684 As Integer
  M268A As String
  M2690 As Integer
  M2696 As String
  M269C As Integer
End Type

Global gv110E() As T265D
Global gv1140 As Integer
Global gv1142 As String
Global Const gc1146 = 2 ' &H2%
Global gv1148 As Integer
Global Const gc114A = "?pmlgcfOTas"
Global Const gc114E = 1 ' &H1%
Global Const gc1150 = 2 ' &H2%
Global Const gc1152 = 3 ' &H3%
Global Const gc1154 = 4 ' &H4%
Global Const gc1156 = 5 ' &H5%
Global Const gc1158 = 6 ' &H6%
Global Const gc115A = 7 ' &H7%
Global Const gc115E = 9 ' &H9%
Global Const gc1164 = "~.[asf14|c"
Global Const gc1168 = 1 ' &H1%
Global Const gc116A = 2 ' &H2%
Global Const gc116C = 3 ' &H3%
Global Const gc116E = 4 ' &H4%
Global Const gc1170 = 5 ' &H5%
Global Const gc1172 = 6 ' &H6%
Global Const gc1174 = 7 ' &H7%
Global Const gc1176 = 8 ' &H8%
Global Const gc1178 = 9 ' &H9%
Global Const gc117A = 10 ' &HA%
Global Const gc117C = "t%&!#@vOT*A$4|"
Global Const gc1180 = "t%&!#@vOT*A$4|"
Global Const gc1184 = "t%&!#@vOT*A$4|1"
Global Const gc1188 = "t%&!#@vOT*A$4|1U"
Global Const gc118C = 1 ' &H1%
Global Const gc118E = 2 ' &H2%
Global Const gc1190 = 3 ' &H3%
Global Const gc1192 = 4 ' &H4%
Global Const gc1194 = 5 ' &H5%
Global Const gc1198 = 7 ' &H7%
Global Const gc119A = 8 ' &H8%
Global Const gc119C = 9 ' &H9%
Global Const gc11A0 = 11 ' &HB%
Global Const gc11A2 = 12 ' &HC%
Global Const gc11A4 = 13 ' &HD%
Global Const gc11A6 = 14 ' &HE%
Global Const gc11A8 = 15 ' &HF%
Global Const gc11AA = 16 ' &H10%
Global gv11AC(31) As Long
Global Const gc11C4 = 128 ' &H80%
Global Const gc11C6 = 64 ' &H40%
Global Const gc11C8 = 32 ' &H20%
Global Const gc11CA = 16 ' &H10%
Global Const gc11D0 = 15 ' &HF%
Global Const gc11D2 = 31 ' &H1F%
Global Const gc11D6 = 128 ' &H80%
Global Const gc11DA = 17 ' &H11%
Global gv11DC(15) As Integer
Global gv11F2(15) As Integer
Global Const gc1208 = "std_p_e.300"
Type T2950
  M1DE4 As Integer
  M12D9 As Long
  M295E As Long
  M2969 As Integer
  M2977 As Long
  M2983 As Integer
  M2993 As Integer
  M299E As Integer
  M29AC As Integer
  M29B9 As Integer
  M29C8 As Integer
  M29D1 As Integer
  M29DB As String * 1
  M29E6 As String * 1
  M29F2 As String * 1
  M29FF As Integer
End Type

Global Const gc129A = 256 ' &H100%
Global Const gc129C = 31 ' &H1F%
Global Const gc12BA = 32 ' &H20%
Global gv12BC(32) As String
Global gv12D2(8) As String
Type T2AF1
  M1D68 As Integer
  M12D9 As Long
  M2B00 As String * 1
  M2B09 As String * 1
  M2B15 As Long
  M2B24 As Integer
  M2B30 As Integer
End Type

Type T2B3B
  M1D68 As Integer
  M2B4B As Integer
  M2B54 As Integer
  M2B61 As Integer
  M2B6E As Integer
  M12D9 As Long
End Type

Type T2B7A
  M239C As Integer
  M1DE4 As Integer
  M29FF As Integer
  M12D9 As Long
  M29AC As String
  M2B85 As String
  M29B9 As String
  M2B94 As Integer
  M2B9D As Integer
  M29DB As Integer
  M2BA8 As Integer
  M2BB3 As Integer
  M2BBD As Integer
  M29E6 As Integer
End Type

Global gv13EC() As T2B7A
Global gv141E As Integer
Type T2BD4
  M13DD As Integer
  M1D68 As String
End Type

Global gv1442() As T2BD4
Global gv1474 As Integer
Global gv1476() As Integer
Global gv14A8 As Integer
Type T2BFF
  M1D68 As String
  M2C0C As String
  M1CB7 As String
End Type

Global gv14DA() As T2BFF
Global gv150C As Integer
Global gv150E() As Integer
Global gv1540 As Integer
Type T2C2F
  M13DD As Integer
  M1D70 As Integer
  M1D68 As String
End Type

Global Const gc156A = "vbdis3i.dat"
Global Const gc156E = "x.dat"
Global Const gc1574 = 10837 ' &H2A55%
Global Const gc1576 = "%&!#@?$"
Global Const gc157A = ""
Global gv157E(1 To 7) As String
Global gv1594(1 To 8) As String
Type T2C9E
  M2CAB As Integer
  M2CB3 As Integer
  M2CBB As Integer
  M2CC3 As Integer
  M2CCB As Integer
End Type

Global gv15DA As String
Global gv15E0 As T2C9E
Global gv15EA As Integer
Type T2CE8
  M2CF5 As Integer
  M2CFB As Integer
  M2CBB As Integer
End Type

Type T2D02
  M2D10(511) As T2CE8
  M2D15(96) As Integer
End Type

Global gv1646 As T2D02
Type T2EEA
  M2EF9 As Integer
  M1C99 As Integer
  M2F00 As Integer
  M2F09 As Integer
End Type

Global gv237A As T2EEA
Global gv2382 As Integer
Type T2F3D
  M267E(10837) As Integer
End Type

Global gv23AA As T2F3D
Global gv7856 As String
Global gv7874 As Integer
Global gv7876() As String
Global gv78A8() As Integer
Global gv78DA As String
Global Const gc78E0 = 1 ' &H1%
Global Const gc78E2 = 2 ' &H2%
Global Const gc78E4 = 3 ' &H3%
Global Const gc78E6 = 4 ' &H4%
Global Const gc78E8 = 5 ' &H5%
Global Const gc78F4 = 11 ' &HB%
Global Const gc78F6 = 12 ' &HC%
Global Const gc78F8 = 13 ' &HD%
Global Const gc78FA = 14 ' &HE%
Global Const gc78FC = 15 ' &HF%
Global Const gc78FE = 16 ' &H10%
Global Const gc7900 = 17 ' &H11%
Global Const gc7902 = 18 ' &H12%
Global Const gc7904 = 19 ' &H13%
Global Const gc7906 = 20 ' &H14%
Global gv790A() As Integer
Global gv793C(21) As String
Global gv7952 As String
Global gv7956() As String
Global gv7988() As Integer
Global Const gc79BA = -1 ' &HFFFF%
Global Const gc79BC = -2 ' &HFFFE%
Global Const gc79C0 = 224 ' &HE0%
Global gv79C2 As Integer
Global Const gc79C4 = 8 ' &H8%
Global Const gc79C6 = 8 ' &H8%
Global Const gc79C8 = 9 ' &H9%
Global Const gc79CA = 10 ' &HA%
Global Const gc79CC = 11 ' &HB%
Global Const gc79CE = 16 ' &H10%
Global Const gc79D0 = 16 ' &H10%
Global Const gc79D2 = 17 ' &H11%
Global Const gc79D6 = 32 ' &H20%
Global Const gc79D8 = 64 ' &H40%
Global Const gc79DC = 96 ' &H60%
Global Const gc79DE = 128 ' &H80%
Global Const gc79E0 = 160 ' &HA0%
Global Const gc79E2 = 160 ' &HA0%
Global Const gc79E4 = 192 ' &HC0%
Global Const gc79E6 = 224 ' &HE0%
Global Const gc79EC = 3 ' &H3%
Global Const gc7A0A = 225 ' &HE1%
Global Const gc7A0C = 128 ' &H80%
Global Const gc7A0E = 226 ' &HE2%
Global Const gc7A14 = 233 ' &HE9%
Global Const gc7A16 = 231 ' &HE7%
Global Const gc7A18 = 32 ' &H20%
Global Const gc7A1A = 227 ' &HE3%
Global gv7A1C As Integer
Global Const gc7A1E = 1 ' &H1%
Global Const gc7A20 = 2 ' &H2%
Global Const gc7A22 = 4 ' &H4%
Global Const gc7A24 = 8 ' &H8%
Global Const gc7A26 = 16 ' &H10%
Global Const gc7A28 = 32 ' &H20%
Global gv7A2A As Integer
Global gv7A70() As T249F
Global gv7AA2 As Integer
Global gv7AA6 As T249F
Global gv7AB6 As Integer
Global gv7AB8 As Integer
Global gv7ABC() As T24D3
Global gv7AEE As Integer
Global gv7AF2 As T14DB
Global gv7B2A As Integer
Global gv7B2C As Integer
Global gv7B2E As Integer
Global gv7B30 As Integer
Global gv7B32 As Integer
Global gv7B34 As Integer
Global gv7B36 As Long
Global gv7B48 As Integer
Global gv7B4A As String
Global Const gc7B4E = 1 ' &H1%
Global Const gc7B50 = 2 ' &H2%
Global Const gc7B52 = 4 ' &H4%
Global Const gc7B54 = 8 ' &H8%
Global Const gc7B56 = 16 ' &H10%
Global Const gc7B58 = 32 ' &H20%
Global Const gc7B5A = 64 ' &H40%
Global Const gc7B5C = 128 ' &H80%
Global Const gc7B5E = 256 ' &H100%
Global Const gc7B60 = 512 ' &H200%
Global Const gc7B62 = 1024 ' &H400%
Global Const gc7B64 = 2048 ' &H800%
Global Const gc7B66 = 4096 ' &H1000%
Global Const gc7B68 = 8192 ' &H2000%
Global Const gc7B6A = 16384 ' &H4000%
Global Const gc7B6C = -32768 ' &H8000%
Global Const gc7B7A = 1 ' &H1%
Global Const gc7B7C = 2 ' &H2%
Global Const gc7B7E = 4 ' &H4%
Global Const gc7B80 = 8 ' &H8%
Global Const gc7B82 = 16 ' &H10%
Global gv7B84
Global gv7B88 As String
Global gv7B8C As String
Global gv7B90 As String
Global gv7B94(64) As String
Global gv7BAA As Integer


Global m_Filename As String
Global m_OutputFolder As String
Global m_SetDataPath As String
