Attribute VB_Name = "modAntiDecompiler"
'###################################################
'modAntiDecompiler
'Basic Encyrption for executables
'###################################################
'dzzie@yahoo.com
'http://sandsprite.com
'Modified by vbgamer45
'Chnage EncryptExe Function, and LoadCrypter

'Assembly Code for CrypterStub
'C7 45 F4 00 00 40 00 mov         dword ptr [ebp-0Ch],400000h
'C7 45 F0 EF BE 00 00 mov         dword ptr [ebp-10h],0BEEFh
'8B 45 F4             mov         eax,dword ptr [ebp-0Ch]
'05 AD DE 00 00       add         eax,0DEADh
'89 45 F4             mov         dword ptr [ebp-0Ch],eax
'C7 45 FC 00 00 00 00 mov         dword ptr [ebp-4],0
'EB 09                jmp         main+43h
'8B 4D FC             mov         ecx,dword ptr [ebp-4]
'83 C1 01             add         ecx,1
'89 4D FC             mov         dword ptr [ebp-4],ecx
'8B 55 FC             mov         edx,dword ptr [ebp-4]
'3B 55 F0             cmp         edx,dword ptr [ebp-10h]
'7D 22                jge         main+6Dh
'8B 45 F4             mov         eax,dword ptr [ebp-0Ch]
'03 45 FC             add         eax,dword ptr [ebp-4]
'8A 08                mov         cl,byte ptr [eax]
'88 4D F8             mov         byte ptr [ebp-8],cl
'0F BE 55 F8          movsx       edx,byte ptr [ebp-8]
'83 F2 0F             xor         edx,0Fh
'88 55 F8             mov         byte ptr [ebp-8],dl
'8B 45 F4             mov         eax,dword ptr [ebp-0Ch]
'03 45 FC             add         eax,dword ptr [ebp-4]
'8A 4D F8             mov         cl,byte ptr [ebp-8]
'88 08                mov         byte ptr [eax],cl
'EB CD                jmp         main+3Ah
'FF 65 F4             jmp         dword ptr [ebp-0Ch]


Option Explicit


Const CrypterStub = "\xC7\x45\xF4\x00\x00\x40\x00\xC7\x45\xF0\xEF\xBE\x00\x00" & _
                    "\x8B\x45\xF4\x05\xAD\xDE\x00\x00\x89\x45\xF4\xC7\x45\xFC" & _
                    "\x00\x00\x00\x00\xEB\x09\x8B\x4D\xFC\x83\xC1\x01\x89\x4D" & _
                    "\xFC\x8B\x55\xFC\x3B\x55\xF0\x7D\x22\x8B\x45\xF4\x03\x45" & _
                    "\xFC\x8A\x08\x88\x4D\xF8\x0F\xBE\x55\xF8\x83\xF2\x0F\x88" & _
                    "\x55\xF8\x8B\x45\xF4\x03\x45\xFC\x8A\x4D\xF8\x88\x08\xEB" & _
                    "\xCD\xFF\x65\xF4"

Dim Crypter() As Byte
'Pe Location Holders
Global peVirtualSizeAddr As Long
Global peAddressOfEntryPoint As Long
Global peCharacteristicsAddr As Long
Sub LoadCrypter()
'*****************************
'Purpose: To load the crypter into a byte array
'*****************************
On Error GoTo errHandle:
    Dim tmp() As String, i As Integer
    'load crypter opcodes
    tmp = Split(CrypterStub, "\x")
    ReDim Crypter(1 To UBound(tmp))
    For i = 1 To UBound(tmp)
        Crypter(i) = CByte(CInt("&h" & tmp(i)))
    Next
Exit Sub
errHandle:
    MsgBox "Error_modAntiDecompiler_LoadCrypter: " & err.Number & " " & err.Description
End Sub
Sub EncryptExe(strFileName As String, strOutput As String)
'*****************************
'Purpose: To modify the exe and add are crypting function
'*****************************
On Error GoTo errHandle:
    Dim f1 As String, f2 As String

    f1 = strFileName
    f2 = strOutput
    
    If Len(f1) = 0 Or Not FileExists(f1) Then
        MsgBox "File Does not exist please choose another file.", vbInformation
        Exit Sub
    End If
    
    If Right$(f1, 4) <> ".exe" Then
        MsgBox "Only exe files please", vbExclamation
        Exit Sub
    End If
    
    If FileExists(f2) Then Kill f2
    FileCopy f1, f2
     
    If OptHeader.ImageBase <> &H400000 Then
        MsgBox "Oops sorry this files image base does not align with the basic crypter stub change it first", vbExclamation
        Exit Sub
    End If
    
    Dim EntryPoint As Long
    Dim VirtualSize As Long
    Dim RawSize As Long
    Dim RawOffset As Long

    EntryPoint = OptHeader.EntryPoint
    VirtualSize = SecHeader(1).Address
    RawSize = SecHeader(1).SizeRawData
    RawOffset = SecHeader(1).RawDataPointer
    
    If (VirtualSize + 125) > RawSize Then 'not enough room
        MsgBox "Humm not enough room to embed decrypter sorry", vbExclamation
        Exit Sub
    End If
    
    Dim F As Long, Length As Long, b As Byte, i As Long, offset As Long
    Dim RawCrypterOffset As Long

    Length = VirtualSize - EntryPoint
  
    F = FreeFile
    Open f2 For Binary As F
    
    'crypt original opcodes
    For i = 1 To Length
        offset = EntryPoint + i
        Get F, offset, b
        b = b Xor &HF
        Put F, offset, b
    Next
           
    'advance file pointer to where we will place crypter routine
    RawCrypterOffset = RawOffset + VirtualSize
    While RawCrypterOffset Mod 16 <> 0
        RawCrypterOffset = RawCrypterOffset + 1
    Wend
    RawCrypterOffset = RawCrypterOffset + 33 'two blank lines in hexeditor after original exe code
    
    'embed base crypter routine
    Put F, RawCrypterOffset, Crypter
    
    'configure crypter stub for length and OEP (see article)
    Put F, (RawCrypterOffset + 10), Length
    Put F, (RawCrypterOffset + 18), EntryPoint
    
    
    Seek F, AppData.OptHeaderOffset + 17
    OptHeader.EntryPoint = RawCrypterOffset - 1
    'Put f, , OptHeader
    Put F, , RawCrypterOffset - 1
    'Seek f, AppData.SecHeaderOffset + 36
    Seek F, AppData.SecHeaderOffset + (41 + 16)
    Put F, , RawSize
    Seek F, AppData.SecHeaderOffset + (41 + 36)
    Put F, , &HE0000020 'read,write, execute flags

    'now do the PE file modifications
    'pe.OptionalHeader.AddressOfEntryPoint = RawCrypterOffset - 1   'file write offsets are 0 based
    Put F, modAntiDecompiler.peAddressOfEntryPoint, RawCrypterOffset - 1
    'sect.VirtualSize = sect.SizeOfRawData
   ' sect.Characteristics = &HE0000020 'read,write, execute flags
    Close F

    
    
    MsgBox "Crypter seems to be successfully implanted!", vbInformation
Exit Sub
errHandle:
    MsgBox "Error_modAntiDecompiler_EncryptExe: " & err.Number & " " & err.Description
End Sub


