Attribute VB_Name = "modAsm"

DefLng A-Z
Option Explicit
Option Base 0


Private Type ASM_OPCODE
    FullOpCode As Integer   'opcode de base (sur 8 ou 16 bits)
    OpCodeLen As Byte ' = 1 ou = 2 ....
    Flag1 As Byte
    Flag2 As Byte
    Flag3 As Byte
    Flag4 As Byte
    Flag5 As Byte
    Flag6 As Byte
    Flag7 As Byte
    Flag8 As Byte
    'description des flags  (les blancs sont en prévision pour le 64bits
    '/0         1
    '/1         2
    '/2         3
    '/3         4
    '/4         5
    '/5         6
    '/6         7
    '/7         8
    '           9...
    '/r         17
    'r/m8       18
    'r/m16      19
    'r/m32      20
    '           21
    'cb         22
    'cw         23
    'cd         24
    '           25
    'ib         26  cp
    'iw         27  cp
    'id         28  cp
    '           29
    '+rb        30
    '+rw        31
    '+rd        32
    '           33
    'rel8       34
    'rel16      35
    'rel32      36
    '           37
    'r8         38
    'r16        39
    'r32        40
    '           41
    'imm8       42
    'imm16      43
    'imm32      44
    '           45
    'ptr16:16   46
    'ptr16:32   47
    '           48
    '           49
    'm          50
    'm8         51
    'm16        52
    'm32        53
    'm64        54
    '           55
    '           56
    'm16:16     60
    'm16:32     61
    '           62
    '           63
    'm16&32     64
    'm16&16     65
    'm32&32     66
    '           67
    '           68
    '           69
    'moffs8     70
    'moffs16    71
    'moffs32    72
    '           73
    '           74
    
    'm32real    128  'fpu
    'm64real    129  'fpu
    'm80real    130  'fpu
    '           131
    'm16int     132  'fpu
    'm32int     133  'fpu
    'm64int     134  'fpu
    '           135
    'ST         159  'fpu
    'ST(0)      159  'fpu
    'ST(i)      160  'fpu
    '+i         160  'fpu
    'mm         192  'mmx
    'mm/m32     200  'mmx
    'mm/m64     201  'mmx

    sInstruct As String  'traduction string de l'opcode
    sEnd As String       's'il y a une fin string préçise
End Type
    Private TblASM_OPCODE() As ASM_OPCODE
    Private TblASM_len As Long

'table des registres, avec bit et nom
Private Type ASM_REGISTER
    r8 As String * 2
    r16 As String * 2
    r32 As String * 3
End Type
    Private TblASM_REG(0 To 7) As ASM_REGISTER

'pointe vers l'entrée asm_opcode dont le premier byte correspond
Private TblPtrASM(0 To 255) As Long

'contient le texte désassemblé, ligne par ligne
Public StrDEASM() As String
    


Sub FileDeAsm(ByVal EntryPoint As Long, ByVal Fpt As Long, ByVal CodeLen As Long, ByVal ImageRva As Long, Optional StopAtRET As Boolean = True)
'désassemble le code commençant à l'offset EntryPoint du fichier ouvert accessible via #Fpt.
'ImageRVA contient l'adresse relative du point d'entrée (nécessaire pour le calcul des JMP rel)
'CodeLen contient la distance maxi du scanner d'instruction (typiquement = LOF(Fpt))
'StopAtRET indique au scanner de s'arrêté dès qu'une instruction RET (C2h ou C3h) est trouvé (eqv End Sub)
Dim i, j, sl, ml, rvai, DataNeed
Dim Fbyte As Byte, FLong As Integer
Dim bArray(1 To 10) As Byte
Dim DumpStr As String
Dim InstructStr As String

sl = 0
i = EntryPoint
ml = i + CodeLen
rvai = ImageRva

    Do
        Get #Fpt, i, Fbyte
        Get #Fpt, i, FLong
        j = GetVASM(TblPtrASM(Fbyte), FLong)
        Get #Fpt, i, bArray()
        
        
        InstructStr = CodeToStr(bArray(), j, rvai, DataNeed)
        DumpStr = bArrayHexStr(bArray(), DataNeed)
        'crée la ligne : "rvaddress: byteshexdump [pad] asminstruction"
        sl = sl + 1
        ReDim Preserve StrDEASM(1 To sl)
        StrDEASM(sl) = Right$("0000" & Hex$(rvai), 8) & ": " & _
                       DumpStr & Space$(13 - Len(DumpStr)) & _
                       InstructStr
    
        If ((j = 385) Or (j = 386)) And StopAtRET Then
            'instruction RET scanné!
            Exit Do
        End If
    
    i = i + DataNeed
    rvai = rvai + DataNeed
    Loop Until i > ml

End Sub

Private Function GetVASM(StartPos As Long, ByVal iOpCode As Integer) As Long
'recherche le nom de l'instruction a partir du byte le plus proche (table inversé)
'renvoi un pointeur dans la table TblASM_OPCODE
Dim i
i = StartPos
    
    Do While i <= TblASM_len
        If TblASM_OPCODE(i).OpCodeLen = 1 Then
            If TblASM_OPCODE(i).FullOpCode = (iOpCode And 255) Then
                Exit Do
            End If
        Else
            If TblASM_OPCODE(i).FullOpCode = iOpCode Then
                Exit Do
            End If
        End If
    i = i + 1
    Loop
    GetVASM = i
    
End Function

Private Function CodeToStr(inCode() As Byte, inOPidx As Long, inRVA As Long, outLU As Long) As String
'texte de l'instruction désassemblé
Dim i, j, k, ol
Dim ib, iw, id
Dim dFlg, eFlg
Dim bMod As Byte, bOP As Byte, bRM As Byte, bReg As Byte
Dim sReg As String
With TblASM_OPCODE(inOPidx)

    ol = .OpCodeLen
    outLU = ol
    CodeToStr = .sInstruct

    dFlg = .Flag1 Or .Flag2 Or .Flag3 Or .Flag4
    eFlg = .Flag5 Or .Flag6 Or .Flag7 Or .Flag8
    If (eFlg + dFlg) = 0 Then
        'pas de flag = instruction direct
        CodeToStr = CodeToStr & .sEnd
        Exit Function
    ElseIf dFlg > 0 Then
        'flag uniquement post : pas de ModRM byte
    End If
    
    If .Flag1 >= 30 And .Flag1 <= 32 Then
        'le premier octet contient la valeur du registre à utiliser
        bReg = inCode(1) - .FullOpCode
        Select Case .Flag1
        Case 30
            sReg = TblASM_REG(bReg).r8
        Case 31
            sReg = TblASM_REG(bReg).r16
        Case 32
            sReg = TblASM_REG(bReg).r32
        End Select
        CodeToStr = CodeToStr & sReg
    End If
    
    
    If .Flag3 > 0 And .Flag3 < 18 Then
        outLU = outLU + 1
        'octet ModR/M utilisé
        ModRM inCode(ol + 1), bMod, bOP, bReg
        Select Case bMod
        Case 0

        Case 1
            ib = inCode(ol + 2)
            outLU = outLU + 1
        Case 2
            id = b4Long(inCode(), ol + 2)
            outLU = outLU + 4
        Case 3

        End Select
    End If

End With
End Function

Private Sub ModRM(inModRM As Byte, outMod As Byte, outReg As Byte, outRM As Byte)
'décompose l'octet ModR/M
    outMod = (inModRM And 192) / 64 '11000000b  mode adressage :
    '0=ptr [rd] : 1=ptr [rd+ib] : 2=ptr [rd+id] : 3=rb
    'ib : l'octet après ModRM contient la valeur
    'id : les 4 octets après ModRM contiennent la valeur
    
    outReg = (inModRM And 56) / 8   '00111000b  Reg/OP
    'complément pour déterminer l'opcode [0-7]
    
    outRM = inModRM And 7           '00000111b  Reg/M
    'contient le numéro du registre rb ou rd (r/m8, r/m32)
    
End Sub

Private Function bArrayHexStr(inB() As Byte, ByVal lTC As Long) As String
'converti un tableau de bytes en string hexadécimale
Dim i As Long
bArrayHexStr = Space$(lTC * 2)
i = 1

    Do While i <= lTC
        Mid$(bArrayHexStr, i, 2) = Right$("0" & Hex$(inB(i)), 2)
        i = i + 1
    Loop
    
End Function

Private Function b4Long(inB() As Byte, Ofs As Long) As Long
'renvoi une variable Long a partir de 4 valeur d'un tableau de byte
    b4Long = inB(Ofs)
    b4Long = b4Long Or CLng(inB(Ofs + 1)) * 256
    b4Long = b4Long Or CLng(inB(Ofs + 2)) * 65536
    
    If inB(Ofs + 3) < 128 Then
        b4Long = b4Long Or CLng(inB(Ofs + 3)) * 16777216
    Else  'putain de variable signé
        b4Long = (b4Long Or (CLng(inB(Ofs + 3) - 128) * 16777216)) Or &H80000000
    End If
    
End Function

Private Sub FastBin(inByte As Byte, outBin() As Byte)
'décompose un octet en valeur binaire, vers un tableau de 8 bytes
    outBin(8) = Abs(CBool(inByte And 128))
    outBin(7) = Abs(CBool(inByte And 64))
    outBin(6) = Abs(CBool(inByte And 32))
    outBin(5) = Abs(CBool(inByte And 16))
    outBin(4) = Abs(CBool(inByte And 8))
    outBin(3) = Abs(CBool(inByte And 4))
    outBin(2) = Abs(CBool(inByte And 2))
    outBin(1) = Abs(CBool(inByte And 1))
End Sub

Sub Init_unASM()
'initialisation du désassembleur
ReDim TblASM_OPCODE(1 To 648)
    
    'registres
    TblASM_REG(0).r8 = "AL": TblASM_REG(0).r16 = "AX": TblASM_REG(0).r32 = "EAX"
    TblASM_REG(1).r8 = "CL": TblASM_REG(1).r16 = "CX": TblASM_REG(1).r32 = "ECX"
    TblASM_REG(2).r8 = "DL": TblASM_REG(2).r16 = "DX": TblASM_REG(2).r32 = "EDX"
    TblASM_REG(3).r8 = "BL": TblASM_REG(3).r16 = "BX": TblASM_REG(3).r32 = "EBX"
    TblASM_REG(4).r8 = "AH": TblASM_REG(4).r16 = "SP": TblASM_REG(4).r32 = "ESP"
    TblASM_REG(5).r8 = "CH": TblASM_REG(5).r16 = "BP": TblASM_REG(5).r32 = "EBP"
    TblASM_REG(6).r8 = "DH": TblASM_REG(6).r16 = "SI": TblASM_REG(6).r32 = "ESI"
    TblASM_REG(7).r8 = "BH": TblASM_REG(7).r16 = "DI": TblASM_REG(7).r32 = "EDI"

    'ajout de toutes les instruction ASM que j'avais sous la main.
    'Meuh non je ne les ais pas taper une par une : voir converti.frm
    AddAOC TblASM_OPCODE(1), &H0, 1, 0, 0, 17, 0, 18, 38, 0, 0, "ADD ", ""
    TblPtrASM(0) = 1
    AddAOC TblASM_OPCODE(2), &H1, 1, 0, 0, 17, 0, 20, 40, 0, 0, "ADD ", ""
    TblPtrASM(1) = 2
    AddAOC TblASM_OPCODE(3), &H2, 1, 0, 0, 17, 0, 38, 18, 0, 0, "ADD ", ""
    TblPtrASM(2) = 3
    AddAOC TblASM_OPCODE(4), &H3, 1, 0, 0, 17, 0, 40, 20, 0, 0, "ADD ", ""
    TblPtrASM(3) = 4
    AddAOC TblASM_OPCODE(5), &H4, 1, 0, 0, 0, 26, 42, 0, 0, 0, "ADD AL,", ""
    TblPtrASM(4) = 5
    AddAOC TblASM_OPCODE(6), &H5, 1, 0, 0, 0, 28, 44, 0, 0, 0, "ADD EAX,", ""
    TblPtrASM(5) = 6
    AddAOC TblASM_OPCODE(7), &H6, 1, 0, 0, 0, 0, 0, 0, 0, 0, "PUSH ES", ""
    TblPtrASM(6) = 7
    AddAOC TblASM_OPCODE(8), &H7, 1, 0, 0, 0, 0, 0, 0, 0, 0, "POP ES", ""
    TblPtrASM(7) = 8
    AddAOC TblASM_OPCODE(9), &H8, 1, 0, 0, 17, 0, 18, 38, 0, 0, "OR ", ""
    TblPtrASM(8) = 9
    AddAOC TblASM_OPCODE(10), &H9, 1, 0, 0, 17, 0, 20, 40, 0, 0, "OR ", ""
    TblPtrASM(9) = 10
    AddAOC TblASM_OPCODE(11), &HA, 1, 0, 0, 17, 0, 38, 18, 0, 0, "OR ", ""
    TblPtrASM(10) = 11
    AddAOC TblASM_OPCODE(12), &HB, 1, 0, 0, 17, 0, 40, 20, 0, 0, "OR ", ""
    TblPtrASM(11) = 12
    AddAOC TblASM_OPCODE(13), &HC, 1, 0, 0, 0, 26, 42, 0, 0, 0, "OR AL,", ""
    TblPtrASM(12) = 13
    AddAOC TblASM_OPCODE(14), &HD, 1, 0, 0, 0, 28, 44, 0, 0, 0, "OR EAX,", ""
    TblPtrASM(13) = 14
    AddAOC TblASM_OPCODE(15), &HE, 1, 0, 0, 0, 0, 0, 0, 0, 0, "PUSH CS", ""
    TblPtrASM(14) = 15
    AddAOC TblASM_OPCODE(16), &HF, 2, 0, 0, 0, 0, 20, 0, 0, 0, "SLDT ", ""
    TblPtrASM(15) = 16
    AddAOC TblASM_OPCODE(17), &HF, 2, 0, 0, 2, 0, 19, 0, 0, 0, "STR ", ""
    AddAOC TblASM_OPCODE(18), &HF, 2, 0, 0, 3, 0, 19, 0, 0, 0, "LLDT ", ""
    AddAOC TblASM_OPCODE(19), &HF, 2, 0, 0, 4, 0, 19, 0, 0, 0, "LTR ", ""
    AddAOC TblASM_OPCODE(20), &HF, 2, 0, 0, 5, 0, 19, 0, 0, 0, "VERR ", ""
    AddAOC TblASM_OPCODE(21), &HF, 2, 0, 0, 6, 0, 19, 0, 0, 0, "VERW ", ""
    AddAOC TblASM_OPCODE(22), &H10F, 2, 0, 0, 0, 0, 50, 0, 0, 0, "SGDT ", ""
    AddAOC TblASM_OPCODE(23), &H10F, 2, 0, 0, 2, 0, 50, 0, 0, 0, "SIDT ", ""
    AddAOC TblASM_OPCODE(24), &H10F, 2, 0, 0, 3, 0, 64, 0, 0, 0, "LGDT ", ""
    AddAOC TblASM_OPCODE(25), &H10F, 2, 0, 0, 4, 0, 64, 0, 0, 0, "LIDT ", ""
    AddAOC TblASM_OPCODE(26), &H10F, 2, 0, 0, 5, 0, 20, 0, 0, 0, "SMSW ", ""
    AddAOC TblASM_OPCODE(27), &H10F, 2, 0, 0, 7, 0, 19, 0, 0, 0, "LMSW ", ""
    AddAOC TblASM_OPCODE(28), &H10F, 2, 0, 0, 8, 0, 50, 0, 0, 0, "INVLPG ", ""
    AddAOC TblASM_OPCODE(29), &H20F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "LAR ", ""
    AddAOC TblASM_OPCODE(30), &H30F, 2, 0, 0, 17, 0, 39, 19, 0, 0, "LSL ", ""
    AddAOC TblASM_OPCODE(31), &H30F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "LSL ", ""
    AddAOC TblASM_OPCODE(32), &H60F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "CLTS", ""
    AddAOC TblASM_OPCODE(33), &H80F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "INVD", ""
    AddAOC TblASM_OPCODE(34), &H90F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "WBINVD", ""
    AddAOC TblASM_OPCODE(35), &HB0F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "UD2", ""
    AddAOC TblASM_OPCODE(36), &H200F, 2, 0, 0, 17, 0, 40, 0, 0, 0, "MOV ", ",CR0"
    AddAOC TblASM_OPCODE(37), &H200F, 2, 0, 0, 17, 0, 40, 0, 0, 0, "MOV ", ",CR2"
    AddAOC TblASM_OPCODE(38), &H200F, 2, 0, 0, 17, 0, 40, 0, 0, 0, "MOV ", ",CR3"
    AddAOC TblASM_OPCODE(39), &H200F, 2, 0, 0, 17, 0, 40, 0, 0, 0, "MOV ", ",CR4"
    AddAOC TblASM_OPCODE(40), &H210F, 2, 0, 0, 17, 0, 40, 0, 0, 0, "MOV ", ",DR0-DR7"
    AddAOC TblASM_OPCODE(41), &H220F, 2, 0, 0, 17, 0, 40, 0, 0, 0, "MOV CR0, ", ""
    AddAOC TblASM_OPCODE(42), &H220F, 2, 0, 0, 17, 0, 40, 0, 0, 0, "MOV CR2, ", ""
    AddAOC TblASM_OPCODE(43), &H220F, 2, 0, 0, 17, 0, 40, 0, 0, 0, "MOV CR3, ", ""
    AddAOC TblASM_OPCODE(44), &H220F, 2, 0, 0, 17, 0, 40, 0, 0, 0, "MOV CR4, ", ""
    AddAOC TblASM_OPCODE(45), &H230F, 2, 0, 0, 17, 0, 40, 0, 0, 0, "MOV DR0-DR7,", ""
    AddAOC TblASM_OPCODE(46), &H300F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "WRMSR", ""
    AddAOC TblASM_OPCODE(47), &H310F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "RDTSC", ""
    AddAOC TblASM_OPCODE(48), &H320F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "RDMSR", ""
    AddAOC TblASM_OPCODE(49), &H330F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "RDPMC", ""
    AddAOC TblASM_OPCODE(50), &H400F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVO ", ""
    AddAOC TblASM_OPCODE(51), &H410F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVNO ", ""
    AddAOC TblASM_OPCODE(52), &H420F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVB ", ""
    AddAOC TblASM_OPCODE(53), &H420F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVC ", ""
    AddAOC TblASM_OPCODE(54), &H420F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVNAE ", ""
    AddAOC TblASM_OPCODE(55), &H430F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVAE ", ""
    AddAOC TblASM_OPCODE(56), &H430F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVNB ", ""
    AddAOC TblASM_OPCODE(57), &H430F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVNC ", ""
    AddAOC TblASM_OPCODE(58), &H440F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVE ", ""
    AddAOC TblASM_OPCODE(59), &H440F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVZ ", ""
    AddAOC TblASM_OPCODE(60), &H450F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVNE ", ""
    AddAOC TblASM_OPCODE(61), &H450F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVNZ ", ""
    AddAOC TblASM_OPCODE(62), &H460F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVBE ", ""
    AddAOC TblASM_OPCODE(63), &H460F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVNA ", ""
    AddAOC TblASM_OPCODE(64), &H470F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVA ", ""
    AddAOC TblASM_OPCODE(65), &H470F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVNBE ", ""
    AddAOC TblASM_OPCODE(66), &H480F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVS ", ""
    AddAOC TblASM_OPCODE(67), &H490F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVNS ", ""
    AddAOC TblASM_OPCODE(68), &H4A0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVP ", ""
    AddAOC TblASM_OPCODE(69), &H4A0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVPE ", ""
    AddAOC TblASM_OPCODE(70), &H4B0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVNP ", ""
    AddAOC TblASM_OPCODE(71), &H4B0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVPO ", ""
    AddAOC TblASM_OPCODE(72), &H4C0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVL ", ""
    AddAOC TblASM_OPCODE(73), &H4C0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVNGE ", ""
    AddAOC TblASM_OPCODE(74), &H4D0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVGE ", ""
    AddAOC TblASM_OPCODE(75), &H4D0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVNL ", ""
    AddAOC TblASM_OPCODE(76), &H4E0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVLE ", ""
    AddAOC TblASM_OPCODE(77), &H4E0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVNG ", ""
    AddAOC TblASM_OPCODE(78), &H4F0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVG ", ""
    AddAOC TblASM_OPCODE(79), &H4F0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "CMOVNLE ", ""
    AddAOC TblASM_OPCODE(80), &H600F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PUNPCKLBW ", ""
    AddAOC TblASM_OPCODE(81), &H610F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PUNPCKLWD ", ""
    AddAOC TblASM_OPCODE(82), &H620F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PUNPCKLDQ ", ""
    AddAOC TblASM_OPCODE(83), &H630F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PACKSSWB ", ""
    AddAOC TblASM_OPCODE(84), &H640F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PCMPGTB ", ""
    AddAOC TblASM_OPCODE(85), &H650F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PCMPGTW ", ""
    AddAOC TblASM_OPCODE(86), &H660F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PCMPGTD ", ""
    AddAOC TblASM_OPCODE(87), &H670F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PACKUSWB ", ""
    AddAOC TblASM_OPCODE(88), &H680F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PUNPCKHBW ", ""
    AddAOC TblASM_OPCODE(89), &H690F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PUNPCKHWD ", ""
    AddAOC TblASM_OPCODE(90), &H6A0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PUNPCKHDQ ", ""
    AddAOC TblASM_OPCODE(91), &H6B0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PACKSSDW ", ""
    AddAOC TblASM_OPCODE(92), &H6E0F, 2, 0, 0, 17, 0, 192, 20, 0, 0, "MOVD ", ""
    AddAOC TblASM_OPCODE(93), &H6F0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "MOVQ ", ""
    AddAOC TblASM_OPCODE(94), &H710F, 2, 0, 0, 3, 26, 192, 42, 0, 0, "PSRLW ", ""
    AddAOC TblASM_OPCODE(95), &H710F, 2, 0, 0, 5, 26, 192, 42, 0, 0, "PSRAW ", ""
    AddAOC TblASM_OPCODE(96), &H710F, 2, 0, 0, 7, 26, 192, 42, 0, 0, "PSLLW ", ""
    AddAOC TblASM_OPCODE(97), &H720F, 2, 0, 0, 3, 26, 192, 42, 0, 0, "PSRLD ", ""
    AddAOC TblASM_OPCODE(98), &H720F, 2, 0, 0, 5, 26, 192, 42, 0, 0, "PSRAD ", ""
    AddAOC TblASM_OPCODE(99), &H720F, 2, 0, 0, 7, 26, 192, 42, 0, 0, "PSLLD ", ""
    AddAOC TblASM_OPCODE(100), &H730F, 2, 0, 0, 3, 26, 192, 42, 0, 0, "PSRLQ ", ""
    AddAOC TblASM_OPCODE(101), &H730F, 2, 0, 0, 7, 26, 192, 42, 0, 0, "PSLLQ ", ""
    AddAOC TblASM_OPCODE(102), &H740F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PCMPEQB ", ""
    AddAOC TblASM_OPCODE(103), &H750F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PCMPEQW ", ""
    AddAOC TblASM_OPCODE(104), &H760F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PCMPEQD ", ""
    AddAOC TblASM_OPCODE(105), &H770F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "EMMS", ""
    AddAOC TblASM_OPCODE(106), &H7E0F, 2, 0, 0, 17, 0, 20, 192, 0, 0, "MOVD ", ""
    AddAOC TblASM_OPCODE(107), &H7F0F, 2, 0, 0, 17, 0, 201, 192, 0, 0, "MOVQ ", ""
    AddAOC TblASM_OPCODE(108), &H800F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JO ", ""
    AddAOC TblASM_OPCODE(109), &H810F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JNO ", ""
    AddAOC TblASM_OPCODE(110), &H820F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JB ", ""
    AddAOC TblASM_OPCODE(111), &H830F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JAE ", ""
    AddAOC TblASM_OPCODE(112), &H840F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JE ", ""
    AddAOC TblASM_OPCODE(113), &H850F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JNE ", ""
    AddAOC TblASM_OPCODE(114), &H860F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JBE ", ""
    AddAOC TblASM_OPCODE(115), &H870F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JA ", ""
    AddAOC TblASM_OPCODE(116), &H880F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JS ", ""
    AddAOC TblASM_OPCODE(117), &H890F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JNS ", ""
    AddAOC TblASM_OPCODE(118), &H8A0F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JPE ", ""
    AddAOC TblASM_OPCODE(119), &H8B0F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JPO ", ""
    AddAOC TblASM_OPCODE(120), &H8C0F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JL ", ""
    AddAOC TblASM_OPCODE(121), &H8D0F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JGE ", ""
    AddAOC TblASM_OPCODE(122), &H8E0F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JLE ", ""
    AddAOC TblASM_OPCODE(123), &H8F0F, 2, 0, 0, 0, 24, 36, 0, 0, 0, "JG ", ""
    AddAOC TblASM_OPCODE(124), &H900F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETO ", ""
    AddAOC TblASM_OPCODE(125), &H910F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETNO ", ""
    AddAOC TblASM_OPCODE(126), &H920F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETB ", ""
    AddAOC TblASM_OPCODE(127), &H930F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETAE ", ""
    AddAOC TblASM_OPCODE(128), &H940F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETE ", ""
    AddAOC TblASM_OPCODE(129), &H950F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETNE ", ""
    AddAOC TblASM_OPCODE(130), &H960F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETBE ", ""
    AddAOC TblASM_OPCODE(131), &H970F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETA ", ""
    AddAOC TblASM_OPCODE(132), &H980F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETS ", ""
    AddAOC TblASM_OPCODE(133), &H990F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETNS ", ""
    AddAOC TblASM_OPCODE(134), &H9A0F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETPE ", ""
    AddAOC TblASM_OPCODE(135), &H9B0F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETPO ", ""
    AddAOC TblASM_OPCODE(136), &H9C0F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETL ", ""
    AddAOC TblASM_OPCODE(137), &H9D0F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETGE ", ""
    AddAOC TblASM_OPCODE(138), &H9E0F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETLE ", ""
    AddAOC TblASM_OPCODE(139), &H9F0F, 2, 0, 0, 17, 0, 18, 0, 0, 0, "SETG ", ""
    AddAOC TblASM_OPCODE(140), &HA00F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "PUSH FS", ""
    AddAOC TblASM_OPCODE(141), &HA10F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "POP FS", ""
    AddAOC TblASM_OPCODE(142), &HA20F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "CPUID", ""
    AddAOC TblASM_OPCODE(143), &HA30F, 2, 0, 0, 17, 0, 20, 40, 0, 0, "BT ", ""
    AddAOC TblASM_OPCODE(144), &HA40F, 2, 0, 0, 17, 26, 20, 40, 42, 0, "SHLD ", ""
    AddAOC TblASM_OPCODE(145), &HA50F, 2, 0, 0, 17, 0, 20, 40, 0, 0, "SHLD ", ",CL"
    AddAOC TblASM_OPCODE(146), &HA80F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "PUSH GS", ""
    AddAOC TblASM_OPCODE(147), &HA90F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "POP GS", ""
    AddAOC TblASM_OPCODE(148), &HAA0F, 2, 0, 0, 0, 0, 0, 0, 0, 0, "RSM", ""
    AddAOC TblASM_OPCODE(149), &HAB0F, 2, 0, 0, 17, 0, 20, 40, 0, 0, "BTS ", ""
    AddAOC TblASM_OPCODE(150), &HAC0F, 2, 0, 0, 17, 26, 20, 40, 42, 0, "SHRD ", ""
    AddAOC TblASM_OPCODE(151), &HAD0F, 2, 0, 0, 17, 0, 20, 40, 0, 0, "SHRD ", ",CL"
    AddAOC TblASM_OPCODE(152), &HAF0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "IMUL ", ""
    AddAOC TblASM_OPCODE(153), &HB00F, 2, 0, 0, 17, 0, 18, 38, 0, 0, "CMPXCHG ", ""
    AddAOC TblASM_OPCODE(154), &HB10F, 2, 0, 0, 17, 0, 20, 40, 0, 0, "CMPXCHG ", ""
    AddAOC TblASM_OPCODE(155), &HB20F, 2, 0, 0, 17, 0, 40, 61, 0, 0, "LSS ", ""
    AddAOC TblASM_OPCODE(156), &HB30F, 2, 0, 0, 17, 0, 20, 40, 0, 0, "BTR ", ""
    AddAOC TblASM_OPCODE(157), &HB40F, 2, 0, 0, 17, 0, 40, 61, 0, 0, "LFS ", ""
    AddAOC TblASM_OPCODE(158), &HB50F, 2, 0, 0, 17, 0, 40, 61, 0, 0, "LGS ", ""
    AddAOC TblASM_OPCODE(159), &HB60F, 2, 0, 0, 17, 0, 40, 18, 0, 0, "MOVZX ", ""
    AddAOC TblASM_OPCODE(160), &HB70F, 2, 0, 0, 17, 0, 40, 19, 0, 0, "MOVZX ", ""
    AddAOC TblASM_OPCODE(161), &HBA0F, 2, 0, 0, 5, 26, 20, 42, 0, 0, "BT ", ""
    AddAOC TblASM_OPCODE(162), &HBA0F, 2, 0, 0, 6, 26, 20, 42, 0, 0, "BTS ", ""
    AddAOC TblASM_OPCODE(163), &HBA0F, 2, 0, 0, 7, 26, 20, 42, 0, 0, "BTR ", ""
    AddAOC TblASM_OPCODE(164), &HBA0F, 2, 0, 0, 8, 26, 20, 42, 0, 0, "BTC ", ""
    AddAOC TblASM_OPCODE(165), &HBB0F, 2, 0, 0, 17, 0, 20, 40, 0, 0, "BTC ", ""
    AddAOC TblASM_OPCODE(166), &HBC0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "BSF ", ""
    AddAOC TblASM_OPCODE(167), &HBD0F, 2, 0, 0, 17, 0, 40, 20, 0, 0, "BSR ", ""
    AddAOC TblASM_OPCODE(168), &HBE0F, 2, 0, 0, 17, 0, 40, 18, 0, 0, "MOVSX ", ""
    AddAOC TblASM_OPCODE(169), &HBF0F, 2, 0, 0, 17, 0, 40, 19, 0, 0, "MOVSX ", ""
    AddAOC TblASM_OPCODE(170), &HC00F, 2, 0, 0, 17, 0, 18, 38, 0, 0, "XADD ", ""
    AddAOC TblASM_OPCODE(171), &HC10F, 2, 0, 0, 17, 0, 19, 39, 0, 0, "XADD ", ""
    AddAOC TblASM_OPCODE(172), &HC10F, 2, 0, 0, 17, 0, 20, 40, 0, 0, "XADD ", ""
    AddAOC TblASM_OPCODE(173), &HC70F, 2, 0, 0, 2, 54, 54, 0, 0, 0, "CMPXCHG8B ", ""
    AddAOC TblASM_OPCODE(174), &HC80F, 2, 0, 32, 0, 0, 40, 0, 0, 0, "BSWAP ", ""
    AddAOC TblASM_OPCODE(175), &HD10F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSRLW ", ""
    AddAOC TblASM_OPCODE(176), &HD20F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSRLD ", ""
    AddAOC TblASM_OPCODE(177), &HD30F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSRLQ ", ""
    AddAOC TblASM_OPCODE(178), &HD50F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PMULLW ", ""
    AddAOC TblASM_OPCODE(179), &HD80F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSUBUSB ", ""
    AddAOC TblASM_OPCODE(180), &HD90F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSUBUSW ", ""
    AddAOC TblASM_OPCODE(181), &HDB0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PAND ", ""
    AddAOC TblASM_OPCODE(182), &HDC0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PADDUSB ", ""
    AddAOC TblASM_OPCODE(183), &HDD0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PADDUSW ", ""
    AddAOC TblASM_OPCODE(184), &HDF0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PANDN ", ""
    AddAOC TblASM_OPCODE(185), &HE10F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSRAW ", ""
    AddAOC TblASM_OPCODE(186), &HE20F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSRAD ", ""
    AddAOC TblASM_OPCODE(187), &HE50F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PMULHW ", ""
    AddAOC TblASM_OPCODE(188), &HE80F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSUBSB ", ""
    AddAOC TblASM_OPCODE(189), &HE90F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSUBSW ", ""
    AddAOC TblASM_OPCODE(190), &HEB0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "POR ", ""
    AddAOC TblASM_OPCODE(191), &HEC0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PADDSB ", ""
    AddAOC TblASM_OPCODE(192), &HED0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PADDSW ", ""
    AddAOC TblASM_OPCODE(193), &HEF0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PXOR ", ""
    AddAOC TblASM_OPCODE(194), &HF10F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSLLW ", ""
    AddAOC TblASM_OPCODE(195), &HF20F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSLLD ", ""
    AddAOC TblASM_OPCODE(196), &HF30F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSLLQ ", ""
    AddAOC TblASM_OPCODE(197), &HF50F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PMADDWD ", ""
    AddAOC TblASM_OPCODE(198), &HF80F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSUBB ", ""
    AddAOC TblASM_OPCODE(199), &HF90F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSUBW ", ""
    AddAOC TblASM_OPCODE(200), &HFA0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PSUBD ", ""
    AddAOC TblASM_OPCODE(201), &HFC0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PADDB ", ""
    AddAOC TblASM_OPCODE(202), &HFD0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PADDW ", ""
    AddAOC TblASM_OPCODE(203), &HFE0F, 2, 0, 0, 17, 0, 192, 201, 0, 0, "PADDD ", ""
    AddAOC TblASM_OPCODE(204), &H10, 1, 0, 0, 17, 0, 18, 38, 0, 0, "ADC ", ""
    TblPtrASM(16) = 204
    AddAOC TblASM_OPCODE(205), &H11, 1, 0, 0, 17, 0, 20, 40, 0, 0, "ADC ", ""
    TblPtrASM(17) = 205
    AddAOC TblASM_OPCODE(206), &H12, 1, 0, 0, 17, 0, 38, 18, 0, 0, "ADC ", ""
    TblPtrASM(18) = 206
    AddAOC TblASM_OPCODE(207), &H13, 1, 0, 0, 17, 0, 40, 20, 0, 0, "ADC ", ""
    TblPtrASM(19) = 207
    AddAOC TblASM_OPCODE(208), &H14, 1, 0, 0, 0, 26, 42, 0, 0, 0, "ADC AL,", ""
    TblPtrASM(20) = 208
    AddAOC TblASM_OPCODE(209), &H15, 1, 0, 0, 0, 28, 44, 0, 0, 0, "ADC EAX,", ""
    TblPtrASM(21) = 209
    AddAOC TblASM_OPCODE(210), &H16, 1, 0, 0, 0, 0, 0, 0, 0, 0, "PUSH SS", ""
    TblPtrASM(22) = 210
    AddAOC TblASM_OPCODE(211), &H17, 1, 0, 0, 0, 0, 0, 0, 0, 0, "POP SS", ""
    TblPtrASM(23) = 211
    AddAOC TblASM_OPCODE(212), &H18, 1, 0, 0, 17, 0, 18, 38, 0, 0, "SBB ", ""
    TblPtrASM(24) = 212
    AddAOC TblASM_OPCODE(213), &H19, 1, 0, 0, 17, 0, 20, 40, 0, 0, "SBB ", ""
    TblPtrASM(25) = 213
    AddAOC TblASM_OPCODE(214), &H1A, 1, 0, 0, 17, 0, 38, 18, 0, 0, "SBB ", ""
    TblPtrASM(26) = 214
    AddAOC TblASM_OPCODE(215), &H1B, 1, 0, 0, 17, 0, 40, 20, 0, 0, "SBB ", ""
    TblPtrASM(27) = 215
    AddAOC TblASM_OPCODE(216), &H1C, 1, 0, 0, 0, 26, 42, 0, 0, 0, "SBB AL,", ""
    TblPtrASM(28) = 216
    AddAOC TblASM_OPCODE(217), &H1D, 1, 0, 0, 0, 28, 44, 0, 0, 0, "SBB EAX,", ""
    TblPtrASM(29) = 217
    AddAOC TblASM_OPCODE(218), &H1E, 1, 0, 0, 0, 0, 0, 0, 0, 0, "PUSH DS", ""
    TblPtrASM(30) = 218
    AddAOC TblASM_OPCODE(219), &H1F, 1, 0, 0, 0, 0, 0, 0, 0, 0, "POP DS", ""
    TblPtrASM(31) = 219
    AddAOC TblASM_OPCODE(220), &H20, 1, 0, 0, 17, 0, 18, 38, 0, 0, "AND ", ""
    TblPtrASM(32) = 220
    AddAOC TblASM_OPCODE(221), &H21, 1, 0, 0, 17, 0, 20, 40, 0, 0, "AND ", ""
    TblPtrASM(33) = 221
    AddAOC TblASM_OPCODE(222), &H22, 1, 0, 0, 17, 0, 38, 18, 0, 0, "AND ", ""
    TblPtrASM(34) = 222
    AddAOC TblASM_OPCODE(223), &H23, 1, 0, 0, 17, 0, 40, 20, 0, 0, "AND ", ""
    TblPtrASM(35) = 223
    AddAOC TblASM_OPCODE(224), &H24, 1, 0, 0, 0, 26, 42, 0, 0, 0, "AND AL,", ""
    TblPtrASM(36) = 224
    AddAOC TblASM_OPCODE(225), &H25, 1, 0, 0, 0, 28, 44, 0, 0, 0, "AND EAX,", ""
    TblPtrASM(37) = 225
    AddAOC TblASM_OPCODE(226), &H26, 1, 0, 0, 0, 0, 0, 0, 0, 0, "ES:", ""
    TblPtrASM(38) = 226
    AddAOC TblASM_OPCODE(227), &H27, 1, 0, 0, 0, 0, 0, 0, 0, 0, "DAA", ""
    TblPtrASM(39) = 227
    AddAOC TblASM_OPCODE(228), &H28, 1, 0, 0, 17, 0, 18, 38, 0, 0, "SUB ", ""
    TblPtrASM(40) = 228
    AddAOC TblASM_OPCODE(229), &H29, 1, 0, 0, 17, 0, 20, 40, 0, 0, "SUB ", ""
    TblPtrASM(41) = 229
    AddAOC TblASM_OPCODE(230), &H2A, 1, 0, 0, 17, 0, 38, 18, 0, 0, "SUB ", ""
    TblPtrASM(42) = 230
    AddAOC TblASM_OPCODE(231), &H2B, 1, 0, 0, 17, 0, 40, 20, 0, 0, "SUB ", ""
    TblPtrASM(43) = 231
    AddAOC TblASM_OPCODE(232), &H2C, 1, 0, 0, 0, 26, 42, 0, 0, 0, "SUB AL,", ""
    TblPtrASM(44) = 232
    AddAOC TblASM_OPCODE(233), &H2D, 1, 0, 0, 0, 28, 44, 0, 0, 0, "SUB EAX,", ""
    TblPtrASM(45) = 233
    AddAOC TblASM_OPCODE(234), &H2E, 1, 0, 0, 0, 0, 0, 0, 0, 0, "CS:", ""
    TblPtrASM(46) = 234
    AddAOC TblASM_OPCODE(235), &H2F, 1, 0, 0, 0, 0, 0, 0, 0, 0, "DAS", ""
    TblPtrASM(47) = 235
    AddAOC TblASM_OPCODE(236), &H30, 1, 0, 0, 17, 0, 18, 38, 0, 0, "XOR ", ""
    TblPtrASM(48) = 236
    AddAOC TblASM_OPCODE(237), &H31, 1, 0, 0, 17, 0, 20, 40, 0, 0, "XOR ", ""
    TblPtrASM(49) = 237
    AddAOC TblASM_OPCODE(238), &H32, 1, 0, 0, 17, 0, 38, 18, 0, 0, "XOR ", ""
    TblPtrASM(50) = 238
    AddAOC TblASM_OPCODE(239), &H33, 1, 0, 0, 17, 0, 40, 20, 0, 0, "XOR ", ""
    TblPtrASM(51) = 239
    AddAOC TblASM_OPCODE(240), &H34, 1, 0, 0, 0, 26, 42, 0, 0, 0, "XOR AL,", ""
    TblPtrASM(52) = 240
    AddAOC TblASM_OPCODE(241), &H35, 1, 0, 0, 0, 28, 44, 0, 0, 0, "XOR EAX,", ""
    TblPtrASM(53) = 241
    AddAOC TblASM_OPCODE(242), &H36, 1, 0, 0, 0, 0, 0, 0, 0, 0, "SS:", ""
    TblPtrASM(54) = 242
    AddAOC TblASM_OPCODE(243), &H37, 1, 0, 0, 0, 0, 0, 0, 0, 0, "AAA", ""
    TblPtrASM(55) = 243
    AddAOC TblASM_OPCODE(244), &H38, 1, 0, 0, 17, 0, 18, 38, 0, 0, "CMP ", ""
    TblPtrASM(56) = 244
    AddAOC TblASM_OPCODE(245), &H39, 1, 0, 0, 17, 0, 20, 40, 0, 0, "CMP ", ""
    TblPtrASM(57) = 245
    AddAOC TblASM_OPCODE(246), &H3A, 1, 0, 0, 17, 0, 38, 18, 0, 0, "CMP ", ""
    TblPtrASM(58) = 246
    AddAOC TblASM_OPCODE(247), &H3B, 1, 0, 0, 17, 0, 40, 20, 0, 0, "CMP ", ""
    TblPtrASM(59) = 247
    AddAOC TblASM_OPCODE(248), &H3C, 1, 0, 0, 0, 26, 42, 0, 0, 0, "CMP AL,", ""
    TblPtrASM(60) = 248
    AddAOC TblASM_OPCODE(249), &H3D, 1, 0, 0, 0, 28, 44, 0, 0, 0, "CMP EAX,", ""
    TblPtrASM(61) = 249
    AddAOC TblASM_OPCODE(250), &H3E, 1, 0, 0, 0, 0, 0, 0, 0, 0, "DS:", ""
    TblPtrASM(62) = 250
    AddAOC TblASM_OPCODE(251), &H3F, 1, 0, 0, 0, 0, 0, 0, 0, 0, "AAS", ""
    TblPtrASM(63) = 251
    
    Init_UnASM_Next
    
End Sub

Private Sub Init_UnASM_Next()
    'Vous connaissiez l'erreur vb "Erreur : Procédure trop grande." ? non ?...
    '...ben fusionnez Init_UnASM() et Init_UnASM_Next() ... tatin!

    AddAOC TblASM_OPCODE(252), &H40, 1, 32, 0, 0, 0, 40, 0, 0, 0, "INC ", ""
    TblPtrASM(64) = 252
    TblPtrASM(65) = 252
    TblPtrASM(66) = 252
    TblPtrASM(67) = 252
    TblPtrASM(68) = 252
    TblPtrASM(69) = 252
    TblPtrASM(70) = 252
    TblPtrASM(71) = 252
    AddAOC TblASM_OPCODE(253), &H48, 1, 32, 0, 0, 0, 40, 0, 0, 0, "DEC ", ""
    TblPtrASM(72) = 253
    TblPtrASM(73) = 253
    TblPtrASM(74) = 253
    TblPtrASM(75) = 253
    TblPtrASM(76) = 253
    TblPtrASM(77) = 253
    TblPtrASM(78) = 253
    TblPtrASM(79) = 253
    AddAOC TblASM_OPCODE(254), &H50, 1, 32, 0, 0, 0, 40, 0, 0, 0, "PUSH ", ""
    TblPtrASM(80) = 254
    TblPtrASM(81) = 254
    TblPtrASM(82) = 254
    TblPtrASM(83) = 254
    TblPtrASM(84) = 254
    TblPtrASM(85) = 254
    TblPtrASM(86) = 254
    TblPtrASM(87) = 254
    AddAOC TblASM_OPCODE(255), &H58, 1, 32, 0, 0, 0, 40, 0, 0, 0, "POP ", ""
    TblPtrASM(88) = 255
    TblPtrASM(89) = 255
    TblPtrASM(90) = 255
    TblPtrASM(91) = 255
    TblPtrASM(92) = 255
    TblPtrASM(93) = 255
    TblPtrASM(94) = 255
    TblPtrASM(95) = 255
    AddAOC TblASM_OPCODE(256), &H60, 1, 0, 0, 0, 0, 0, 0, 0, 0, "PUSHAD", ""
    TblPtrASM(96) = 256
    AddAOC TblASM_OPCODE(257), &H61, 1, 0, 0, 0, 0, 0, 0, 0, 0, "POPAD", ""
    TblPtrASM(97) = 257
    AddAOC TblASM_OPCODE(258), &H62, 1, 0, 0, 17, 0, 40, 66, 0, 0, "BOUND ", ""
    TblPtrASM(98) = 258
    AddAOC TblASM_OPCODE(259), &H63, 1, 0, 0, 17, 0, 19, 39, 0, 0, "ARPL ", ""
    TblPtrASM(99) = 259
    AddAOC TblASM_OPCODE(260), &H64, 1, 0, 0, 0, 0, 0, 0, 0, 0, "FS:", ""
    TblPtrASM(100) = 260
    AddAOC TblASM_OPCODE(261), &H65, 1, 0, 0, 0, 0, 0, 0, 0, 0, "GS:", ""
    TblPtrASM(101) = 261
    AddAOC TblASM_OPCODE(262), &H66, 1, 0, 0, 0, 0, 0, 0, 0, 0, "Ops", ""
    TblPtrASM(102) = 262
    AddAOC TblASM_OPCODE(263), &H67, 1, 0, 0, 0, 0, 0, 0, 0, 0, "Add", ""
    TblPtrASM(103) = 263
    AddAOC TblASM_OPCODE(264), &H68, 1, 0, 0, 0, 28, 44, 0, 0, 0, "PUSH ", ""
    TblPtrASM(104) = 264
    AddAOC TblASM_OPCODE(265), &H69, 1, 0, 0, 17, 28, 40, 44, 0, 0, "IMUL ", ""
    TblPtrASM(105) = 265
    AddAOC TblASM_OPCODE(266), &H69, 1, 0, 0, 17, 28, 40, 20, 44, 0, "IMUL ", ""
    AddAOC TblASM_OPCODE(267), &H6A, 1, 0, 0, 0, 26, 42, 0, 0, 0, "PUSH ", ""
    TblPtrASM(106) = 267
    AddAOC TblASM_OPCODE(268), &H6B, 1, 0, 0, 17, 26, 40, 42, 0, 0, "IMUL ", ""
    TblPtrASM(107) = 268
    AddAOC TblASM_OPCODE(269), &H6B, 1, 0, 0, 17, 26, 40, 20, 42, 0, "IMUL ", ""
    AddAOC TblASM_OPCODE(270), &H6C, 1, 0, 0, 0, 0, 51, 0, 0, 0, "INS ", ""
    TblPtrASM(108) = 270
    AddAOC TblASM_OPCODE(271), &H6D, 1, 0, 0, 0, 0, 53, 0, 0, 0, "INS ", ""
    TblPtrASM(109) = 271
    AddAOC TblASM_OPCODE(272), &H6E, 1, 0, 0, 0, 0, 51, 0, 0, 0, "OUTS DX,", ""
    TblPtrASM(110) = 272
    AddAOC TblASM_OPCODE(273), &H6F, 1, 0, 0, 0, 0, 53, 0, 0, 0, "OUTS DX,", ""
    TblPtrASM(111) = 273
    AddAOC TblASM_OPCODE(274), &H70, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JO ", ""
    TblPtrASM(112) = 274
    AddAOC TblASM_OPCODE(275), &H71, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JNO ", ""
    TblPtrASM(113) = 275
    AddAOC TblASM_OPCODE(276), &H72, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JC ", ""
    TblPtrASM(114) = 276
    AddAOC TblASM_OPCODE(277), &H73, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JAE ", ""
    TblPtrASM(115) = 277
    AddAOC TblASM_OPCODE(278), &H74, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JE ", ""
    TblPtrASM(116) = 278
    AddAOC TblASM_OPCODE(279), &H75, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JNE ", ""
    TblPtrASM(117) = 279
    AddAOC TblASM_OPCODE(280), &H76, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JBE ", ""
    TblPtrASM(118) = 280
    AddAOC TblASM_OPCODE(281), &H77, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JA ", ""
    TblPtrASM(119) = 281
    AddAOC TblASM_OPCODE(282), &H78, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JS ", ""
    TblPtrASM(120) = 282
    AddAOC TblASM_OPCODE(283), &H79, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JNS ", ""
    TblPtrASM(121) = 283
    AddAOC TblASM_OPCODE(284), &H7A, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JPE ", ""
    TblPtrASM(122) = 284
    AddAOC TblASM_OPCODE(285), &H7B, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JPO ", ""
    TblPtrASM(123) = 285
    AddAOC TblASM_OPCODE(286), &H7C, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JL ", ""
    TblPtrASM(124) = 286
    AddAOC TblASM_OPCODE(287), &H7D, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JGE ", ""
    TblPtrASM(125) = 287
    AddAOC TblASM_OPCODE(288), &H7E, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JLE ", ""
    TblPtrASM(126) = 288
    AddAOC TblASM_OPCODE(289), &H7F, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JG ", ""
    TblPtrASM(127) = 289
    AddAOC TblASM_OPCODE(290), &H80, 1, 0, 0, 0, 26, 18, 42, 0, 0, "ADD ", ""
    TblPtrASM(128) = 290
    AddAOC TblASM_OPCODE(291), &H80, 1, 0, 0, 2, 26, 18, 42, 0, 0, "OR ", ""
    AddAOC TblASM_OPCODE(292), &H80, 1, 0, 0, 3, 26, 18, 42, 0, 0, "ADC ", ""
    AddAOC TblASM_OPCODE(293), &H80, 1, 0, 0, 4, 26, 18, 42, 0, 0, "SBB ", ""
    AddAOC TblASM_OPCODE(294), &H80, 1, 0, 0, 5, 26, 18, 42, 0, 0, "AND ", ""
    AddAOC TblASM_OPCODE(295), &H80, 1, 0, 0, 6, 26, 18, 42, 0, 0, "SUB ", ""
    AddAOC TblASM_OPCODE(296), &H80, 1, 0, 0, 7, 26, 18, 42, 0, 0, "XOR ", ""
    AddAOC TblASM_OPCODE(297), &H80, 1, 0, 0, 8, 26, 18, 42, 0, 0, "CMP ", ""
    AddAOC TblASM_OPCODE(298), &H81, 1, 0, 0, 0, 28, 20, 44, 0, 0, "ADD ", ""
    TblPtrASM(129) = 298
    AddAOC TblASM_OPCODE(299), &H81, 1, 0, 0, 2, 28, 20, 44, 0, 0, "OR ", ""
    AddAOC TblASM_OPCODE(300), &H81, 1, 0, 0, 3, 28, 20, 44, 0, 0, "ADC ", ""
    AddAOC TblASM_OPCODE(301), &H81, 1, 0, 0, 4, 28, 20, 44, 0, 0, "SBB ", ""
    AddAOC TblASM_OPCODE(302), &H81, 1, 0, 0, 5, 28, 20, 44, 0, 0, "AND ", ""
    AddAOC TblASM_OPCODE(303), &H81, 1, 0, 0, 6, 28, 20, 44, 0, 0, "SUB ", ""
    AddAOC TblASM_OPCODE(304), &H81, 1, 0, 0, 7, 28, 20, 44, 0, 0, "XOR ", ""
    AddAOC TblASM_OPCODE(305), &H81, 1, 0, 0, 8, 28, 20, 44, 0, 0, "CMP ", ""
    AddAOC TblASM_OPCODE(306), &H83, 1, 0, 0, 0, 26, 20, 42, 0, 0, "ADD ", ""
    TblPtrASM(131) = 306
    AddAOC TblASM_OPCODE(307), &H83, 1, 0, 0, 2, 26, 20, 42, 0, 0, "OR ", ""
    AddAOC TblASM_OPCODE(308), &H83, 1, 0, 0, 3, 26, 20, 42, 0, 0, "ADC ", ""
    AddAOC TblASM_OPCODE(309), &H83, 1, 0, 0, 4, 26, 20, 42, 0, 0, "SBB ", ""
    AddAOC TblASM_OPCODE(310), &H83, 1, 0, 0, 5, 26, 20, 42, 0, 0, "AND ", ""
    AddAOC TblASM_OPCODE(311), &H83, 1, 0, 0, 6, 26, 20, 42, 0, 0, "SUB ", ""
    AddAOC TblASM_OPCODE(312), &H83, 1, 0, 0, 7, 26, 20, 42, 0, 0, "XOR ", ""
    AddAOC TblASM_OPCODE(313), &H83, 1, 0, 0, 8, 26, 20, 42, 0, 0, "CMP ", ""
    AddAOC TblASM_OPCODE(314), &H84, 1, 0, 0, 17, 0, 18, 38, 0, 0, "TEST ", ""
    TblPtrASM(132) = 314
    AddAOC TblASM_OPCODE(315), &H85, 1, 0, 0, 17, 0, 19, 39, 0, 0, "TEST ", ""
    TblPtrASM(133) = 315
    AddAOC TblASM_OPCODE(316), &H85, 1, 0, 0, 17, 0, 20, 40, 0, 0, "TEST ", ""
    AddAOC TblASM_OPCODE(317), &H86, 1, 0, 0, 17, 0, 18, 38, 0, 0, "XCHG ", ""
    TblPtrASM(134) = 317
    AddAOC TblASM_OPCODE(318), &H86, 1, 0, 0, 17, 0, 38, 18, 0, 0, "XCHG ", ""
    AddAOC TblASM_OPCODE(319), &H87, 1, 0, 0, 17, 0, 20, 40, 0, 0, "XCHG ", ""
    TblPtrASM(135) = 319
    AddAOC TblASM_OPCODE(320), &H87, 1, 0, 0, 17, 0, 40, 20, 0, 0, "XCHG ", ""
    AddAOC TblASM_OPCODE(321), &H88, 1, 0, 0, 17, 0, 18, 38, 0, 0, "MOV ", ""
    TblPtrASM(136) = 321
    AddAOC TblASM_OPCODE(322), &H89, 1, 0, 0, 17, 0, 20, 40, 0, 0, "MOV ", ""
    TblPtrASM(137) = 322
    AddAOC TblASM_OPCODE(323), &H8A, 1, 0, 0, 17, 0, 38, 18, 0, 0, "MOV ", ""
    TblPtrASM(138) = 323
    AddAOC TblASM_OPCODE(324), &H8B, 1, 0, 0, 17, 0, 40, 20, 0, 0, "MOV ", ""
    TblPtrASM(139) = 324
    AddAOC TblASM_OPCODE(325), &H8C, 1, 0, 0, 17, 0, 19, 0, 0, 0, "MOV ", ""
    TblPtrASM(140) = 325
    AddAOC TblASM_OPCODE(326), &H8D, 1, 0, 0, 17, 0, 40, 50, 0, 0, "LEA ", ""
    TblPtrASM(141) = 326
    AddAOC TblASM_OPCODE(327), &H8E, 1, 0, 0, 17, 0, 0, 0, 0, 0, "MOV S", ""
    TblPtrASM(142) = 327
    AddAOC TblASM_OPCODE(328), &H8F, 1, 0, 0, 0, 0, 53, 0, 0, 0, "POP ", ""
    TblPtrASM(143) = 328
    AddAOC TblASM_OPCODE(329), &H90, 1, 0, 0, 0, 0, 0, 0, 0, 0, "NOP", ""
    TblPtrASM(144) = 329
    
    'equivoque ???
    AddAOC TblASM_OPCODE(330), &H90, 1, 32, 0, 0, 0, 40, 0, 0, 0, "XCHG EAX,", ""
    TblPtrASM(145) = 330
    TblPtrASM(146) = 330
    TblPtrASM(147) = 330
    TblPtrASM(148) = 330
    TblPtrASM(149) = 330
    TblPtrASM(150) = 330
    TblPtrASM(151) = 330
    AddAOC TblASM_OPCODE(331), &H90, 1, 32, 0, 0, 0, 40, 0, 0, 0, "XCHG ", ",EAX"
    'ee
    
    AddAOC TblASM_OPCODE(332), &H98, 1, 0, 0, 0, 0, 0, 0, 0, 0, "CBW", ""
    TblPtrASM(152) = 332
    AddAOC TblASM_OPCODE(333), &H99, 1, 0, 0, 0, 0, 0, 0, 0, 0, "CDQ", ""
    TblPtrASM(153) = 333
    AddAOC TblASM_OPCODE(334), &H99, 1, 0, 0, 0, 0, 0, 0, 0, 0, "CWD", ""
    AddAOC TblASM_OPCODE(335), &H9A, 1, 0, 0, 0, 0, 0, 0, 0, 0, "CALL pt", ""
    TblPtrASM(154) = 335
    AddAOC TblASM_OPCODE(336), &H9B, 1, 0, 0, 0, 0, 0, 0, 0, 0, "FWAIT", ""
    TblPtrASM(155) = 336
    AddAOC TblASM_OPCODE(337), &H9B, 1, 0, 0, 0, 0, 0, 0, 0, 0, "WAIT", ""
    AddAOC TblASM_OPCODE(338), &HD99B, 2, 0, 0, 7, 0, 0, 0, 0, 0, "FSTENV ", ""
    AddAOC TblASM_OPCODE(339), &HD99B, 2, 0, 0, 8, 0, 0, 0, 0, 0, "FSTCW ", ""
    AddAOC TblASM_OPCODE(340), &HDB9B, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FCLEX", ""
    AddAOC TblASM_OPCODE(341), &HDB9B, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FINIT", ""
    AddAOC TblASM_OPCODE(342), &HDD9B, 2, 0, 0, 7, 0, 0, 0, 0, 0, "FSAVE ", ""
    AddAOC TblASM_OPCODE(343), &HDD9B, 2, 0, 0, 8, 0, 0, 0, 0, 0, "FSTSW ", ""
    AddAOC TblASM_OPCODE(344), &HDF9B, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FSTSW AX", ""
    AddAOC TblASM_OPCODE(345), &H9C, 1, 0, 0, 0, 0, 0, 0, 0, 0, "PUSHFD", ""
    TblPtrASM(156) = 345
    AddAOC TblASM_OPCODE(346), &H9D, 1, 0, 0, 0, 0, 0, 0, 0, 0, "POPFD", ""
    TblPtrASM(157) = 346
    AddAOC TblASM_OPCODE(347), &H9E, 1, 0, 0, 0, 0, 0, 0, 0, 0, "SAHF", ""
    TblPtrASM(158) = 347
    AddAOC TblASM_OPCODE(348), &H9F, 1, 0, 0, 0, 0, 0, 0, 0, 0, "LAHF", ""
    TblPtrASM(159) = 348
    AddAOC TblASM_OPCODE(349), &HA0, 1, 0, 0, 0, 0, 0, 0, 0, 0, "MOV AL, ", ""
    TblPtrASM(160) = 349
    AddAOC TblASM_OPCODE(350), &HA1, 1, 0, 0, 0, 0, 0, 0, 0, 0, "MOV AX, ", ""
    TblPtrASM(161) = 350
    AddAOC TblASM_OPCODE(351), &HA1, 1, 0, 0, 0, 0, 0, 0, 0, 0, "MOV EAX, ", ""
    AddAOC TblASM_OPCODE(352), &HA2, 1, 0, 0, 0, 0, 0, 0, 0, 0, "MOV ", ",AL"
    TblPtrASM(162) = 352
    AddAOC TblASM_OPCODE(353), &HA3, 1, 0, 0, 0, 0, 0, 0, 0, 0, "MOV ", ",AX"
    TblPtrASM(163) = 353
    AddAOC TblASM_OPCODE(354), &HA3, 1, 0, 0, 0, 0, 0, 0, 0, 0, "MOV ", ",EAX"
    AddAOC TblASM_OPCODE(355), &HA4, 1, 0, 0, 0, 0, 51, 51, 0, 0, "MOVS ", ""
    TblPtrASM(164) = 355
    AddAOC TblASM_OPCODE(356), &HA5, 1, 0, 0, 0, 0, 53, 53, 0, 0, "MOVS ", ""
    TblPtrASM(165) = 356
    AddAOC TblASM_OPCODE(357), &HA6, 1, 0, 0, 0, 0, 0, 0, 0, 0, "CMPSB", ""
    TblPtrASM(166) = 357
    AddAOC TblASM_OPCODE(358), &HA7, 1, 0, 0, 0, 0, 0, 0, 0, 0, "CMPSD", ""
    TblPtrASM(167) = 358
    AddAOC TblASM_OPCODE(359), &HA8, 1, 0, 0, 0, 26, 42, 0, 0, 0, "TEST AL,", ""
    TblPtrASM(168) = 359
    AddAOC TblASM_OPCODE(360), &HA9, 1, 0, 0, 0, 28, 44, 0, 0, 0, "TEST EAX,", ""
    TblPtrASM(169) = 360
    AddAOC TblASM_OPCODE(361), &HAA, 1, 0, 0, 0, 0, 51, 0, 0, 0, "STOS ", ""
    TblPtrASM(170) = 361
    AddAOC TblASM_OPCODE(362), &HAB, 1, 0, 0, 0, 0, 53, 0, 0, 0, "STOS ", ""
    TblPtrASM(171) = 362
    AddAOC TblASM_OPCODE(363), &HAC, 1, 0, 0, 0, 0, 51, 0, 0, 0, "LODS ", ""
    TblPtrASM(172) = 363
    AddAOC TblASM_OPCODE(364), &HAD, 1, 0, 0, 0, 0, 53, 0, 0, 0, "LODS ", ""
    TblPtrASM(173) = 364
    AddAOC TblASM_OPCODE(365), &HAE, 1, 0, 0, 0, 0, 51, 0, 0, 0, "SCAS ", ""
    TblPtrASM(174) = 365
    AddAOC TblASM_OPCODE(366), &HAF, 1, 0, 0, 0, 0, 53, 0, 0, 0, "SCAS ", ""
    TblPtrASM(175) = 366
    AddAOC TblASM_OPCODE(367), &HB0, 1, 30, 0, 0, 0, 38, 42, 0, 0, "MOV ", ""
    TblPtrASM(176) = 367
    TblPtrASM(177) = 367
    TblPtrASM(178) = 367
    TblPtrASM(179) = 367
    TblPtrASM(180) = 367
    TblPtrASM(181) = 367
    TblPtrASM(182) = 367
    TblPtrASM(183) = 367
    AddAOC TblASM_OPCODE(368), &HB8, 1, 32, 0, 0, 0, 40, 44, 0, 0, "MOV ", ""
    TblPtrASM(184) = 368
    TblPtrASM(185) = 368
    TblPtrASM(186) = 368
    TblPtrASM(187) = 368
    TblPtrASM(188) = 368
    TblPtrASM(189) = 368
    TblPtrASM(190) = 368
    TblPtrASM(191) = 368
    AddAOC TblASM_OPCODE(369), &HC0, 1, 0, 0, 0, 26, 18, 42, 0, 0, "ROL ", ""
    TblPtrASM(192) = 369
    AddAOC TblASM_OPCODE(370), &HC0, 1, 0, 0, 2, 26, 18, 42, 0, 0, "ROR ", ""
    AddAOC TblASM_OPCODE(371), &HC0, 1, 0, 0, 3, 26, 18, 42, 0, 0, "RCL ", ""
    AddAOC TblASM_OPCODE(372), &HC0, 1, 0, 0, 4, 26, 18, 42, 0, 0, "RCR ", ""
    AddAOC TblASM_OPCODE(373), &HC0, 1, 0, 0, 5, 26, 18, 42, 0, 0, "SAL ", ""
    AddAOC TblASM_OPCODE(374), &HC0, 1, 0, 0, 5, 26, 18, 42, 0, 0, "SHL ", ""
    AddAOC TblASM_OPCODE(375), &HC0, 1, 0, 0, 6, 26, 18, 42, 0, 0, "SHR ", ""
    AddAOC TblASM_OPCODE(376), &HC0, 1, 0, 0, 8, 26, 18, 42, 0, 0, "SAR ", ""
    AddAOC TblASM_OPCODE(377), &HC1, 1, 0, 0, 0, 26, 20, 42, 0, 0, "ROL ", ""
    TblPtrASM(193) = 377
    AddAOC TblASM_OPCODE(378), &HC1, 1, 0, 0, 2, 26, 20, 42, 0, 0, "ROR ", ""
    AddAOC TblASM_OPCODE(379), &HC1, 1, 0, 0, 3, 26, 20, 42, 0, 0, "RCL ", ""
    AddAOC TblASM_OPCODE(380), &HC1, 1, 0, 0, 4, 26, 20, 42, 0, 0, "RCR ", ""
    AddAOC TblASM_OPCODE(381), &HC1, 1, 0, 0, 5, 26, 20, 42, 0, 0, "SAL ", ""
    AddAOC TblASM_OPCODE(382), &HC1, 1, 0, 0, 5, 26, 20, 42, 0, 0, "SHL ", ""
    AddAOC TblASM_OPCODE(383), &HC1, 1, 0, 0, 6, 26, 20, 42, 0, 0, "SHR ", ""
    AddAOC TblASM_OPCODE(384), &HC1, 1, 0, 0, 8, 26, 20, 42, 0, 0, "SAR ", ""
    AddAOC TblASM_OPCODE(385), &HC2, 1, 0, 0, 0, 27, 43, 0, 0, 0, "RET ", ""
    TblPtrASM(194) = 385
    AddAOC TblASM_OPCODE(386), &HC3, 1, 0, 0, 0, 0, 0, 0, 0, 0, "RET", ""
    TblPtrASM(195) = 386
    AddAOC TblASM_OPCODE(387), &HC4, 1, 0, 0, 17, 0, 40, 61, 0, 0, "LES ", ""
    TblPtrASM(196) = 387
    AddAOC TblASM_OPCODE(388), &HC5, 1, 0, 0, 17, 0, 40, 61, 0, 0, "LDS ", ""
    TblPtrASM(197) = 388
    AddAOC TblASM_OPCODE(389), &HC6, 1, 0, 0, 0, 26, 18, 42, 0, 0, "MOV ", ""
    TblPtrASM(198) = 389
    AddAOC TblASM_OPCODE(390), &HC7, 1, 0, 0, 0, 28, 20, 44, 0, 0, "MOV ", ""
    TblPtrASM(199) = 390
    AddAOC TblASM_OPCODE(391), &HC8, 1, 0, 0, 27, 0, 43, 0, 0, 0, "ENTER ", ",0"
    TblPtrASM(200) = 391
    AddAOC TblASM_OPCODE(392), &HC8, 1, 0, 0, 27, 0, 43, 0, 0, 0, "ENTER ", ",1"
    AddAOC TblASM_OPCODE(393), &HC8, 1, 0, 0, 27, 26, 43, 42, 0, 0, "ENTER ", ""
    AddAOC TblASM_OPCODE(394), &HC9, 1, 0, 0, 0, 0, 0, 0, 0, 0, "LEAVE", ""
    TblPtrASM(201) = 394
    AddAOC TblASM_OPCODE(395), &HCA, 1, 0, 0, 0, 27, 43, 0, 0, 0, "RET ", ""
    TblPtrASM(202) = 395
    AddAOC TblASM_OPCODE(396), &HCB, 1, 0, 0, 0, 0, 0, 0, 0, 0, "RET", ""
    TblPtrASM(203) = 396
    AddAOC TblASM_OPCODE(397), &HCC, 1, 0, 0, 0, 0, 0, 0, 0, 0, "INT 3", ""
    TblPtrASM(204) = 397
    AddAOC TblASM_OPCODE(398), &HCD, 1, 0, 0, 0, 26, 42, 0, 0, 0, "INT ", ""
    TblPtrASM(205) = 398
    AddAOC TblASM_OPCODE(399), &HCE, 1, 0, 0, 0, 0, 0, 0, 0, 0, "INTO", ""
    TblPtrASM(206) = 399
    AddAOC TblASM_OPCODE(400), &HCF, 1, 0, 0, 0, 0, 0, 0, 0, 0, "IRETD", ""
    TblPtrASM(207) = 400
    AddAOC TblASM_OPCODE(401), &HD0, 1, 0, 0, 0, 0, 18, 0, 0, 0, "ROL ", ",1"
    TblPtrASM(208) = 401
    AddAOC TblASM_OPCODE(402), &HD0, 1, 0, 0, 2, 0, 18, 0, 0, 0, "ROR ", ",1"
    AddAOC TblASM_OPCODE(403), &HD0, 1, 0, 0, 3, 0, 18, 0, 0, 0, "RCL ", ",1"
    AddAOC TblASM_OPCODE(404), &HD0, 1, 0, 0, 4, 0, 18, 0, 0, 0, "RCR ", ",1"
    AddAOC TblASM_OPCODE(405), &HD0, 1, 0, 0, 5, 0, 18, 0, 0, 0, "SAL ", ",1"
    AddAOC TblASM_OPCODE(406), &HD0, 1, 0, 0, 5, 0, 18, 0, 0, 0, "SHL ", ",1"
    AddAOC TblASM_OPCODE(407), &HD0, 1, 0, 0, 6, 0, 18, 0, 0, 0, "SHR ", ",1"
    AddAOC TblASM_OPCODE(408), &HD0, 1, 0, 0, 8, 0, 18, 0, 0, 0, "SAR ", ",1"
    AddAOC TblASM_OPCODE(409), &HD1, 1, 0, 0, 0, 0, 20, 0, 0, 0, "ROL ", ",1"
    TblPtrASM(209) = 409
    AddAOC TblASM_OPCODE(410), &HD1, 1, 0, 0, 2, 0, 20, 0, 0, 0, "ROR ", ",1"
    AddAOC TblASM_OPCODE(411), &HD1, 1, 0, 0, 3, 0, 20, 0, 0, 0, "RCL ", ",1"
    AddAOC TblASM_OPCODE(412), &HD1, 1, 0, 0, 4, 0, 20, 0, 0, 0, "RCR ", ",1"
    AddAOC TblASM_OPCODE(413), &HD1, 1, 0, 0, 5, 0, 20, 0, 0, 0, "SAL ", ",1"
    AddAOC TblASM_OPCODE(414), &HD1, 1, 0, 0, 5, 0, 20, 0, 0, 0, "SHL ", ",1"
    AddAOC TblASM_OPCODE(415), &HD1, 1, 0, 0, 6, 0, 20, 0, 0, 0, "SHR ", ",1"
    AddAOC TblASM_OPCODE(416), &HD1, 1, 0, 0, 8, 0, 20, 0, 0, 0, "SAR ", ",1"
    AddAOC TblASM_OPCODE(417), &HD2, 1, 0, 0, 0, 0, 18, 0, 0, 0, "ROL ", ",CL"
    TblPtrASM(210) = 417
    AddAOC TblASM_OPCODE(418), &HD2, 1, 0, 0, 2, 0, 18, 0, 0, 0, "ROR ", ",CL"
    AddAOC TblASM_OPCODE(419), &HD2, 1, 0, 0, 3, 0, 18, 0, 0, 0, "RCL ", ",CL"
    AddAOC TblASM_OPCODE(420), &HD2, 1, 0, 0, 4, 0, 18, 0, 0, 0, "RCR ", ",CL"
    AddAOC TblASM_OPCODE(421), &HD2, 1, 0, 0, 5, 0, 18, 0, 0, 0, "SAL ", ",CL"
    AddAOC TblASM_OPCODE(422), &HD2, 1, 0, 0, 5, 0, 18, 0, 0, 0, "SHL ", ",CL"
    AddAOC TblASM_OPCODE(423), &HD2, 1, 0, 0, 6, 0, 18, 0, 0, 0, "SHR ", ",CL"
    AddAOC TblASM_OPCODE(424), &HD2, 1, 0, 0, 8, 0, 18, 0, 0, 0, "SAR ", ",CL"
    AddAOC TblASM_OPCODE(425), &HD3, 1, 0, 0, 0, 0, 20, 0, 0, 0, "ROL ", ",CL"
    TblPtrASM(211) = 425
    AddAOC TblASM_OPCODE(426), &HD3, 1, 0, 0, 2, 0, 20, 0, 0, 0, "ROR ", ",CL"
    AddAOC TblASM_OPCODE(427), &HD3, 1, 0, 0, 3, 0, 20, 0, 0, 0, "RCL ", ",CL"
    AddAOC TblASM_OPCODE(428), &HD3, 1, 0, 0, 4, 0, 20, 0, 0, 0, "RCR ", ",CL"
    AddAOC TblASM_OPCODE(429), &HD3, 1, 0, 0, 5, 0, 20, 0, 0, 0, "SAL ", ",CL"
    AddAOC TblASM_OPCODE(430), &HD3, 1, 0, 0, 5, 0, 20, 0, 0, 0, "SHL ", ",CL"
    AddAOC TblASM_OPCODE(431), &HD3, 1, 0, 0, 6, 0, 20, 0, 0, 0, "SHR ", ",CL"
    AddAOC TblASM_OPCODE(432), &HD3, 1, 0, 0, 8, 0, 20, 0, 0, 0, "SAR ", ",CL"
    AddAOC TblASM_OPCODE(433), &HAD4, 2, 0, 0, 0, 0, 0, 0, 0, 0, "AAM", ""
    TblPtrASM(212) = 433
    AddAOC TblASM_OPCODE(434), &HAD5, 2, 0, 0, 0, 0, 0, 0, 0, 0, "AAD", ""
    TblPtrASM(213) = 434
    AddAOC TblASM_OPCODE(435), &HD6, 1, 0, 0, 0, 0, 0, 0, 0, 0, "SETALC", ""
    TblPtrASM(214) = 435
    AddAOC TblASM_OPCODE(436), &HD7, 1, 0, 0, 0, 0, 51, 0, 0, 0, "XLAT ", ""
    TblPtrASM(215) = 436
    AddAOC TblASM_OPCODE(437), &HD8, 1, 0, 0, 0, 0, 128, 0, 0, 0, "FADD ", ""
    TblPtrASM(216) = 437
    AddAOC TblASM_OPCODE(438), &HD8, 1, 0, 0, 2, 0, 128, 0, 0, 0, "FMUL ", ""
    AddAOC TblASM_OPCODE(439), &HD8, 1, 0, 0, 3, 0, 128, 0, 0, 0, "FCOM ", ""
    AddAOC TblASM_OPCODE(440), &HD8, 1, 0, 0, 4, 0, 128, 0, 0, 0, "FCOMP ", ""
    AddAOC TblASM_OPCODE(441), &HD8, 1, 0, 0, 5, 0, 128, 0, 0, 0, "FSUB ", ""
    AddAOC TblASM_OPCODE(442), &HD8, 1, 0, 0, 6, 0, 128, 0, 0, 0, "FSUBR ", ""
    AddAOC TblASM_OPCODE(443), &HD8, 1, 0, 0, 7, 0, 128, 0, 0, 0, "FDIV ", ""
    AddAOC TblASM_OPCODE(444), &HD8, 1, 0, 0, 8, 0, 128, 0, 0, 0, "FDIVR ", ""
    AddAOC TblASM_OPCODE(445), &HC0D8, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FADD ST(0),ST(", ")"
    AddAOC TblASM_OPCODE(446), &HC8D8, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FMUL ST(0),ST(", ")"
    AddAOC TblASM_OPCODE(447), &HD0D8, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FCOM ST(", ""
    AddAOC TblASM_OPCODE(448), &HD1D8, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FCOM", ""
    AddAOC TblASM_OPCODE(449), &HD8D8, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FCOMP ST(", ")"
    AddAOC TblASM_OPCODE(450), &HD9D8, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FCOMP", ""
    AddAOC TblASM_OPCODE(451), &HE0D8, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FSUB ST(0),ST(", ")"
    AddAOC TblASM_OPCODE(452), &HE8D8, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FSUBR ST(0),ST(", ")"
    AddAOC TblASM_OPCODE(453), &HF0D8, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FDIV ST(0),ST(", ")"
    AddAOC TblASM_OPCODE(454), &HF8D8, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FDIVR ST(0),ST(", ")"
    AddAOC TblASM_OPCODE(455), &HD9, 1, 0, 0, 0, 0, 128, 0, 0, 0, "FLD ", ""
    TblPtrASM(217) = 455
    AddAOC TblASM_OPCODE(456), &HD9, 1, 0, 0, 3, 0, 128, 0, 0, 0, "FST ", ""
    AddAOC TblASM_OPCODE(457), &HD9, 1, 0, 0, 4, 0, 128, 0, 0, 0, "FSTP ", ""
    AddAOC TblASM_OPCODE(458), &HD9, 1, 0, 0, 5, 0, 0, 0, 0, 0, "FLDENV ", ""
    AddAOC TblASM_OPCODE(459), &HD9, 1, 0, 0, 6, 0, 0, 0, 0, 0, "FLDCW ", ""
    AddAOC TblASM_OPCODE(460), &HD9, 1, 0, 0, 7, 0, 0, 0, 0, 0, "FNSTENV ", ""
    AddAOC TblASM_OPCODE(461), &HD9, 1, 0, 0, 8, 0, 0, 0, 0, 0, "FNSTCW ", ""
    AddAOC TblASM_OPCODE(462), &HC0D9, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FLD ST(", ")"
    AddAOC TblASM_OPCODE(463), &HC8D9, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FXCH ST(", ")"
    AddAOC TblASM_OPCODE(464), &HC9D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FXCH", ""
    AddAOC TblASM_OPCODE(465), &HD0D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FNOP", ""
    AddAOC TblASM_OPCODE(466), &HE0D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FCHS", ""
    AddAOC TblASM_OPCODE(467), &HE1D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FABS", ""
    AddAOC TblASM_OPCODE(468), &HE4D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FTST", ""
    AddAOC TblASM_OPCODE(469), &HE5D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FXAM", ""
    AddAOC TblASM_OPCODE(470), &HE8D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FLD1", ""
    AddAOC TblASM_OPCODE(471), &HE9D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FLDL2T", ""
    AddAOC TblASM_OPCODE(472), &HEAD9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FLDL2E", ""
    AddAOC TblASM_OPCODE(473), &HEBD9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FLDPI", ""
    AddAOC TblASM_OPCODE(474), &HECD9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FLDLG2", ""
    AddAOC TblASM_OPCODE(475), &HEDD9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FLDLN2", ""
    AddAOC TblASM_OPCODE(476), &HEED9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FLDZ", ""
    AddAOC TblASM_OPCODE(477), &HF0D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "F2XM1", ""
    AddAOC TblASM_OPCODE(478), &HF1D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FYL2X", ""
    AddAOC TblASM_OPCODE(479), &HF2D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FPTAN", ""
    AddAOC TblASM_OPCODE(480), &HF3D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FPATAN", ""
    AddAOC TblASM_OPCODE(481), &HF4D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FXTRACT", ""
    AddAOC TblASM_OPCODE(482), &HF5D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FPREM1", ""
    AddAOC TblASM_OPCODE(483), &HF6D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FDECSTP", ""
    AddAOC TblASM_OPCODE(484), &HF7D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FINCSTP", ""
    AddAOC TblASM_OPCODE(485), &HF8D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FPREM", ""
    AddAOC TblASM_OPCODE(486), &HF9D9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FYL2XP1", ""
    AddAOC TblASM_OPCODE(487), &HFAD9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FSQRT", ""
    AddAOC TblASM_OPCODE(488), &HFBD9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FSINCOS", ""
    AddAOC TblASM_OPCODE(489), &HFCD9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FRNDINT", ""
    AddAOC TblASM_OPCODE(490), &HFDD9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FSCALE", ""
    AddAOC TblASM_OPCODE(491), &HFED9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FSIN", ""
    AddAOC TblASM_OPCODE(492), &HFFD9, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FCOS", ""
    AddAOC TblASM_OPCODE(493), &HDA, 1, 0, 0, 0, 0, 133, 0, 0, 0, "FIADD ", ""
    TblPtrASM(218) = 493
    AddAOC TblASM_OPCODE(494), &HDA, 1, 0, 0, 2, 0, 133, 0, 0, 0, "FIMUL ", ""
    AddAOC TblASM_OPCODE(495), &HDA, 1, 0, 0, 3, 0, 133, 0, 0, 0, "FICOM ", ""
    AddAOC TblASM_OPCODE(496), &HDA, 1, 0, 0, 4, 0, 133, 0, 0, 0, "FICOMP ", ""
    AddAOC TblASM_OPCODE(497), &HDA, 1, 0, 0, 5, 0, 133, 0, 0, 0, "FISUB ", ""
    AddAOC TblASM_OPCODE(498), &HDA, 1, 0, 0, 6, 0, 133, 0, 0, 0, "FISUBR ", ""
    AddAOC TblASM_OPCODE(499), &HDA, 1, 0, 0, 7, 0, 133, 0, 0, 0, "FIDIV ", ""
    AddAOC TblASM_OPCODE(500), &HDA, 1, 0, 0, 8, 0, 133, 0, 0, 0, "FIDIVR ", ""
    AddAOC TblASM_OPCODE(501), &HC0DA, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FCMOVB ST(0),ST(", ")"
    AddAOC TblASM_OPCODE(502), &HC8DA, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FCMOVE ST(0),ST(", ")"
    AddAOC TblASM_OPCODE(503), &HD0DA, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FCMOVBE ST(0),ST(", ")"
    AddAOC TblASM_OPCODE(504), &HD8DA, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FCMOVU ST(0),ST(", ")"
    AddAOC TblASM_OPCODE(505), &HE9DA, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FUCOMPP", ""
    AddAOC TblASM_OPCODE(506), &HDB, 1, 0, 0, 0, 0, 133, 0, 0, 0, "FILD ", ""
    TblPtrASM(219) = 506
    AddAOC TblASM_OPCODE(507), &HDB, 1, 0, 0, 3, 0, 133, 0, 0, 0, "FIST ", ""
    AddAOC TblASM_OPCODE(508), &HDB, 1, 0, 0, 4, 0, 133, 0, 0, 0, "FISTP ", ""
    AddAOC TblASM_OPCODE(509), &HDB, 1, 0, 0, 6, 0, 130, 0, 0, 0, "FLD ", ""
    AddAOC TblASM_OPCODE(510), &HDB, 1, 0, 0, 8, 0, 130, 0, 0, 0, "FSTP ", ""
    AddAOC TblASM_OPCODE(511), &HC0DB, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FCMOVNB ST(0),ST(", ")"
    AddAOC TblASM_OPCODE(512), &HC8DB, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FCMOVNE ST(0),ST(", ")"
    AddAOC TblASM_OPCODE(513), &HD0DB, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FCMOVNBE ST(0),ST(", ")"
    AddAOC TblASM_OPCODE(514), &HD8DB, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FCMOVNU ST(0),ST(", ")"
    AddAOC TblASM_OPCODE(515), &HE2DB, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FNCLEX", ""
    AddAOC TblASM_OPCODE(516), &HE3DB, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FNINIT", ""
    AddAOC TblASM_OPCODE(517), &HE8DB, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FUCOMI ST,ST(", ""
    AddAOC TblASM_OPCODE(518), &HF0DB, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FCOMI ST,ST(", ""
    AddAOC TblASM_OPCODE(519), &HDC, 1, 0, 0, 0, 0, 129, 0, 0, 0, "FADD ", ""
    TblPtrASM(220) = 519
    AddAOC TblASM_OPCODE(520), &HDC, 1, 0, 0, 2, 0, 129, 0, 0, 0, "FMUL ", ""
    AddAOC TblASM_OPCODE(521), &HDC, 1, 0, 0, 3, 0, 129, 0, 0, 0, "FCOM ", ""
    AddAOC TblASM_OPCODE(522), &HDC, 1, 0, 0, 4, 0, 129, 0, 0, 0, "FCOMP ", ""
    AddAOC TblASM_OPCODE(523), &HDC, 1, 0, 0, 5, 0, 129, 0, 0, 0, "FSUB ", ""
    AddAOC TblASM_OPCODE(524), &HDC, 1, 0, 0, 6, 0, 129, 0, 0, 0, "FSUBR ", ""
    AddAOC TblASM_OPCODE(525), &HDC, 1, 0, 0, 7, 0, 129, 0, 0, 0, "FDIV ", ""
    AddAOC TblASM_OPCODE(526), &HDC, 1, 0, 0, 8, 0, 129, 0, 0, 0, "FDIVR ", ""
    AddAOC TblASM_OPCODE(527), &HC0DC, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FADD ST(", "),ST(0)"
    AddAOC TblASM_OPCODE(528), &HC8DC, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FMUL ST(", "),ST(0)"
    AddAOC TblASM_OPCODE(529), &HE0DC, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FSUBR ST(", "),ST(0)"
    AddAOC TblASM_OPCODE(530), &HE8DC, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FSUB ST(", "),ST(0)"
    AddAOC TblASM_OPCODE(531), &HF0DC, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FDIVR ST(", "),ST(0)"
    AddAOC TblASM_OPCODE(532), &HF8DC, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FDIV ST(", "),ST(0)"
    AddAOC TblASM_OPCODE(533), &HDD, 1, 0, 0, 0, 0, 129, 0, 0, 0, "FLD ", ""
    TblPtrASM(221) = 533
    AddAOC TblASM_OPCODE(534), &HDD, 1, 0, 0, 3, 0, 129, 0, 0, 0, "FST ", ""
    AddAOC TblASM_OPCODE(535), &HDD, 1, 0, 0, 4, 0, 129, 0, 0, 0, "FSTP ", ""
    AddAOC TblASM_OPCODE(536), &HDD, 1, 0, 0, 5, 0, 0, 0, 0, 0, "FRSTOR ", ""
    AddAOC TblASM_OPCODE(537), &HDD, 1, 0, 0, 7, 0, 0, 0, 0, 0, "FNSAVE ", ""
    AddAOC TblASM_OPCODE(538), &HDD, 1, 0, 0, 8, 0, 0, 0, 0, 0, "FNSTSW ", ""
    AddAOC TblASM_OPCODE(539), &HC0DD, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FFREE ST(", ")"
    AddAOC TblASM_OPCODE(540), &HD0DD, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FST ST(", ")"
    AddAOC TblASM_OPCODE(541), &HD8DD, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FSTP ST(", ")"
    AddAOC TblASM_OPCODE(542), &HE0DD, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FUCOM ST(", ")"
    AddAOC TblASM_OPCODE(543), &HE1DD, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FUCOM", ""
    AddAOC TblASM_OPCODE(544), &HE8DD, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FUCOMP ST(", ")"
    AddAOC TblASM_OPCODE(545), &HE9DD, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FUCOMP", ""
    AddAOC TblASM_OPCODE(546), &HDE, 1, 0, 0, 0, 0, 132, 0, 0, 0, "FIADD ", ""
    TblPtrASM(222) = 546
    AddAOC TblASM_OPCODE(547), &HDE, 1, 0, 0, 2, 0, 132, 0, 0, 0, "FIMUL ", ""
    AddAOC TblASM_OPCODE(548), &HDE, 1, 0, 0, 3, 0, 132, 0, 0, 0, "FICOM ", ""
    AddAOC TblASM_OPCODE(549), &HDE, 1, 0, 0, 4, 0, 132, 0, 0, 0, "FICOMP ", ""
    AddAOC TblASM_OPCODE(550), &HDE, 1, 0, 0, 5, 0, 132, 0, 0, 0, "FISUB ", ""
    AddAOC TblASM_OPCODE(551), &HDE, 1, 0, 0, 6, 0, 132, 0, 0, 0, "FISUBR ", ""
    AddAOC TblASM_OPCODE(552), &HDE, 1, 0, 0, 7, 0, 132, 0, 0, 0, "FIDIV ", ""
    AddAOC TblASM_OPCODE(553), &HDE, 1, 0, 0, 8, 0, 132, 0, 0, 0, "FIDIVR ", ""
    AddAOC TblASM_OPCODE(554), &HC0DE, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FADDP ST(", "),ST(0)"
    AddAOC TblASM_OPCODE(555), &HC1DE, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FADDP", ""
    AddAOC TblASM_OPCODE(556), &HC8DE, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FMULP ST(", "),ST(0)"
    AddAOC TblASM_OPCODE(557), &HC9DE, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FMULP", ""
    AddAOC TblASM_OPCODE(558), &HD9DE, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FCOMPP", ""
    AddAOC TblASM_OPCODE(559), &HE0DE, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FSUBRP ST(", "),ST(0)"
    AddAOC TblASM_OPCODE(560), &HE1DE, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FSUBRP", ""
    AddAOC TblASM_OPCODE(561), &HE8DE, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FSUBP ST(", "),ST(0)"
    AddAOC TblASM_OPCODE(562), &HE9DE, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FSUBP", ""
    AddAOC TblASM_OPCODE(563), &HF0DE, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FDIVRP ST(", "),ST(0)"
    AddAOC TblASM_OPCODE(564), &HF1DE, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FDIVRP", ""
    AddAOC TblASM_OPCODE(565), &HF8DE, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FDIVP ST(", "),ST(0)"
    AddAOC TblASM_OPCODE(566), &HF9DE, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FDIVP", ""
    AddAOC TblASM_OPCODE(567), &HDF, 1, 0, 0, 0, 0, 132, 0, 0, 0, "FILD ", ""
    TblPtrASM(223) = 567
    AddAOC TblASM_OPCODE(568), &HDF, 1, 0, 0, 3, 0, 132, 0, 0, 0, "FIST ", ""
    AddAOC TblASM_OPCODE(569), &HDF, 1, 0, 0, 4, 0, 132, 0, 0, 0, "FISTP ", ""
    AddAOC TblASM_OPCODE(570), &HDF, 1, 0, 0, 5, 0, 0, 0, 0, 0, "FBLD ", ""
    AddAOC TblASM_OPCODE(571), &HDF, 1, 0, 0, 6, 0, 134, 0, 0, 0, "FILD ", ""
    AddAOC TblASM_OPCODE(572), &HDF, 1, 0, 0, 7, 0, 0, 0, 0, 0, "FBSTP ", ""
    AddAOC TblASM_OPCODE(573), &HDF, 1, 0, 0, 8, 0, 134, 0, 0, 0, "FISTP ", ""
    AddAOC TblASM_OPCODE(574), &HE0DF, 2, 0, 0, 0, 0, 0, 0, 0, 0, "FNSTSW AX", ""
    AddAOC TblASM_OPCODE(575), &HE8DF, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FUCOMIP ST,ST(", ")"
    AddAOC TblASM_OPCODE(576), &HF0DF, 2, 0, 160, 0, 0, 0, 0, 0, 0, "FCOMIP ST,ST(", ")"
    AddAOC TblASM_OPCODE(577), &HE0, 1, 0, 0, 0, 22, 34, 0, 0, 0, "LOOPNE ", ""
    TblPtrASM(224) = 577
    AddAOC TblASM_OPCODE(578), &HE0, 1, 0, 0, 0, 22, 34, 0, 0, 0, "LOOPNZ ", ""
    AddAOC TblASM_OPCODE(579), &HE1, 1, 0, 0, 0, 22, 34, 0, 0, 0, "LOOPE ", ""
    TblPtrASM(225) = 579
    AddAOC TblASM_OPCODE(580), &HE1, 1, 0, 0, 0, 22, 34, 0, 0, 0, "LOOPZ ", ""
    AddAOC TblASM_OPCODE(581), &HE2, 1, 0, 0, 0, 22, 34, 0, 0, 0, "LOOP ", ""
    TblPtrASM(226) = 581
    AddAOC TblASM_OPCODE(582), &HE3, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JECXZ ", ""
    TblPtrASM(227) = 582
    AddAOC TblASM_OPCODE(583), &HE4, 1, 0, 0, 0, 26, 42, 0, 0, 0, "IN AL,", ""
    TblPtrASM(228) = 583
    AddAOC TblASM_OPCODE(584), &HE5, 1, 0, 0, 0, 26, 42, 0, 0, 0, "IN EAX,", ""
    TblPtrASM(229) = 584
    AddAOC TblASM_OPCODE(585), &HE6, 1, 0, 0, 0, 26, 42, 0, 0, 0, "OUT ", ",AL"
    TblPtrASM(230) = 585
    AddAOC TblASM_OPCODE(586), &HE7, 1, 0, 0, 0, 26, 42, 0, 0, 0, "OUT ", ",EAX"
    TblPtrASM(231) = 586
    AddAOC TblASM_OPCODE(587), &HE8, 1, 0, 0, 0, 24, 36, 0, 0, 0, "CALL ", ""
    TblPtrASM(232) = 587
    AddAOC TblASM_OPCODE(588), &HE9, 1, 0, 0, 0, 24, 36, 0, 0, 0, "JMP ", ""
    TblPtrASM(233) = 588
    AddAOC TblASM_OPCODE(589), &HEA, 1, 0, 0, 0, 0, 0, 0, 0, 0, "JMP pt", ""
    TblPtrASM(234) = 589
    AddAOC TblASM_OPCODE(590), &HEB, 1, 0, 0, 0, 22, 34, 0, 0, 0, "JMP ", ""
    TblPtrASM(235) = 590
    AddAOC TblASM_OPCODE(591), &HEC, 1, 0, 0, 0, 0, 0, 0, 0, 0, "IN AL,DX", ""
    TblPtrASM(236) = 591
    AddAOC TblASM_OPCODE(592), &HED, 1, 0, 0, 0, 0, 0, 0, 0, 0, "IN EAX,DX", ""
    TblPtrASM(237) = 592
    AddAOC TblASM_OPCODE(593), &HEE, 1, 0, 0, 0, 0, 0, 0, 0, 0, "OUT DX,AL", ""
    TblPtrASM(238) = 593
    AddAOC TblASM_OPCODE(594), &HEF, 1, 0, 0, 0, 0, 0, 0, 0, 0, "OUT DX,EAX", ""
    TblPtrASM(239) = 594
    AddAOC TblASM_OPCODE(595), &HF0, 1, 0, 0, 0, 0, 0, 0, 0, 0, "LOCK", ""
    TblPtrASM(240) = 595
    AddAOC TblASM_OPCODE(596), &HF1, 1, 0, 0, 0, 0, 0, 0, 0, 0, "INT1", ""
    TblPtrASM(241) = 596
    AddAOC TblASM_OPCODE(597), &HA6F2, 2, 0, 0, 0, 0, 51, 51, 0, 0, "REPNE CMPS ", ""
    TblPtrASM(242) = 597
    AddAOC TblASM_OPCODE(598), &HA7F2, 2, 0, 0, 0, 0, 53, 53, 0, 0, "REPNE CMPS ", ""
    AddAOC TblASM_OPCODE(599), &HAEF2, 2, 0, 0, 0, 0, 51, 0, 0, 0, "REPNE SCAS ", ""
    AddAOC TblASM_OPCODE(600), &HAFF2, 2, 0, 0, 0, 0, 53, 0, 0, 0, "REPNE SCAS ", ""
    AddAOC TblASM_OPCODE(601), &H6CF3, 2, 0, 0, 0, 0, 51, 0, 0, 0, "REP INS ", ",DX"
    TblPtrASM(243) = 601
    AddAOC TblASM_OPCODE(602), &H6DF3, 2, 0, 0, 0, 0, 53, 0, 0, 0, "REP INS ", ",DX"
    AddAOC TblASM_OPCODE(603), &H6EF3, 2, 0, 0, 0, 0, 51, 0, 0, 0, "REP OUTS DX,", ""
    AddAOC TblASM_OPCODE(604), &H6FF3, 2, 0, 0, 0, 0, 53, 0, 0, 0, "REP OUTS DX,", ""
    AddAOC TblASM_OPCODE(605), &HA4F3, 2, 0, 0, 0, 0, 51, 51, 0, 0, "REP MOVS ", ""
    AddAOC TblASM_OPCODE(606), &HA5F3, 2, 0, 0, 0, 0, 53, 53, 0, 0, "REP MOVS ", ""
    AddAOC TblASM_OPCODE(607), &HA6F3, 2, 0, 0, 0, 0, 51, 51, 0, 0, "REPE CMPS ", ""
    AddAOC TblASM_OPCODE(608), &HA7F3, 2, 0, 0, 0, 0, 53, 53, 0, 0, "REPE CMPS ", ""
    AddAOC TblASM_OPCODE(609), &HAAF3, 2, 0, 0, 0, 0, 51, 0, 0, 0, "REP STOS ", ""
    AddAOC TblASM_OPCODE(610), &HABF3, 2, 0, 0, 0, 0, 53, 0, 0, 0, "REP STOS ", ""
    AddAOC TblASM_OPCODE(611), &HACF3, 2, 0, 0, 0, 0, 0, 0, 0, 0, "REP LODS AL", ""
    AddAOC TblASM_OPCODE(612), &HADF3, 2, 0, 0, 0, 0, 0, 0, 0, 0, "REP LODS EAX", ""
    AddAOC TblASM_OPCODE(613), &HAEF3, 2, 0, 0, 0, 0, 51, 0, 0, 0, "REPE SCAS ", ""
    AddAOC TblASM_OPCODE(614), &HAFF3, 2, 0, 0, 0, 0, 53, 0, 0, 0, "REPE SCAS ", ""
    AddAOC TblASM_OPCODE(615), &HF4, 1, 0, 0, 0, 0, 0, 0, 0, 0, "HLT", ""
    TblPtrASM(244) = 615
    AddAOC TblASM_OPCODE(616), &HF5, 1, 0, 0, 0, 0, 0, 0, 0, 0, "CMC", ""
    TblPtrASM(245) = 616
    AddAOC TblASM_OPCODE(617), &HF6, 1, 0, 0, 3, 0, 18, 0, 0, 0, "NOT ", ""
    TblPtrASM(246) = 617
    AddAOC TblASM_OPCODE(618), &HF6, 1, 0, 0, 4, 0, 18, 0, 0, 0, "NEG ", ""
    AddAOC TblASM_OPCODE(619), &HF6, 1, 0, 0, 5, 0, 18, 0, 0, 0, "MUL ", ""
    AddAOC TblASM_OPCODE(620), &HF6, 1, 0, 0, 6, 0, 18, 0, 0, 0, "IMUL ", ""
    AddAOC TblASM_OPCODE(621), &HF6, 1, 0, 0, 7, 0, 18, 0, 0, 0, "DIV ", ""
    AddAOC TblASM_OPCODE(622), &HF6, 1, 0, 0, 8, 0, 18, 0, 0, 0, "IDIV ", ""
    AddAOC TblASM_OPCODE(623), &HF6, 1, 0, 0, 0, 26, 18, 42, 0, 0, "TEST ", ""
    AddAOC TblASM_OPCODE(624), &HF7, 1, 0, 0, 3, 0, 20, 0, 0, 0, "NOT ", ""
    TblPtrASM(247) = 624
    AddAOC TblASM_OPCODE(625), &HF7, 1, 0, 0, 4, 0, 20, 0, 0, 0, "NEG ", ""
    AddAOC TblASM_OPCODE(626), &HF7, 1, 0, 0, 5, 0, 20, 0, 0, 0, "MUL ", ""
    AddAOC TblASM_OPCODE(627), &HF7, 1, 0, 0, 6, 0, 20, 0, 0, 0, "IMUL ", ""
    AddAOC TblASM_OPCODE(628), &HF7, 1, 0, 0, 7, 0, 19, 0, 0, 0, "DIV ", ""
    AddAOC TblASM_OPCODE(629), &HF7, 1, 0, 0, 7, 0, 20, 0, 0, 0, "DIV ", ""
    AddAOC TblASM_OPCODE(630), &HF7, 1, 0, 0, 8, 0, 20, 0, 0, 0, "IDIV ", ""
    AddAOC TblASM_OPCODE(631), &HF7, 1, 0, 0, 0, 28, 20, 44, 0, 0, "TEST ", ""
    AddAOC TblASM_OPCODE(632), &HF8, 1, 0, 0, 0, 0, 0, 0, 0, 0, "CLC", ""
    TblPtrASM(248) = 632
    AddAOC TblASM_OPCODE(633), &HF9, 1, 0, 0, 0, 0, 0, 0, 0, 0, "STC", ""
    TblPtrASM(249) = 633
    AddAOC TblASM_OPCODE(634), &HFA, 1, 0, 0, 0, 0, 0, 0, 0, 0, "CLI", ""
    TblPtrASM(250) = 634
    AddAOC TblASM_OPCODE(635), &HFB, 1, 0, 0, 0, 0, 0, 0, 0, 0, "STI", ""
    TblPtrASM(251) = 635
    AddAOC TblASM_OPCODE(636), &HFC, 1, 0, 0, 0, 0, 0, 0, 0, 0, "CLD", ""
    TblPtrASM(252) = 636
    AddAOC TblASM_OPCODE(637), &HFD, 1, 0, 0, 0, 0, 0, 0, 0, 0, "STD", ""
    TblPtrASM(253) = 637
    AddAOC TblASM_OPCODE(638), &HFE, 1, 0, 0, 0, 0, 18, 0, 0, 0, "INC ", ""
    TblPtrASM(254) = 638
    AddAOC TblASM_OPCODE(639), &HFE, 1, 0, 0, 2, 0, 18, 0, 0, 0, "DEC ", ""
    AddAOC TblASM_OPCODE(640), &HFF, 1, 0, 0, 0, 0, 20, 0, 0, 0, "INC ", ""
    TblPtrASM(255) = 640
    AddAOC TblASM_OPCODE(641), &HFF, 1, 0, 0, 2, 0, 20, 0, 0, 0, "DEC ", ""
    AddAOC TblASM_OPCODE(642), &HFF, 1, 0, 0, 3, 0, 20, 0, 0, 0, "CALL ", ""
    AddAOC TblASM_OPCODE(643), &HFF, 1, 0, 0, 4, 0, 61, 0, 0, 0, "CALL ", ""
    AddAOC TblASM_OPCODE(644), &HFF, 1, 0, 0, 5, 0, 20, 0, 0, 0, "JMP ", ""
    AddAOC TblASM_OPCODE(645), &HFF, 1, 0, 0, 7, 0, 20, 0, 0, 0, "PUSH ", ""
    AddAOC TblASM_OPCODE(646), &HFF, 1, 0, 0, 17, 0, 61, 0, 0, 0, "JMP ", ""
   
End Sub

Private Sub AddAOC(tAOC As ASM_OPCODE, FOC As Integer, OpLen As Byte, _
                   f1 As Byte, f2 As Byte, f3 As Byte, f4 As Byte, _
                   f5 As Byte, f6 As Byte, f7 As Byte, f8 As Byte, _
                   EqvStr As String, EStr As String)
'paramètre une entrée ASM_OPCODE (balèze, hein!)

    tAOC.FullOpCode = FOC
    tAOC.OpCodeLen = OpLen
    tAOC.Flag1 = f1
    tAOC.Flag2 = f2
    tAOC.Flag3 = f3
    tAOC.Flag4 = f4
    tAOC.Flag5 = f5
    tAOC.Flag6 = f6
    tAOC.Flag7 = f7
    tAOC.Flag8 = f8
    tAOC.sInstruct = EqvStr
    tAOC.sEnd = EStr
    
End Sub

