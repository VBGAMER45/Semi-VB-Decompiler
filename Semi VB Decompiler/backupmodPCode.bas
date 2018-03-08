Attribute VB_Name = "modPCode"
'#############Begin Information#############
'Informative:
'Takes no bytes tells how to process data
'>'      put the following hex in subsegments up to next
'         offset following ArgStr char should be "p" for
'         Procedure Address
'h'      return hex output of following typechars. possible(°,%,&);
'}'      End Procedure

'Arguments
'Will usually take bytes from the datastream

'.'      name Of Object at the Address specified by a Long off the datastream
'b'      a byte off the datastream - formerly '°'
'%'      an integer off the datastream
'&'      a long off the datastream
'!'      a single off the datastream
'a'      an argument reference. Followed by an Int and a type char.  Takes variable out of the ConstantPool
'c'      return the control index,uses one int from the datastream
'l'      return Local variable reference(uses int off datastream)
'L'      take (Value of Int off DataStream) local variable references
'm'      return Local Variable reference followed by typechar
'n'      return hex Integer
'o'      return item off the stack(Pop)
'p'      return (value of Integer  off datastream) + Procedure Base Address
't'      followed by typechar('o' return ObjectName;'c' return control name)
'u'      push...not used anymore
'v'      vTable this is slightly complicated ;)
'z'      return Null-Termed Unicode String From File(not used?)


'Type Characters
' b     Byte
' ?     Boolean
' %     Integer
' !     Single
' &    Long
' ~     Variant
' z     String


'Pcode Opcode Meanings
'Imp=import
'Ad = Address
'St/Ld=Store/Load
'I2 = Integer
'Lit=Literal(ie "Hi",2,8 )
'Cy=Currency
'R4=
'R8=Single
'Str=String
'DOC=Duplicate/Redundante Opcode(in the table it will redirect you to another opcode)
'#############End Information#############
Option Explicit

Private Type OpcodeType
    Mnemonic    As String
    Size        As Long
    Flag        As Byte
End Type



Private Type ObjEntry
    ObjectName(7) As Byte
    VirtualSize   As Long
    SectionRVA   As Long
    PhysicalSize   As Long
    PhysicalOffset   As Long
    Reserved(11)   As Byte
    ObjectFlags   As Long
End Type

Private Type ImportTable
    LookUpRVA  As Long
    TimeDateStamp  As Long
    Chains  As Long
    NameRVA  As Long
    AddressTableRVA  As Long
End Type

Private Type RecordTableInfo
    Data00  As Long
    Vftable As Long
    Layout  As Long
    data0C  As Long
    data10  As Long
    Data14  As Long
    Info(15)  As Byte
    Flag    As Integer
    Len     As Integer
    len2    As Integer
    Len3    As Integer
    RecAddr As Long
    unk(2) As Long
    NameTab As Long
End Type

Private Type RecordType
    TabAddr As Long
    Data04  As Long
    Import  As Long
    data0C  As Long
    data10  As Long
    Data14  As Long
    ModName As Long
    Owner   As Long
    Names   As Long
    data20  As Long
    data24  As Long
    data2C  As Long
End Type

Private Type ProcDscInfo 'CodeInfo
    table           As Long
    field_4         As Integer
    FrameSize       As Integer '24
    ProcSize        As Integer '22
    field_A         As Integer '20
    field_C         As Integer '18
    field_E         As Integer '16
    field_10        As Integer '14
    field_12        As Integer '12
    field_14        As Integer '10
    field_16        As Integer '8
    field_18        As Integer '6
    field_1A        As Integer '4
    Flag            As Integer '2
End Type

Private Type TableInfo 'ObjectInfo
    data0       As Long
    Record      As Long '4
    data8       As Long
    data0C      As Long
    data10      As Long
    Owner       As Long '14
    rtl         As Long '18
    data1C      As Long
    data20      As Long
    data24      As Long
    JmpCnt      As Integer  '28
    data2A      As Integer
    data2C      As Long
    data30      As Long
    ConstPool   As Long
End Type

Global Const unkno = 0
Global Const std = 1
Global Const idx = 2

Global Const none = 99
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private base&, start&, PESize&, Table1&, Table2&, Table3&, RecTable&

Private ObjTable(10) As ObjEntry
Private ImpTab(50) As ImportTable
Private Record(256) As RecordType
Private RecordNames(256) As String
Private RecordTable As RecordTableInfo
Private OPCode(5, 255) As OpcodeType
Private File() As Byte
Private Map() As Byte
Public SubName() As String
Private RefName() As String
Private ProcList() As Long
Private ProcCnt As Long
Global EventProcList() As Long
Global bSkipProcedure As Boolean
'Holds SubNames from OpenVBExe
Private Type subNameListType
    strName As String
    offset As Long
End Type
Global SubNamelist() As subNameListType

Sub Decode(Filename As String)

'*****************************
'Purpose: To decode a P-Code excutable and return all procedures in P-Code tokens
'*****************************
    Dim c As Long, a As Long
    Dim i As Long
    Dim f As Integer
    Dim f2 As Integer
    ReDim ProcList(0)
    ProcCnt = 0

    LoadPE2 Filename
    LoadPcode2
    
    frmMain.cmdSkipProcedure.Visible = True
    
    Dim ProcAddr() As Long
    Dim g As Integer
    'Get all procedures
    f = FreeFile
    Open SFilePath For Binary Access Read As #f
    For i = 0 To UBound(gObjectInfoHolder)
        If gObjectInfoHolder(i).NumberOfProcs > 0 Then
        ReDim ProcAddr(gObjectInfoHolder(i).NumberOfProcs - 1)
        Seek #f, gObjectInfoHolder(i).aProcTable + 1 - OptHeader.ImageBase
        Get #f, , ProcAddr
        For g = 0 To UBound(ProcAddr)
            If ProcAddr(g) <> 0 And ProcAddr(g) <> -1 Then
                If ProcAddr(g) < UBound(SubName) And ProcAddr(g) > LBound(SubName) Then
                SubName(ProcAddr(g)) = gObjectNameArray(i) & ".Proc" & ProcAddr(g)
                AddProc ProcAddr(g)
                End If
            End If
        Next
        End If
    Next
        Dim addrSubMain As Long
        If gVBHeader.aSubMain <> 0 Then
            Seek #f, gVBHeader.aSubMain + 2 - OptHeader.ImageBase
            Get #f, , addrSubMain
          Dim stemp
            stemp = Split(SubName(addrSubMain), ".")
            SubName(addrSubMain) = stemp(0) & ".Sub Main"
        End If
    Close #f
    'Add Event ProcLists
    
    For i = 0 To UBound(EventProcList) - 1
        If EventProcList(i) <> 0 Then
        AddProc EventProcList(i)
        'MsgBox "Added" & EventProcList(i)
        End If
    Next
    'For i = 1 To UBound(exeIMPORT_APINAME)
       'On Error Resume Next
        'SubName(exeIMPORT_APINAME(i).Address) = exeIMPORT_APINAME(i).ApiName
   ' Next
    For i = 0 To UBound(SubNamelist) - 1
        If SubNamelist(i).offset < UBound(SubName) Then
            SubName(SubNamelist(i).offset) = SubNamelist(i).strName
        End If
      '  MsgBox "SubName:" & SubNamelist(i).Offset
    Next
    'Reset Registers
    Call modPCodeToVB.ResetAsmRegister
    

    f = FreeFile
   
    
    Open App.Path & "\dump\" & SFile & "\PcodeToVB.txt" For Output As #f
     f2 = FreeFile
    Open App.Path & "\dump\" & SFile & "\PcodeOut.txt" For Output As #2
        Print #f2, "Semi VB Decompiler - VisualBasicZone.com"
        Print #f2, "P-Code Output for : " & Filename
        Print #f2, "---------------------------------"
        Print #f, "Semi VB Decompiler - VisualBasicZone.com"
        Print #f, "P-Code To VB Output for : " & Filename
        Print #f, "---------------------------------"
        frmMain.txtStatus.Text = frmMain.txtStatus.Text & "P-Code Procedure Count: " & ProcCnt & vbCrLf
        frmMain.txtStatus.Refresh
        Do
            c = 0
            For a = 0 To ProcCnt - 1
                If ProcList(a) <> 0 Then
                    'MsgBox a & " " & ProcList(a)
                    frmMain.txtStatus.Text = frmMain.txtStatus.Text & "Decoding procedure: " & ProcList(a) & vbCrLf
                    frmMain.txtStatus.Refresh
                    Print #f2, DecompileProc(ProcList(a))
                    Print #f2, ""
                    Print #f, DecompileProcToVB(ProcList(a))

                    
                    ProcList(a) = 0
                    c = 1
                End If
            Next
        Loop While c
    Close #f2
    Close #f

    frmMain.cmdSkipProcedure.Visible = False
End Sub
Sub LoadPE2(strFileName As String)
'*****************************
'Purpose: To get the PE data of the filename
'*****************************
    Dim a As Long, l As Long, b As Long, t() As Byte
    Dim addrtab As Long, libname As String, fname As String, i As Long, tmp As String, func As Long
    Dim f As Long
    f = FreeFile
    Open strFileName For Binary As #f

    For a = 0 To PEHeader.NumSections - 1
        l = SecHeader(a).Misc + SecHeader(a).Address
        If PESize < l Then PESize = l
        l = SecHeader(a).Misc + SecHeader(a).Address
        If PESize < l Then PESize = l
    Next
    base& = OptHeader.ImageBase
    ReDim File(base To base + PESize)
    ReDim Map(base To base + PESize)
    ReDim SubName(base To base + PESize)
    ReDim RefName(base To base + PESize)
    For a = 0 To PEHeader.NumSections - 1
        If SecHeader(a).SizeRawData > 0 Then
            Seek #1, SecHeader(a).RawDataPointer + 1
            t() = InputB(SecHeader(a).SizeRawData, #f)
            l& = SecHeader(a).Address + base
            CopyMemory File(l), t(0), SecHeader(a).SizeRawData
        End If
    Next
    Close #f
    
    i& = 0
    CopyMemory ImpTab(0), File(OptHeader.DataDirectory(1).Address + base), OptHeader.DataDirectory(1).Size
    Do While ImpTab(i).LookUpRVA <> 0
        func = ImpTab(i).LookUpRVA + base
        tmp = Hex(func)
        libname = FileZ(base + ImpTab(i).NameRVA)
        addrtab = ImpTab(i).AddressTableRVA + base
        Do While File32(func) > 0 And File32(func) < PESize
            fname = FileZ(File32(func) + base + 2)
            SubName(addrtab) = libname + "." + fname
          
            addrtab = addrtab + 4
            func = func + 4
        Loop
        i = i + 1
    Loop

  


End Sub


Sub LoadPcode2()
    Dim a&
    Table1 = File32(OptHeader.EntryPoint + base + 1)
    Table3 = File32(Table1 + &H30)
    RecTable = File32(Table3 + 4)
    CopyMemory RecordTable, File(RecTable), Len(RecordTable)
    For a = 0 To RecordTable.Len - 1
        CopyMemory Record(a), File(RecordTable.RecAddr + &H30 * a), &H30
        RecordNames(a) = FileZ(Record(a).ModName)
    Next
End Sub
Function DecompileProc$(addr As Long)
'*****************************
'Purpose: Decompile a P-Code procedure from a CodeInfo address
'*****************************
    Dim i&, t&, t2&, u$, m As OpcodeType, a&, q$
    Dim pd As ProcDscInfo, output$, pc&, pass&, sp$, sr&
    Dim tt As TableInfo, tt2 As TableInfo
    On Error Resume Next
    sp$ = Space$(14)
    CopyMemory pd, File(addr), Len(pd)
    CopyMemory tt, File(pd.table), Len(tt)
    pc = addr - pd.ProcSize
    'pc = 4198804
    
    output = Hex(pc) + " " + FastName("proc", addr) + ":" + Chr(13) + Chr(10)
    For pass = 0 To 1
        i = pc
        Do
            DoEvents
            If bSkipProcedure = True Then
                bSkipProcedure = False
                Exit Function
            End If
            
            If HasRef(i) And pass Then
                output = output + Chr(13) + Chr(10) + Hex(i) + " " + FastName("loc", i) + ":"
                'output = output + Chr(13) + Chr(10) + " " + FastName("loc", i) + ":"
                output = output + Chr(9) + Chr(9) + Chr(9) + "; " + RefName(i) + Chr(13) + Chr(10)
            End If
            u = Hex(i) + sp
            AddMap i
            t = File(i)
            i = i + 1
            u$ = u + MakeHex(t)
            Select Case t
                Case &HFB To &HFF
                    AddMap i
                    t2 = File(i)
                    i = i + 1
                    u$ = u$ + MakeHex(t2)
                    m = OPCode(t - &HFB + 1, t2)
                Case Else
                    m = OPCode(0, t)
            End Select
            u$ = u$ + " "
            If m.Size > 0 Then
                For a = 1 To m.Size - 1
                    u = u + MakeHex(0 + File(i + a - 1))
                Next
            End If
            u = Left(u$ + Space(32), 38)
            If m.Size > 0 Then
                u = u + Chr(9)
            Else
                u = u + Chr(9)
                q$ = vbNullString
                t = File16(i)
                i = i + 2
                If t < 48 Then
                    For a = 0 To t - 2 Step 2
                        t2 = File16(i + a)
                        u = u + MakeArg(t2) + " "
                    Next
                    i = i + t
                End If
            End If
            
            Select Case m.Flag
                Case std
                    u = u + ConvertStr(m.Mnemonic, i, tt.ConstPool, tt.ConstPool, pc)
                Case idx
                    u = u + ConvertStr(m.Mnemonic, i, sr, tt.ConstPool, pc)
                Case none
                    u = u + (m.Mnemonic)
                    
                Case Else
                    u = u + m.Mnemonic + "  ???"
            End Select
            u = u + Chr(13) + Chr(10)
            If pass Then output = output + u
            If m.Size > 0 Then i = i + m.Size - 1
        Loop While i < addr
    Next
    
    DecompileProc = output
End Function
Function DecompileProcToVB(addr As Long, Optional bReturnProcedureName As Boolean = False) As String
'*****************************
'Purpose: Decompile a P-Code procedure from a CodeInfo address
'*****************************
    Dim i As Long, t As Long, t2 As Long, u As String, m As OpcodeType, a&, q$
    Dim pd As ProcDscInfo, output As String, pc&, pass&, sp$, sr&
    Dim tt As TableInfo, tt2 As TableInfo
    sp$ = Space$(14)
    On Error Resume Next
    CopyMemory pd, File(addr), Len(pd)
    CopyMemory tt, File(pd.table), Len(tt)
    pc = addr - pd.ProcSize
  
    
    'output = Hex(pc) + " " + FastName("proc", addr) + ":" + Chr(13) + Chr(10)
    If bReturnProcedureName = True Then
        DecompileProcToVB = FastName("proc", addr)
        Exit Function
    Else
        output = output & "Sub " & FastName("proc", addr) + "()" & vbCrLf
    End If
    output = output & "'ProcInfo: StartAddress=" & Hex(pc) & " ProcSize: " & pd.ProcSize & vbCrLf
    For pass = 0 To 1
        i = pc
        Do
            DoEvents
            If bSkipProcedure = True Then
                bSkipProcedure = False
                Exit Function
            End If
            If HasRef(i) And pass Then
                output = output + Chr(13) + Chr(10) + Hex(i) + " " + FastName("loc", i) + ":"
                'output = output + Chr(13) + Chr(10) + " " + FastName("loc", i) + ":"
                output = output + Chr(9) + Chr(9) + Chr(9) + "; " + RefName(i) + Chr(13) + Chr(10)
            End If
           ' u = Hex(i) + sp
            AddMap i
            t = File(i)
            i = i + 1
           ' u$ = u + MakeHex(t)
            Select Case t
                Case &HFB To &HFF
                    AddMap i
                    t2 = File(i)
                    i = i + 1
                   ' u$ = u$ + MakeHex(t2)
                    m = OPCode(t - &HFB + 1, t2)
                    
                Case Else
                    m = OPCode(0, t)
            End Select
            'u$ = u$ + " "
            If m.Size > 0 Then
                For a = 1 To m.Size - 1
                  '  u = u + MakeHex(0 + File(i + a - 1))
                Next
            End If
           ' u = Left(u$ + Space(32), 38)
            If m.Size > 0 Then
             '   u = u + Chr(9)
            Else
             '   u = u + Chr(9)
                q$ = vbNullString
                t = File16(i)
                i = i + 2
                If t < 48 Then
                    For a = 0 To t - 2 Step 2
                        t2 = File16(i + a)
                      '  u = u + MakeArg(t2) + " "
                    Next
                    i = i + t
                End If
            End If
            u = vbNullString
            'Ident Code a little
            'u = u & Space(5)
            'u = u & modPCodeToVB.ReturnVBCodeByPcodeToken(i, m.Mnemonic)
            'MsgBox m.Mnemonic
            'MsgBox modPCodeToVB.ReturnVBCodeByPcodeToken(i, m.Mnemonic)
            Select Case m.Flag
                Case std
                  '#  u = u + ConvertStr(m.Mnemonic, i, tt.ConstPool, tt.ConstPool, pc)
                  u = u & modPCodeToVB.ReturnVBCodeByPcodeToken(i, m.Mnemonic, ConvertStrToVB(m.Mnemonic, i, tt.ConstPool, tt.ConstPool, pc))
                Case idx
                  '#  u = u + ConvertStr(m.Mnemonic, i, sr, tt.ConstPool, pc)
                    u = u & modPCodeToVB.ReturnVBCodeByPcodeToken(i, m.Mnemonic, ConvertStrToVB(m.Mnemonic, i, sr, tt.ConstPool, pc))
                Case none
                  '#  u = u + (m.Mnemonic)
                    u = u & modPCodeToVB.ReturnVBCodeByPcodeToken(i, m.Mnemonic, "-1")
                Case Else
                 u = u & modPCodeToVB.ReturnVBCodeByPcodeToken(i, m.Mnemonic, "-1")
                 '#   u = u + m.Mnemonic + "  ???"
            End Select
            u = u + Chr(13) + Chr(10)
            If pass Then output = output + u
            If m.Size > 0 Then i = i + m.Size - 1
        Loop While i < addr
    Next
    'output = output & "End Sub" & vbCrLf
    DecompileProcToVB = output
End Function

Function ConvertStr(Mnem As String, addr As Long, pool As Long, origpool As Long, ProcPC As Long)
'Arguement Type
    
    Dim a&, c$, t&, u$, i&, j&
    
    i = addr
    For a = 1 To Len(Mnem)
        c = Mid(Mnem, a, 1)
        If c <> "%" Then
            u = u + c
        Else
            a = a + 1
            c = Mid(Mnem, a, 1)
            Select Case c
                Case "a"
                    t = File16(i)
                    u = u + MakeArg(t)
                    i = i + 2
                Case "c"
                    t = File32(File16(i) * 4 + pool)
                    u = u + MakeAddr(t)
                    i = i + 2
                Case "e"
                    t = File32(File16(i) * 4 + pool) + File16(i + 2)
                    u = u + MakeAddr(t)
                    i = i + 4
                Case "s"
                    t = File32(File16(i) * 4 + pool)
                    u = u + FastName("v", t) + " '" + FileW(t) + "' "
                    i = i + 2
                Case "l"
                    t = ProcPC + File16(i)
                    u = u + FastName("loc", t)
                    AddRef t, i - 1
                    i = i + 2
                Case "1", "2", "4"
                    For j = 1 To Val(c)
                        u = u + MakeHex(File(i + Val(c) - j))
                    Next
                    i = i + Val(c)
                Case "t"
                    t = File32(File16(i) * 4 + origpool)
                    u = u + FastName("xxx", t)
                    pool = t
                Case Else
              
                    u = u + c
            End Select
        End If
    Next
    ConvertStr = u
End Function
Function MakeArg$(t As Long)
        If t < 0 Then
            MakeArg = "var_" + Hex(-t)
        Else
            MakeArg = "arg_" + Hex(t)
        End If
End Function

Sub Init0()
    MakeOpcode 0, &H0, 2, 0, "LargeBos"
    MakeOpcode 0, &H1, 5, 0, "InvalidExcode"
    MakeOpcode 0, &H2, 2, 0, "SelectCaseByte"
    MakeOpcode 0, &H3, 1, 0, "---"
    MakeOpcode 0, &H4, 3, std, "FLdRfVar ::lea %a"
    MakeOpcode 0, &H5, 3, std, "ImpAdLdRf ::lea %c"
    MakeOpcode 0, &H6, 3, 0, "MemLdRfVar"
    MakeOpcode 0, &H7, 5, 0, "FMemLdRf"
    MakeOpcode 0, &H8, 3, std, "FLdPr ::mov SR,%a"
    MakeOpcode 0, &H9, 5, std, "ImpAdCallHresult ::call %c(%a)"
    MakeOpcode 0, &HA, 5, std, "ImpAdCallFPR4 ::call %c(%a)"
    MakeOpcode 0, &HB, 5, std, "ImpAdCallI2 ::call %c(%a)"
    MakeOpcode 0, &HC, 5, std, "ImpAdCallCy ::call %c(%a)"
    'MakeOpcode 0, &HD, 5, std, "VCallHresult ::call %c(%c)"
     MakeOpcode 0, &HD, 5, std, "VCallHresult ::call %c(%c)"
    MakeOpcode 0, &HE, 3, 0, "VCallFPR8"
    MakeOpcode 0, &HF, 3, std, "VCallAd"
    MakeOpcode 0, &H10, 5, 0, "ThisVCallHresult"
    MakeOpcode 0, &H11, 3, 0, "ThisVCall"
    MakeOpcode 0, &H12, 3, 0, "ThisVCallAd"
    MakeOpcode 0, &H13, 1, none, "ExitProcHresult ::ret"
    MakeOpcode 0, &H14, 1, none, "ExitProc ::ret"
    MakeOpcode 0, &H15, 1, none, "ExitProcI2 ::retw"
    MakeOpcode 0, &H16, 1, none, "ExitProcR4 ::retf"
    MakeOpcode 0, &H17, 1, none, "ExitProcR8 ::retf8"
    MakeOpcode 0, &H18, 1, none, "ExitProcCy ::retc"
    MakeOpcode 0, &H19, 3, 0, "FStAdFunc"
    MakeOpcode 0, &H1A, 3, 0, "FFree1Ad"
    MakeOpcode 0, &H1B, 3, std, "LitStr ::lea %s"
    MakeOpcode 0, &H1C, 3, std, "BranchF ::jnz %l"
    MakeOpcode 0, &H1D, 3, std, "BranchT ::jz %l"
    MakeOpcode 0, &H1E, 3, std, "Branch ::jmp %l"
    MakeOpcode 0, &H1F, 3, std, "CRec2Ansi %2"
    MakeOpcode 0, &H20, 3, 0, "CRec2Uni"
    MakeOpcode 0, &H21, 1, none, "FLdPrThis"
    MakeOpcode 0, &H22, 3, std, "ImpAdLdPr ::push [%c]"
    MakeOpcode 0, &H23, 3, 0, "FStStrNoPop"
    MakeOpcode 0, &H24, 3, idx, "NewIfNullPr ::newnull %t"
    MakeOpcode 0, &H25, 1, none, "PopAdLdVar"
    MakeOpcode 0, &H26, 3, 0, "AryDescTemp"
    MakeOpcode 0, &H27, 3, 0, "LitVar_Missing"
    MakeOpcode 0, &H28, 5, 0, "LitVarI2 ::mov %a,%2"
    MakeOpcode 0, &H29, -1, none, "FFreeAd:"
    MakeOpcode 0, &H2A, 1, none, "ConcatStr"
    MakeOpcode 0, &H2B, 3, 0, "PopTmpLdAd2"
    MakeOpcode 0, &H2C, 5, 0, "LateIdSt"
    MakeOpcode 0, &H2D, 3, 0, "AryUnlock"
    MakeOpcode 0, &H2E, 3, 0, "AryLock"
    MakeOpcode 0, &H2F, 3, 0, "FFree1Str"
    MakeOpcode 0, &H30, 3, 0, "PopTmpLdAd8"
    MakeOpcode 0, &H31, 3, 0, "FStStr"
    MakeOpcode 0, &H32, -1, none, "FFreeStr"
    MakeOpcode 0, &H33, 3, std, "LdFixedStr ::lea %s"
    MakeOpcode 0, &H34, 1, none, "CStr2Ansi"
    MakeOpcode 0, &H35, 3, 0, "FFree1Var"
    MakeOpcode 0, &H36, -1, none, "FFreeVar"
    MakeOpcode 0, &H37, 1, none, "PopFPR4"
    MakeOpcode 0, &H38, 3, 0, "CopyBytes"
    MakeOpcode 0, &H39, 1, none, "PopFPR8"
    'MakeOpcode 0, &H3A, 5, 0, "LitVarStr"
    MakeOpcode 0, &H3A, 5, std, "LitVarStr :: %a %s"
    MakeOpcode 0, &H3B, 1, none, "Ary1StStrCopy"
    MakeOpcode 0, &H3C, 1, none, "SetLastSystemError"
    MakeOpcode 0, &H3D, 3, 0, "CastAd"
    MakeOpcode 0, &H3E, 3, 0, "FLdZeroAd"
    MakeOpcode 0, &H3F, 3, 0, "CVarCy"
    MakeOpcode 0, &H40, 1, none, "Ary1LdRf"
    MakeOpcode 0, &H41, 1, none, "Ary1LdPr"
    MakeOpcode 0, &H42, 1, none, "CR4Var"
    MakeOpcode 0, &H43, 3, std, "FStStrCopy ::strcpy %a"
    MakeOpcode 0, &H44, 3, 0, "CVarI2"
    MakeOpcode 0, &H45, 1, none, "Error"
    MakeOpcode 0, &H46, 3, 0, "CVarStr"
    MakeOpcode 0, &H47, 3, std, "StFixedStr %l %a"
    MakeOpcode 0, &H48, 3, 0, "ILdPr"
    MakeOpcode 0, &H49, 1, none, "PopAdLd4"
    MakeOpcode 0, &H4A, 1, none, "FnLenStr ::strlen"
    MakeOpcode 0, &H4B, 3, std, "OnErrorGoto %l"
    MakeOpcode 0, &H4C, 1, none, "FnLBound"
    MakeOpcode 0, &H4D, 5, 0, "CVarRef:"
    MakeOpcode 0, &H4E, 3, 0, "FStVarCopyObj"
    MakeOpcode 0, &H4F, 3, 0, "MidStr"
    MakeOpcode 0, &H50, 1, none, "CI4Str"
    MakeOpcode 0, &H51, 3, 0, "FLdZeroAd"
    MakeOpcode 0, &H52, 1, none, "Ary1StVar"
    MakeOpcode 0, &H53, 1, none, "CBoolCy"
    MakeOpcode 0, &H54, 5, 0, "FMemStStrCopy"
    MakeOpcode 0, &H55, 1, none, "CI2Var"
    MakeOpcode 0, &H56, 3, 0, "NewIfNullAd"
    MakeOpcode 0, &H57, 5, 0, "LateMemLdVar"
    MakeOpcode 0, &H58, 3, 0, "MemLdPr"
    MakeOpcode 0, &H59, 3, 0, "PopTmpLdAdStr"
    MakeOpcode 0, &H5A, 1, none, "Erase"
    MakeOpcode 0, &H5B, 3, 0, "FStAdFuncNoPop"
    MakeOpcode 0, &H5C, 3, 0, "BranchFVar"
    MakeOpcode 0, &H5D, 1, none, "HardType"
    MakeOpcode 0, &H5E, 11, std, "call %c(%a)"
    MakeOpcode 0, &H5F, 5, 0, "FMemLdPr"
    MakeOpcode 0, &H60, 1, none, "CStrVarTmp"
    MakeOpcode 0, &H61, 7, 0, "LateIdLdVar"
    MakeOpcode 0, &H62, 3, 0, "IStDarg"
    MakeOpcode 0, &H63, 3, 0, "LitVar_TRUE"
    MakeOpcode 0, &H64, 5, 0, "NextI2:"
    MakeOpcode 0, &H65, 5, 0, "NextStepI2:"
    MakeOpcode 0, &H66, 5, 0, "NextI4:"
    MakeOpcode 0, &H67, 5, 0, "NextStepI4:"
    MakeOpcode 0, &H68, 5, 0, "NextStepR4:"
    MakeOpcode 0, &H69, 5, 0, "NextStepR8:"
    MakeOpcode 0, &H6A, 5, 0, "NextStepCy"
    MakeOpcode 0, &H6B, 3, std, "FLdI2 ::push [%a]"
    MakeOpcode 0, &H6C, 3, std, "ILdRf ::push [%a]"
    MakeOpcode 0, &H6D, 3, 0, "FLdR8 ::push"
    MakeOpcode 0, &H6E, 3, 0, "FLdFPR4"
    MakeOpcode 0, &H6F, 3, 0, "FLdFPR8"
    MakeOpcode 0, &H70, 3, std, "FStI2 ::pop [%a]"
    MakeOpcode 0, &H71, 3, std, "FStR4 ::pop [%a]"
    MakeOpcode 0, &H72, 3, 0, "FStR8"
    MakeOpcode 0, &H73, 3, 0, "FStFPR4"
    MakeOpcode 0, &H74, 3, 0, "FStFPR8"
    MakeOpcode 0, &H75, 3, 0, "ImpAdLdI2"
    MakeOpcode 0, &H76, 3, std, "ImpAdLdI4 ::push [%c]"
    MakeOpcode 0, &H77, 3, std, "ImpAdLdCy %c"
    MakeOpcode 0, &H78, 3, std, "ImpAdLdFPR4 %c"
    MakeOpcode 0, &H79, 3, std, "ImpAdLdFPR8 %c"
    MakeOpcode 0, &H7A, 3, std, "ImpAdStI2 %c"
    MakeOpcode 0, &H7B, 3, std, "ImpAdStR4 %c"
    MakeOpcode 0, &H7C, 3, std, "ImpAdStCy %c"
    MakeOpcode 0, &H7D, 3, std, "ImpAdStFPR4 %c"
    MakeOpcode 0, &H7E, 3, std, "ImpAdStFPR8 %c"
    MakeOpcode 0, &H7F, 3, std, "ILdI2 %c"
    MakeOpcode 0, &H80, 3, std, "ILdI4 %c"
    MakeOpcode 0, &H81, 3, std, "ILdR8 %c"
    MakeOpcode 0, &H82, 3, std, "ILdFPR4 %c"
    MakeOpcode 0, &H83, 3, std, "ILdFPR8 %c"
    MakeOpcode 0, &H84, 3, std, "IStI2 %c"
    MakeOpcode 0, &H85, 3, std, "IStI4 %c"
    MakeOpcode 0, &H86, 3, 0, "IStR8"
    MakeOpcode 0, &H87, 3, 0, "IStFPR4"
    MakeOpcode 0, &H88, 3, 0, "IStFPR8"
    MakeOpcode 0, &H89, 3, idx, "MemLdI2 ::push [%2+SR]"
    MakeOpcode 0, &H8A, 3, 0, "MemLdStr"
    MakeOpcode 0, &H8B, 3, 0, "MemLdR8"
    MakeOpcode 0, &H8C, 3, 0, "MemLdFPR4"
    MakeOpcode 0, &H8D, 3, 0, "MemLdFPR8"
    MakeOpcode 0, &H8E, 3, 0, "MemStI2"
    MakeOpcode 0, &H8F, 3, 0, "MemStI4"
    MakeOpcode 0, &H90, 3, 0, "MemStR8"
    MakeOpcode 0, &H91, 3, 0, "MemStFPR4"
    MakeOpcode 0, &H92, 3, 0, "MemStFPR8"
    MakeOpcode 0, &H93, 5, 0, "FMemLdI2"
    MakeOpcode 0, &H94, 5, 0, "FMemLdR4"
    MakeOpcode 0, &H95, 5, 0, "FMemLdCy"
    MakeOpcode 0, &H96, 5, 0, "FMemLdFPR4"
    MakeOpcode 0, &H97, 5, 0, "FMemLdFPR8"
    MakeOpcode 0, &H98, 5, std, "FMemStI2 ::popw [[%a]+%2]"
    MakeOpcode 0, &H99, 5, std, "FMemStI4 ::popd [[%a]+%2]"
    MakeOpcode 0, &H9A, 5, std, "FMemStR8 ::popf [[%a]+%2]"
    MakeOpcode 0, &H9B, 5, std, "FMemStFPR4 ::popf [[%a]+%2]"
    MakeOpcode 0, &H9C, 5, std, "FMemStFPR8 ::popf [[%a]+%2]"
    MakeOpcode 0, &H9D, 1, none, "Ary1LdI2"
    MakeOpcode 0, &H9E, 1, none, "Ary1LdI4"
    MakeOpcode 0, &H9F, 1, none, "Ary1LdCy"
    MakeOpcode 0, &HA0, 1, none, "Ary1LdFPR4"
    MakeOpcode 0, &HA1, 1, none, "Ary1LdFPR8"
    MakeOpcode 0, &HA2, 1, none, "Ary1StI2"
    MakeOpcode 0, &HA3, 1, none, "Ary1StI4"
    MakeOpcode 0, &HA4, 1, none, "Ary1StCy"
    MakeOpcode 0, &HA5, 1, none, "Ary1StFPR4"
    MakeOpcode 0, &HA6, 1, none, "Ary1StFPR8"
    MakeOpcode 0, &HA7, 3, 0, "AryLdPr"
    MakeOpcode 0, &HA8, 3, 0, "AryLdRf"
    MakeOpcode 0, &HA9, 1, none, "AddI2 ::addw"
    MakeOpcode 0, &HAA, 1, none, "AddI4 ::addd"
    MakeOpcode 0, &HAB, 1, none, "AddR8 ::addf"
    MakeOpcode 0, &HAC, 1, none, "AddCy ::addc"
    MakeOpcode 0, &HAD, 1, none, "SubI2 ::subw"
    MakeOpcode 0, &HAE, 1, none, "SubI4 ::subd"
    MakeOpcode 0, &HAF, 1, none, "SubR4 ::subf"
    MakeOpcode 0, &HB0, 1, none, "SubCy ::subc"
    MakeOpcode 0, &HB1, 1, none, "MulI2 ::mulw"
    MakeOpcode 0, &HB2, 1, none, "MulI4 ::muld"
    MakeOpcode 0, &HB3, 1, none, "MulR8 ::mulf"
    MakeOpcode 0, &HB4, 1, none, "MulCy ::mulc"
    MakeOpcode 0, &HB5, 1, none, "MulCyI2"
    MakeOpcode 0, &HB6, 1, none, "DivR8 ::divf"
    MakeOpcode 0, &HB7, 1, none, "UMiI2"
    MakeOpcode 0, &HB8, 1, none, "UMiI4"
    MakeOpcode 0, &HB9, 1, none, "UMiR8"
    MakeOpcode 0, &HBA, 1, none, "UMiCy"
    MakeOpcode 0, &HBB, 1, none, "FnAbsI2"
    MakeOpcode 0, &HBC, 1, none, "FnAbsI4"
    MakeOpcode 0, &HBD, 1, none, "FnAbsR4"
    MakeOpcode 0, &HBE, 1, none, "FnAbsCy"
    MakeOpcode 0, &HBF, 1, none, "IDvI2"
    MakeOpcode 0, &HC0, 1, none, "IDvI4"
    MakeOpcode 0, &HC1, 1, none, "ModI2 ::modw"
    MakeOpcode 0, &HC2, 1, none, "ModI4 ::modd"
    MakeOpcode 0, &HC3, 1, none, "NotI4 ::notd"
    MakeOpcode 0, &HC4, 1, none, "AndI4 ::andd"
    MakeOpcode 0, &HC5, 1, none, "OrI4 ::ord"
    MakeOpcode 0, &HC6, 1, none, "EqI2 ::cmpw"
    MakeOpcode 0, &HC7, 1, none, "EqI4 ::cmpd"
    MakeOpcode 0, &HC8, 1, none, "EqR4 ::cmpf"
    MakeOpcode 0, &HC9, 1, none, "EqCy"
    MakeOpcode 0, &HCA, 1, none, "EqCyR8"
    MakeOpcode 0, &HCB, 1, none, "NeI2 ::new"
    MakeOpcode 0, &HCC, 1, none, "NeI4 ::ned"
    MakeOpcode 0, &HCD, 1, none, "NeR8 ::nef"
    MakeOpcode 0, &HCE, 1, none, "NeCy ::nec"
    MakeOpcode 0, &HCF, 1, none, "NeCyR8"
    MakeOpcode 0, &HD0, 1, none, "LtI2 ::ltw"
    MakeOpcode 0, &HD1, 1, none, "LtI4 ::ltd"
    MakeOpcode 0, &HD2, 1, none, "LtR8 ::ltf"
    MakeOpcode 0, &HD3, 1, none, "LtCy ::ltc"
    MakeOpcode 0, &HD4, 1, none, "LtCyR8"
    MakeOpcode 0, &HD5, 1, none, "LeI2 ::lew"
    MakeOpcode 0, &HD6, 1, none, "LeI4 ::led"
    MakeOpcode 0, &HD7, 1, none, "LeR8 ::lef"
    MakeOpcode 0, &HD8, 1, none, "LeCy ::lec"
    MakeOpcode 0, &HD9, 1, none, "LeCyR8"
    MakeOpcode 0, &HDA, 1, none, "GtI2 ::gtw"
    MakeOpcode 0, &HDB, 1, none, "GtI4 ::gtd"
    MakeOpcode 0, &HDC, 1, none, "GtR4 ::gtf"
    MakeOpcode 0, &HDD, 1, none, "GtCy ::gtc"
    MakeOpcode 0, &HDE, 1, none, "GtCyR8"
    MakeOpcode 0, &HDF, 1, none, "GeI2 ::gew"
    MakeOpcode 0, &HE0, 1, none, "GeI4 ::ged"
    MakeOpcode 0, &HE1, 1, none, "GeR8 ::gedf"
    MakeOpcode 0, &HE2, 1, none, "GeCy ::gec"
    MakeOpcode 0, &HE3, 1, none, "GeCyR8"
    MakeOpcode 0, &HE4, 1, none, "CI2I4"
    MakeOpcode 0, &HE5, 1, none, "CI2R8"
    MakeOpcode 0, &HE6, 1, none, "CI2Cy"
    MakeOpcode 0, &HE7, 1, none, "CI4UI1"
    MakeOpcode 0, &HE8, 1, none, "CI4R8"
    MakeOpcode 0, &HE9, 1, none, "CI4Cy"
    MakeOpcode 0, &HEA, 1, none, "CR4R4"
    MakeOpcode 0, &HEB, 1, none, "CR8I2"
    MakeOpcode 0, &HEC, 1, none, "CR8I4"
    MakeOpcode 0, &HED, 1, none, "CR8R8"
    MakeOpcode 0, &HEE, 1, none, "CR8Cy"
    MakeOpcode 0, &HEF, 1, none, "CCyI2"
    MakeOpcode 0, &HF0, 1, none, "CCyI4"
    MakeOpcode 0, &HF1, 1, none, "CCyR4"
    MakeOpcode 0, &HF2, 1, none, "CDateR8"
    MakeOpcode 0, &HF3, 3, std, "LitI2 ::push %2"
    MakeOpcode 0, &HF4, 2, std, "LitI2_Byte ::push %1"
    MakeOpcode 0, &HF5, 5, std, "LitI4 ::push %4"
    MakeOpcode 0, &HF6, 9, 0, "LitCy:"
    MakeOpcode 0, &HF7, 5, 0, "LitCy4:"
    MakeOpcode 0, &HF8, 3, 0, "LitI2FP:"
    MakeOpcode 0, &HF9, 5, 0, "LitR4FP:"
    MakeOpcode 0, &HFA, 9, 0, "LitDate:"
    MakeOpcode 0, &HFB, 1, none, "Lead0"
    MakeOpcode 0, &HFC, 1, none, "Lead1"
    MakeOpcode 0, &HFD, 1, none, "Lead2"
    MakeOpcode 0, &HFE, 1, none, "Lead3"
    MakeOpcode 0, &HFF, 1, none, "Lead4"
End Sub

Sub Init1()
    MakeOpcode 1, &H0, 0, 0, "---"
    MakeOpcode 1, &H1, 1, none, "ImpUI1"
    MakeOpcode 1, &H2, 1, none, "ImpI4"
    MakeOpcode 1, &H3, 1, none, "ImpI4"
    MakeOpcode 1, &H4, 0, 0, "---"
    MakeOpcode 1, &H5, 0, 0, "---"
    MakeOpcode 1, &H6, 0, 0, "---"
    MakeOpcode 1, &H7, 3, 0, "ImpVar"
    MakeOpcode 1, &H8, 0, 0, "---"
    MakeOpcode 1, &H9, 1, none, "EqvUI1"
    MakeOpcode 1, &HA, 1, none, "EqvI4"
    MakeOpcode 1, &HB, 1, none, "EqvI4"
    MakeOpcode 1, &HC, 0, 0, "---"
    MakeOpcode 1, &HD, 0, 0, "---"
    MakeOpcode 1, &HE, 0, 0, "---"
    MakeOpcode 1, &HF, 3, 0, "EqvVar"
    MakeOpcode 1, &H10, 0, 0, "---"
    MakeOpcode 1, &H11, 1, none, "XorI4"
    MakeOpcode 1, &H12, 1, none, "XorI4"
    MakeOpcode 1, &H13, 1, none, "XorI4"
    MakeOpcode 1, &H14, 0, 0, "---"
    MakeOpcode 1, &H15, 0, 0, "---"
    MakeOpcode 1, &H16, 0, 0, "---"
    MakeOpcode 1, &H17, 3, 0, "XorVar"
    MakeOpcode 1, &H18, 0, 0, "---"
    MakeOpcode 1, &H19, 1, none, "OrI2"
    MakeOpcode 1, &H1A, 1, none, "OrI2"
    MakeOpcode 1, &H1B, 1, none, "OrI2"
    MakeOpcode 1, &H1C, 0, 0, "---"
    MakeOpcode 1, &H1D, 0, 0, "---"
    MakeOpcode 1, &H1E, 0, 0, "---"
    MakeOpcode 1, &H1F, 3, 0, "OrVar"
    MakeOpcode 1, &H20, 0, 0, "---"
    MakeOpcode 1, &H21, 1, none, "AndUI1"
    MakeOpcode 1, &H22, 1, none, "AndUI1"
    MakeOpcode 1, &H23, 1, none, "AndUI1"
    MakeOpcode 1, &H24, 0, 0, "---"
    MakeOpcode 1, &H25, 0, 0, "---"
    MakeOpcode 1, &H26, 0, 0, "---"
    MakeOpcode 1, &H27, 3, 0, "AndVar"
    MakeOpcode 1, &H28, 0, 0, "---"
    MakeOpcode 1, &H29, 1, none, "EqI2"
    MakeOpcode 1, &H2A, 1, none, "EqI2"
    MakeOpcode 1, &H2B, 1, none, "EqI4"
    MakeOpcode 1, &H2C, 1, none, "EqR8"
    MakeOpcode 1, &H2D, 1, none, "EqR8"
    MakeOpcode 1, &H2E, 1, none, "EqCy"
    MakeOpcode 1, &H2F, 3, 0, "EqVar"
    MakeOpcode 1, &H30, 1, none, "EqStr"
    MakeOpcode 1, &H31, 3, 0, "EqTextVar"
    MakeOpcode 1, &H32, 1, none, "EqTextStr"
    MakeOpcode 1, &H33, 1, none, "EqVarBool"
    MakeOpcode 1, &H34, 1, none, "EqTextVarBool"
    MakeOpcode 1, &H35, 1, none, "EqCyR8"
    MakeOpcode 1, &H36, 1, none, "NeUI1"
    MakeOpcode 1, &H37, 1, none, "NeUI1"
    MakeOpcode 1, &H38, 1, none, "NeI4"
    MakeOpcode 1, &H39, 1, none, "NeR4"
    MakeOpcode 1, &H3A, 1, none, "NeR4"
    MakeOpcode 1, &H3B, 1, none, "NeCy"
    MakeOpcode 1, &H3C, 3, 0, "NeVar"
    MakeOpcode 1, &H3D, 1, none, "NeStr"
    MakeOpcode 1, &H3E, 3, 0, "NeTextVar"
    MakeOpcode 1, &H3F, 1, none, "NeTextStr"
    MakeOpcode 1, &H40, 1, none, "NeVarBool"
    MakeOpcode 1, &H41, 1, none, "NeTextVarBool"
    MakeOpcode 1, &H42, 1, none, "NeCyR8"
    MakeOpcode 1, &H43, 1, none, "LeUI1"
    MakeOpcode 1, &H44, 1, none, "LeI2"
    MakeOpcode 1, &H45, 1, none, "LeI4"
    MakeOpcode 1, &H46, 1, none, "LeR4"
    MakeOpcode 1, &H47, 1, none, "LeR4"
    MakeOpcode 1, &H48, 1, none, "LeCy"
    MakeOpcode 1, &H49, 3, 0, "LeVar"
    MakeOpcode 1, &H4A, 1, none, "LeStr"
    MakeOpcode 1, &H4B, 3, 0, "LeTextVar"
    MakeOpcode 1, &H4C, 1, none, "LeTextStr"
    MakeOpcode 1, &H4D, 1, none, "LeVarBool"
    MakeOpcode 1, &H4E, 1, none, "LeTextVarBool"
    MakeOpcode 1, &H4F, 1, none, "LeCyR8"
    MakeOpcode 1, &H50, 1, none, "GeUI1"
    MakeOpcode 1, &H51, 1, none, "GeI2"
    MakeOpcode 1, &H52, 1, none, "GeI4"
    MakeOpcode 1, &H53, 1, none, "GeR4"
    MakeOpcode 1, &H54, 1, none, "GeR4"
    MakeOpcode 1, &H55, 1, none, "GeCy"
    MakeOpcode 1, &H56, 3, 0, "GeVar"
    MakeOpcode 1, &H57, 1, none, "GeStr"
    MakeOpcode 1, &H58, 3, 0, "GeTextVar"
    MakeOpcode 1, &H59, 1, none, "GeTextStr"
    MakeOpcode 1, &H5A, 1, none, "GeVarBool"
    MakeOpcode 1, &H5B, 1, none, "GeTextVarBool"
    MakeOpcode 1, &H5C, 1, none, "GeCyR8"
    MakeOpcode 1, &H5D, 1, none, "LtUI1"
    MakeOpcode 1, &H5E, 1, none, "LtI2"
    MakeOpcode 1, &H5F, 1, none, "LtI4"
    MakeOpcode 1, &H60, 1, none, "LtR4"
    MakeOpcode 1, &H61, 1, none, "LtR4"
    MakeOpcode 1, &H62, 1, none, "LtCy"
    MakeOpcode 1, &H63, 3, 0, "LtVar"
    MakeOpcode 1, &H64, 1, none, "LtStr"
    MakeOpcode 1, &H65, 3, 0, "LtTextVar"
    MakeOpcode 1, &H66, 1, none, "LtTextStr"
    MakeOpcode 1, &H67, 1, none, "LtVarBool"
    MakeOpcode 1, &H68, 1, none, "LtTextVarBool"
    MakeOpcode 1, &H69, 1, none, "LtCyR8"
    MakeOpcode 1, &H6A, 1, none, "GtUI1"
    MakeOpcode 1, &H6B, 1, none, "GtI2"
    MakeOpcode 1, &H6C, 1, none, "GtI4"
    MakeOpcode 1, &H6D, 1, none, "GtR4"
    MakeOpcode 1, &H6E, 1, none, "GtR4"
    MakeOpcode 1, &H6F, 1, none, "GtCy"
    MakeOpcode 1, &H70, 3, 0, "GtVar"
    MakeOpcode 1, &H71, 1, none, "GtStr"
    MakeOpcode 1, &H72, 3, 0, "GtTextVar"
    MakeOpcode 1, &H73, 1, none, "GtTextStr"
    MakeOpcode 1, &H74, 1, none, "GtVarBool"
    MakeOpcode 1, &H75, 1, none, "GtTextVarBool"
    MakeOpcode 1, &H76, 1, none, "GtCyR8"
    MakeOpcode 1, &H77, 0, 0, "---"
    MakeOpcode 1, &H78, 0, 0, "---"
    MakeOpcode 1, &H79, 0, 0, "---"
    MakeOpcode 1, &H7A, 0, 0, "---"
    MakeOpcode 1, &H7B, 0, 0, "---"
    MakeOpcode 1, &H7C, 0, 0, "---"
    MakeOpcode 1, &H7D, 3, 0, "LikeVar"
    MakeOpcode 1, &H7E, 1, none, "LikeStr"
    MakeOpcode 1, &H7F, 3, 0, "LikeTextVar"
    MakeOpcode 1, &H80, 1, none, "LikeTextStr"
    MakeOpcode 1, &H81, 1, none, "LikeVarBool"
    MakeOpcode 1, &H82, 1, none, "LikeTextVarBool"
    MakeOpcode 1, &H83, 0, 0, "---"
    MakeOpcode 1, &H84, 1, none, "BetweenUI1"
    MakeOpcode 1, &H85, 1, none, "BetweenI2"
    MakeOpcode 1, &H86, 1, none, "BetweenI4"
    MakeOpcode 1, &H87, 1, none, "BetweenR4"
    MakeOpcode 1, &H88, 1, none, "BetweenR4"
    MakeOpcode 1, &H89, 1, none, "BetweenCy"
    MakeOpcode 1, &H8A, 1, none, "BetweenVar"
    MakeOpcode 1, &H8B, 1, none, "BetweenStr"
    MakeOpcode 1, &H8C, 1, none, "BetweenTextVar"
    MakeOpcode 1, &H8D, 1, none, "BetweenTextStr"
    MakeOpcode 1, &H8E, 1, none, "AddUI1"
    MakeOpcode 1, &H8F, 1, none, "AddI2"
    MakeOpcode 1, &H90, 1, none, "AddI4"
    MakeOpcode 1, &H91, 1, none, "AddR4"
    MakeOpcode 1, &H92, 1, none, "AddR4"
    MakeOpcode 1, &H93, 1, none, "AddCy"
    MakeOpcode 1, &H94, 3, 0, "AddVar"
    MakeOpcode 1, &H95, 0, 0, "---"
    MakeOpcode 1, &H96, 1, none, "SubUI1"
    MakeOpcode 1, &H97, 1, none, "SubI2"
    MakeOpcode 1, &H98, 1, none, "SubI4"
    MakeOpcode 1, &H99, 1, none, "SubR4"
    MakeOpcode 1, &H9A, 1, none, "SubR4"
    MakeOpcode 1, &H9B, 1, none, "SubCy"
    MakeOpcode 1, &H9C, 3, 0, "SubVar"
    MakeOpcode 1, &H9D, 0, 0, "---"
    MakeOpcode 1, &H9E, 1, none, "ModUI1"
    MakeOpcode 1, &H9F, 1, none, "ModI2"
    MakeOpcode 1, &HA0, 1, none, "ModI4"
    MakeOpcode 1, &HA1, 0, 0, "---"
    MakeOpcode 1, &HA2, 0, 0, "---"
    MakeOpcode 1, &HA3, 0, 0, "---"
    MakeOpcode 1, &HA4, 3, 0, "ModVar"
    MakeOpcode 1, &HA5, 0, 0, "---"
    MakeOpcode 1, &HA6, 1, none, "IDvUI1"
    MakeOpcode 1, &HA7, 1, none, "IDvI2"
    MakeOpcode 1, &HA8, 1, none, "IDvI4"
    MakeOpcode 1, &HA9, 0, 0, "---"
    MakeOpcode 1, &HAA, 0, 0, "---"
    MakeOpcode 1, &HAB, 0, 0, "---"
    MakeOpcode 1, &HAC, 3, 0, "IDvVar"
    MakeOpcode 1, &HAD, 0, 0, "Unknow"
    MakeOpcode 1, &HAE, 1, none, "MulUI1"
    MakeOpcode 1, &HAF, 1, none, "MulI2"
    MakeOpcode 1, &HB0, 1, none, "MulI4"
    MakeOpcode 1, &HB1, 1, none, "MulR4"
    MakeOpcode 1, &HB2, 1, none, "MulR4"
    MakeOpcode 1, &HB3, 1, none, "MulCy"
    MakeOpcode 1, &HB4, 3, 0, "MulVar"
    MakeOpcode 1, &HB5, 0, 0, "---"
    MakeOpcode 1, &HB6, 0, 0, "---"
    MakeOpcode 1, &HB7, 0, 0, "---"
    MakeOpcode 1, &HB8, 0, 0, "---"
    MakeOpcode 1, &HB9, 1, none, "DivR8"
    MakeOpcode 1, &HBA, 1, none, "DivR8"
    MakeOpcode 1, &HBB, 0, 0, "---"
    MakeOpcode 1, &HBC, 3, 0, "DivVar"
    MakeOpcode 1, &HBD, 0, 0, "---"
    MakeOpcode 1, &HBE, 1, none, "NotUI1"
    MakeOpcode 1, &HBF, 1, none, "NotI4"
    MakeOpcode 1, &HC0, 1, none, "NotI4"
    MakeOpcode 1, &HC1, 0, 0, "---"
    MakeOpcode 1, &HC2, 0, 0, "---"
    MakeOpcode 1, &HC3, 0, 0, "---"
    MakeOpcode 1, &HC4, 3, 0, "NotVar"
    MakeOpcode 1, &HC5, 0, 0, "---"
    MakeOpcode 1, &HC6, 0, 0, "---"
    MakeOpcode 1, &HC7, 1, none, "UMiI2"
    MakeOpcode 1, &HC8, 1, none, "UMiI2"
    MakeOpcode 1, &HC9, 1, none, "UMiR4"
    MakeOpcode 1, &HCA, 1, none, "UMiR4"
    MakeOpcode 1, &HCB, 1, none, "UMiCy"
    MakeOpcode 1, &HCC, 3, 0, "UMiVar"
    MakeOpcode 1, &HCD, 0, 0, "---"
    MakeOpcode 1, &HCE, 3, 0, "PwrVar"
    MakeOpcode 1, &HCF, 1, none, "PwrR8R8"
    MakeOpcode 1, &HD0, 1, none, "PwrR8I2"
    MakeOpcode 1, &HD1, 1, none, "MulCyI2"
    MakeOpcode 1, &HD2, 1, none, "Is"
    MakeOpcode 1, &HD3, 0, 0, "---"
    MakeOpcode 1, &HD4, 1, none, "FnAbsI2"
    MakeOpcode 1, &HD5, 1, none, "FnAbsI4"
    MakeOpcode 1, &HD6, 1, none, "FnAbsR4"
    MakeOpcode 1, &HD7, 1, none, "FnAbsR4"
    MakeOpcode 1, &HD8, 1, none, "FnAbsCy"
    MakeOpcode 1, &HD9, 3, 0, "FnAbsVar"
    MakeOpcode 1, &HDA, 0, 0, "---"
    MakeOpcode 1, &HDB, 0, 0, "---"
    MakeOpcode 1, &HDC, 0, 0, "---"
    MakeOpcode 1, &HDD, 0, 0, "---"
    MakeOpcode 1, &HDE, 1, none, "FnFixR8"
    MakeOpcode 1, &HDF, 1, none, "FnFixR8"
    MakeOpcode 1, &HE0, 1, none, "FnFixCy"
    MakeOpcode 1, &HE1, 3, 0, "FnFixVar"
    MakeOpcode 1, &HE2, 0, 0, "---"
    MakeOpcode 1, &HE3, 0, 0, "---"
    MakeOpcode 1, &HE4, 0, 0, "---"
    MakeOpcode 1, &HE5, 0, 0, "---"
    MakeOpcode 1, &HE6, 1, none, "FnIntR8"
    MakeOpcode 1, &HE7, 1, none, "FnIntR8"
    MakeOpcode 1, &HE8, 1, none, "FnIntCy"
    MakeOpcode 1, &HE9, 3, 0, "FnIntVar"
    MakeOpcode 1, &HEA, 0, 0, "---"
    MakeOpcode 1, &HEB, 3, 0, "FnLenVar"
    MakeOpcode 1, &HEC, 1, none, "FnLenStr"
    MakeOpcode 1, &HED, 3, 0, "FnLenBVar"
    MakeOpcode 1, &HEE, 1, none, "FnLenBStr"
    MakeOpcode 1, &HEF, 3, 0, "ConcatVar"
    MakeOpcode 1, &HF0, 1, none, "ConcatStr"
    MakeOpcode 1, &HF1, 0, 0, "---"
    MakeOpcode 1, &HF2, 1, none, "FnSgnUI1"
    MakeOpcode 1, &HF3, 1, none, "FnSgnUI1"
    MakeOpcode 1, &HF4, 1, none, "FnSgnI4"
    MakeOpcode 1, &HF5, 1, none, "FnSgnR8"
    MakeOpcode 1, &HF6, 1, none, "FnSgnR4"
    MakeOpcode 1, &HF7, 1, none, "FnSgnCy"
    MakeOpcode 1, &HF8, 0, 0, "---"
    MakeOpcode 1, &HF9, 0, 0, "---"
    MakeOpcode 1, &HFA, 1, none, "SeekFile"
    MakeOpcode 1, &HFB, 1, none, "NameFile"
    MakeOpcode 1, &HFC, 1, none, "CStrI2"
    MakeOpcode 1, &HFD, 1, none, "CStrUI1"
    MakeOpcode 1, &HFE, 1, none, "CStrI4"
    MakeOpcode 1, &HFF, 1, none, "CStrR4"
End Sub

Sub Init2()
    MakeOpcode 2, &H0, 1, none, "CStrR8"
    MakeOpcode 2, &H1, 1, none, "CStrCy"
    MakeOpcode 2, &H2, 1, none, "CStrVar"
    MakeOpcode 2, &H3, 0, 0, "---"
    MakeOpcode 2, &H4, 1, none, "CCyI2"
    MakeOpcode 2, &H5, 1, none, "CCyI2"
    MakeOpcode 2, &H6, 1, none, "CCyI4"
    MakeOpcode 2, &H7, 1, none, "CCyR4"
    MakeOpcode 2, &H8, 1, none, "CCyR4"
    MakeOpcode 2, &H9, 0, 0, "---"
    MakeOpcode 2, &HA, 1, none, "CCyVar"
    MakeOpcode 2, &HB, 1, none, "CCyStr"
    MakeOpcode 2, &HC, 0, 0, "---"
    MakeOpcode 2, &HD, 1, none, "CUI1I2"
    MakeOpcode 2, &HE, 1, none, "CUI1I4"
    MakeOpcode 2, &HF, 1, none, "CUI1R4"
    MakeOpcode 2, &H10, 1, none, "CUI1R4"
    MakeOpcode 2, &H11, 1, none, "CUI1Cy"
    MakeOpcode 2, &H12, 1, none, "CUI1Var"
    MakeOpcode 2, &H13, 1, none, "CUI1Str"
    MakeOpcode 2, &H14, 1, none, "CI2UI1"
    MakeOpcode 2, &H15, 0, 0, "---"
    MakeOpcode 2, &H16, 1, none, "CI2I4"
    MakeOpcode 2, &H17, 1, none, "CI2R8"
    MakeOpcode 2, &H18, 1, none, "CI2R8"
    MakeOpcode 2, &H19, 1, none, "CI2Cy"
    MakeOpcode 2, &H1A, 1, none, "CI2Var"
    MakeOpcode 2, &H1B, 1, none, "CI2Str"
    MakeOpcode 2, &H1C, 1, none, "CI4UI1"
    MakeOpcode 2, &H1D, 1, none, "CI4UI1"
    MakeOpcode 2, &H1E, 0, 0, "---"
    MakeOpcode 2, &H1F, 1, none, "CI4R8"
    MakeOpcode 2, &H20, 1, none, "CI4R8"
    MakeOpcode 2, &H21, 1, none, "CI4Cy"
    MakeOpcode 2, &H22, 1, none, "CI4Var"
    MakeOpcode 2, &H23, 1, none, "CI4Str"
    MakeOpcode 2, &H24, 1, none, "FnCSngI2"
    MakeOpcode 2, &H25, 1, none, "FnCSngI2"
    MakeOpcode 2, &H26, 1, none, "CR4I4"
    MakeOpcode 2, &H27, 1, none, "CR4R8"
    MakeOpcode 2, &H28, 1, none, "CR4R8"
    MakeOpcode 2, &H29, 1, none, "CR8Cy"
    MakeOpcode 2, &H2A, 1, none, "CR8Var"
    MakeOpcode 2, &H2B, 1, none, "CR4Str"
    MakeOpcode 2, &H2C, 1, none, "FnCSngI2"
    MakeOpcode 2, &H2D, 1, none, "FnCSngI2"
    MakeOpcode 2, &H2E, 1, none, "CR4I4"
    MakeOpcode 2, &H2F, 1, none, "CR8R4"
    MakeOpcode 2, &H30, 1, none, "CR8R4"
    MakeOpcode 2, &H31, 1, none, "CR8Cy"
    MakeOpcode 2, &H32, 1, none, "CR8Var"
    MakeOpcode 2, &H33, 1, none, "CR8Str"
    MakeOpcode 2, &H34, 1, none, "CAdVar"
    MakeOpcode 2, &H35, 1, none, "CRefVarAry"
    MakeOpcode 2, &H36, 0, 0, "---"
    MakeOpcode 2, &H37, 0, 0, "---"
    MakeOpcode 2, &H38, 1, none, "CUI1Bool"
    MakeOpcode 2, &H39, 1, none, "FnCDblCy"
    MakeOpcode 2, &H3A, 1, none, "FnCDblR8"
    MakeOpcode 2, &H3B, 1, none, "FnCDblR8"
    MakeOpcode 2, &H3C, 1, none, "FnCSngI2"
    MakeOpcode 2, &H3D, 1, none, "FnCSngI2"
    MakeOpcode 2, &H3E, 1, none, "FnCSngI4"
    MakeOpcode 2, &H3F, 1, none, "CSng"
    MakeOpcode 2, &H40, 1, none, "CSng"
    MakeOpcode 2, &H41, 1, none, "FnCSngCy"
    MakeOpcode 2, &H42, 1, none, "FnCSngVar"
    MakeOpcode 2, &H43, 1, none, "FnCSngStr"
    MakeOpcode 2, &H44, 1, none, "FnCByteVar"
    MakeOpcode 2, &H45, 1, none, "FnCIntVar"
    MakeOpcode 2, &H46, 1, none, "FnCLngVar"
    MakeOpcode 2, &H47, 1, none, "CDateR8"
    MakeOpcode 2, &H48, 1, none, "FnCDblVar"
    MakeOpcode 2, &H49, 1, none, "FnCCurVar"
    MakeOpcode 2, &H4A, 0, 0, "---"
    MakeOpcode 2, &H4B, 1, none, "FnCStrVar"
    MakeOpcode 2, &H4C, 0, 0, "---"
    MakeOpcode 2, &H4D, 1, none, "FnCBoolVar"
    MakeOpcode 2, &H4E, 1, none, "FnCDateVar"
    MakeOpcode 2, &H4F, 1, none, "FnCDateVar"
    MakeOpcode 2, &H50, 1, none, "CBoolUI1"
    MakeOpcode 2, &H51, 1, none, "CBoolUI1"
    MakeOpcode 2, &H52, 1, none, "CBoolI4"
    MakeOpcode 2, &H53, 1, none, "CBoolR4"
    MakeOpcode 2, &H54, 1, none, "CBoolR4"
    MakeOpcode 2, &H55, 1, none, "CBoolCy"
    MakeOpcode 2, &H56, 1, none, "CBoolVar"
    MakeOpcode 2, &H57, 1, none, "CBoolStr"
    MakeOpcode 2, &H58, 1, none, "CStr2Uni"
    MakeOpcode 2, &H59, 1, none, "CStrAry2Uni"
    MakeOpcode 2, &H5A, 1, none, "CStr2Ansi"
    MakeOpcode 2, &H5B, 1, none, "CStrAry2Ansi"
    MakeOpcode 2, &H5C, 1, none, "PopAdLd4"
    MakeOpcode 2, &H5D, 3, 0, "CRecAnsi2Uni"
    MakeOpcode 2, &H5E, 3, 0, "CRecUni2Ansi"
    MakeOpcode 2, &H5F, 3, 0, "CStr2Vec"
    MakeOpcode 2, &H60, 3, 0, "CVar2Vec"
    MakeOpcode 2, &H61, 5, 0, "CVec2Var"
    MakeOpcode 2, &H62, 1, none, "GetLastError"
    MakeOpcode 2, &H63, 1, none, "LitNothing"
    MakeOpcode 2, &H64, 2, 0, "LitVar_Null"
    MakeOpcode 2, &H65, 3, 0, "LitVar_TRUE"
    MakeOpcode 2, &H66, 3, 0, "LitVar_FALSE"
    MakeOpcode 2, &H67, 3, 0, "LitVar_Empty"
    MakeOpcode 2, &H68, 3, 0, "LitVar_Missing"
    MakeOpcode 2, &H69, 5, 0, "VCallHresult"
    MakeOpcode 2, &H6A, 5, 0, "ThisVCallHresult"
    MakeOpcode 2, &H6B, 0, 0, "---"
    MakeOpcode 2, &H6C, 0, 0, "---"
    MakeOpcode 2, &H6D, 1, none, "ExitProcHresult"
    MakeOpcode 2, &H6E, 0, 0, "---"
    MakeOpcode 2, &H6F, 3, 0, "CheckTypeVar"
    MakeOpcode 2, &H70, 0, 0, "---"
    MakeOpcode 2, &H71, 1, none, "CUnkVar"
    MakeOpcode 2, &H72, 3, 0, "CVarUnk"
    MakeOpcode 2, &H73, 1, none, "LdPrUnkVar"
    MakeOpcode 2, &H74, 9, 0, "FLdLateIdUnkVar"
    MakeOpcode 2, &H75, 1, none, "GetRec3"
    MakeOpcode 2, &H76, 1, none, "GetRec4"
    MakeOpcode 2, &H77, 1, none, "PutRec3"
    MakeOpcode 2, &H78, 1, none, "PutRec4"
    MakeOpcode 2, &H79, -1, none, "GetRecOwner3"
    MakeOpcode 2, &H7A, -1, none, "GetRecOwner4"
    MakeOpcode 2, &H7B, -1, none, "PutRecOwner3"
    MakeOpcode 2, &H7C, -1, none, "PutRecOwner4"
    MakeOpcode 2, &H7D, 1, none, "Input"
    MakeOpcode 2, &H7E, 1, none, "InputDone"
    MakeOpcode 2, &H7F, 1, none, "InputItemUI1"
    MakeOpcode 2, &H80, 1, none, "InputItemI2"
    MakeOpcode 2, &H81, 1, none, "InputItemI4"
    MakeOpcode 2, &H82, 1, none, "InputItemR4"
    MakeOpcode 2, &H83, 1, none, "InputItemR8"
    MakeOpcode 2, &H84, 1, none, "InputItemCy"
    MakeOpcode 2, &H85, 1, none, "InputItemVar"
    MakeOpcode 2, &H86, 1, none, "InputItemStr"
    MakeOpcode 2, &H87, 1, none, "InputItemBool"
    MakeOpcode 2, &H88, 1, none, "InputItemDate"
    MakeOpcode 2, &H89, 1, none, "PopFPR4"
    MakeOpcode 2, &H8A, 1, none, "PopFPR8"
    MakeOpcode 2, &H8B, 1, none, "PopAd"
    MakeOpcode 2, &H8C, 1, none, "PopAdLdVar"
    MakeOpcode 2, &H8D, 3, 0, "AryLdPr"
    MakeOpcode 2, &H8E, 3, 0, "AryLdRf"
    MakeOpcode 2, &H8F, 5, 0, "ParmAry1St"
    MakeOpcode 2, &H90, 1, none, "Ary1LdUI1"
    MakeOpcode 2, &H91, 1, none, "Ary1LdI2"
    MakeOpcode 2, &H92, 1, none, "Ary1LdI4"
    MakeOpcode 2, &H93, 1, none, "Ary1LdI4"
    MakeOpcode 2, &H94, 1, none, "Ary1LdR8"
    MakeOpcode 2, &H95, 1, none, "Ary1LdR8"
    MakeOpcode 2, &H96, 1, none, "Ary1LdRfVar"
    MakeOpcode 2, &H97, 1, none, "Ary1LdI4"
    MakeOpcode 2, &H98, 1, none, "Ary1LdI4"
    MakeOpcode 2, &H99, 1, none, "Ary1LdFPR4"
    MakeOpcode 2, &H9A, 1, none, "Ary1LdFPR8"
    MakeOpcode 2, &H9B, 1, none, "Ary1LdPr"
    MakeOpcode 2, &H9C, 1, none, "Ary1LdRf"
    MakeOpcode 2, &H9D, 1, none, "Ary1LdVar"
    MakeOpcode 2, &H9E, 0, 0, "---"
    MakeOpcode 2, &H9F, 0, 0, "---"
    MakeOpcode 2, &HA0, 1, none, "Ary1StUI1"
    MakeOpcode 2, &HA1, 1, none, "Ary1StI2"
    MakeOpcode 2, &HA2, 1, none, "Ary1StR4"
    MakeOpcode 2, &HA3, 1, none, "Ary1StR4"
    MakeOpcode 2, &HA4, 1, none, "Ary1StCy"
    MakeOpcode 2, &HA5, 1, none, "Ary1StCy"
    MakeOpcode 2, &HA6, 1, none, "Ary1StVar"
    MakeOpcode 2, &HA7, 1, none, "Ary1StStr"
    MakeOpcode 2, &HA8, 1, none, "Ary1StAd"
    MakeOpcode 2, &HA9, 1, none, "Ary1StFPR4"
    MakeOpcode 2, &HAA, 1, none, "Ary1StFPR8"
    MakeOpcode 2, &HAB, 1, none, "Ary1StVarAd"
    MakeOpcode 2, &HAC, 1, none, "Ary1StVarAdFunc"
    MakeOpcode 2, &HAD, 1, none, "Ary1StVarUnk"
    MakeOpcode 2, &HAE, 1, none, "Ary1StVarUnkFunc"
    MakeOpcode 2, &HAF, 1, none, "Ary1StAdFunc"
    MakeOpcode 2, &HB0, 1, none, "Ary1StVarCopy"
    MakeOpcode 2, &HB1, 1, none, "Ary1StStrCopy"
    MakeOpcode 2, &HB2, 3, 0, "Ary1LdRfVarg"
    MakeOpcode 2, &HB3, 1, none, "Ary1LdVarg"
    MakeOpcode 2, &HB4, 1, none, "Ary1LdRfVargParam"
    MakeOpcode 2, &HB5, 1, none, "Ary1StVarg"
    MakeOpcode 2, &HB6, 1, none, "Ary1StVargCopy"
    MakeOpcode 2, &HB7, 1, none, "Ary1StVargAd"
    MakeOpcode 2, &HB8, 1, none, "Ary1StVargAdFunc"
    MakeOpcode 2, &HB9, 1, none, "Ary1StVargUnk"
    MakeOpcode 2, &HBA, 1, none, "Ary1StVargUnkFunc"
    MakeOpcode 2, &HBB, 1, none, "MidVar"
    MakeOpcode 2, &HBC, 3, 0, "MidBStr"
    MakeOpcode 2, &HBD, 1, none, "MidBVar"
    MakeOpcode 2, &HBE, 3, 0, "MidBStrB"
    MakeOpcode 2, &HBF, 1, none, "LineInputVar"
    MakeOpcode 2, &HC0, 1, none, "LineInputStr"
    MakeOpcode 2, &HC1, 1, none, "Error"
    MakeOpcode 2, &HC2, 1, none, "Stop"
    MakeOpcode 2, &HC3, 1, none, "Erase"
    MakeOpcode 2, &HC4, 2, 0, "LargeBos"
    MakeOpcode 2, &HC5, 0, 0, "---"
    MakeOpcode 2, &HC6, 0, 0, "---"
    MakeOpcode 2, &HC7, 0, 0, "---"
    MakeOpcode 2, &HC8, 1, none, "End"
    MakeOpcode 2, &HC9, 1, none, "Return"
    MakeOpcode 2, &HCA, 1, none, "FnLBound"
    MakeOpcode 2, &HCB, 1, none, "FnUBound"
    MakeOpcode 2, &HCC, 1, none, "ExitProcUI1"
    MakeOpcode 2, &HCD, 1, none, "ExitProcI2"
    MakeOpcode 2, &HCE, 1, none, "ExitProcStr"
    MakeOpcode 2, &HCF, 1, none, "ExitProcR4"
    MakeOpcode 2, &HD0, 1, none, "ExitProcR8"
    MakeOpcode 2, &HD1, 1, none, "ExitProcCy"
    MakeOpcode 2, &HD2, 0, 0, "---"
    MakeOpcode 2, &HD3, 1, none, "ExitProcStr"
    MakeOpcode 2, &HD4, 1, none, "ExitProcStr"
    MakeOpcode 2, &HD5, 1, none, "ExitProcStr"
    MakeOpcode 2, &HD6, 0, 0, "---"
    MakeOpcode 2, &HD7, 0, 0, "---"
    MakeOpcode 2, &HD8, 0, 0, "---"
    MakeOpcode 2, &HD9, 0, 0, "---"
    MakeOpcode 2, &HDA, 0, 0, "---"
    MakeOpcode 2, &HDB, 0, 0, "---"
    MakeOpcode 2, &HDC, 0, 0, "---"
    MakeOpcode 2, &HDD, 0, 0, "---"
    MakeOpcode 2, &HDE, 0, 0, "---"
    MakeOpcode 2, &HDF, 0, 0, "---"
    MakeOpcode 2, &HE0, 3, 0, "FLdUI1"
    MakeOpcode 2, &HE1, 3, 0, "FLdI2"
    MakeOpcode 2, &HE2, 3, 0, "FLdR4"
    MakeOpcode 2, &HE3, 3, 0, "FLdR4"
    MakeOpcode 2, &HE4, 3, 0, "FLdR8"
    MakeOpcode 2, &HE5, 3, 0, "FLdR8"
    MakeOpcode 2, &HE6, 3, 0, "FLdRfVar"
    MakeOpcode 2, &HE7, 3, 0, "FLdR4"
    MakeOpcode 2, &HE8, 3, 0, "FLdR4"
    MakeOpcode 2, &HE9, 3, 0, "FLdFPR4"
    MakeOpcode 2, &HEA, 3, 0, "FLdFPR8"
    MakeOpcode 2, &HEB, 3, 0, "FLdPr"
    MakeOpcode 2, &HEC, 3, 0, "FLdRfVar"
    MakeOpcode 2, &HED, 3, 0, "FLdVar"
    MakeOpcode 2, &HEE, 0, 0, "---"
    MakeOpcode 2, &HEF, 0, 0, "---"
    MakeOpcode 2, &HF0, 3, 0, "FStUI1"
    MakeOpcode 2, &HF1, 3, 0, "FStI2"
    MakeOpcode 2, &HF2, 3, 0, "FStR4"
    MakeOpcode 2, &HF3, 3, 0, "FStR4"
    MakeOpcode 2, &HF4, 3, 0, "FStR8"
    MakeOpcode 2, &HF5, 3, 0, "FStR8"
    MakeOpcode 2, &HF6, 3, 0, "FStVar"
    MakeOpcode 2, &HF7, 3, 0, "FStStr"
    MakeOpcode 2, &HF8, 3, 0, "FStAd "
    MakeOpcode 2, &HF9, 3, 0, "FStFPR4"
    MakeOpcode 2, &HFA, 3, 0, "FStFPR8"
    MakeOpcode 2, &HFB, 3, 0, "FStVarAd"
    MakeOpcode 2, &HFC, 3, 0, "FStVarAdFunc"
    MakeOpcode 2, &HFD, 3, 0, "FStVarUnk"
    MakeOpcode 2, &HFE, 3, 0, "FStVarUnkFunc"
    MakeOpcode 2, &HFF, 3, 0, "FStAdFunc"
End Sub

Sub Init3()
    MakeOpcode 3, &H0, 3, 0, "FStVarCopy"
    MakeOpcode 3, &H1, 3, 0, "FStStrCopy"
    MakeOpcode 3, &H2, 1, none, "HardType"
    MakeOpcode 3, &H3, 3, 0, "Branch"
    MakeOpcode 3, &H4, 3, 0, "BranchF"
    MakeOpcode 3, &H5, 3, 0, "BranchFVar"
    MakeOpcode 3, &H6, 3, 0, "BranchFVarFree"
    MakeOpcode 3, &H7, 3, 0, "BranchT"
    MakeOpcode 3, &H8, 3, 0, "BranchTVar"
    MakeOpcode 3, &H9, 3, 0, "BranchTVarFree"
    MakeOpcode 3, &HA, 3, 0, "Gosub"
    MakeOpcode 3, &HB, 3, 0, "OnErrorGoto"
    MakeOpcode 3, &HC, 3, 0, "Resume"
    MakeOpcode 3, &HD, 3, 0, "AryLock"
    MakeOpcode 3, &HE, 3, 0, "AryUnlock"
    MakeOpcode 3, &HF, 3, 0, "AryDescTemp"
    MakeOpcode 3, &H10, 3, 0, "ILdUI1"
    MakeOpcode 3, &H11, 3, 0, "ILdI2"
    MakeOpcode 3, &H12, 3, 0, "ILdAd"
    MakeOpcode 3, &H13, 3, 0, "ILdAd"
    MakeOpcode 3, &H14, 3, 0, "ILdR8"
    MakeOpcode 3, &H15, 3, 0, "ILdR8"
    MakeOpcode 3, &H16, 5, 0, "ILdRfDarg"
    MakeOpcode 3, &H17, 3, 0, "ILdAd"
    MakeOpcode 3, &H18, 3, 0, "ILdAd"
    MakeOpcode 3, &H19, 3, 0, "ILdFPR4"
    MakeOpcode 3, &H1A, 3, 0, "ILdFPR8"
    MakeOpcode 3, &H1B, 3, 0, "ILdPr"
    MakeOpcode 3, &H1C, 3, 0, "FLdR4"
    MakeOpcode 3, &H1D, 3, 0, "ILdDarg"
    MakeOpcode 3, &H1E, 0, 0, "---"
    MakeOpcode 3, &H1F, 0, 0, "---"
    MakeOpcode 3, &H20, 3, 0, "IStUI1"
    MakeOpcode 3, &H21, 3, 0, "IStI2"
    MakeOpcode 3, &H22, 3, 0, "IStI4"
    MakeOpcode 3, &H23, 3, 0, "IStI4"
    MakeOpcode 3, &H24, 3, 0, "IStR8"
    MakeOpcode 3, &H25, 3, 0, "IStR8"
    MakeOpcode 3, &H26, 3, 0, "IStDarg"
    MakeOpcode 3, &H27, 3, 0, "IStStr"
    MakeOpcode 3, &H28, 3, 0, "IStAd"
    MakeOpcode 3, &H29, 3, 0, "IStFPR4"
    MakeOpcode 3, &H2A, 3, 0, "IStFPR8"
    MakeOpcode 3, &H2B, 3, 0, "IStDargAd"
    MakeOpcode 3, &H2C, 3, 0, "IStDargAdFunc"
    MakeOpcode 3, &H2D, 3, 0, "IStDargUnk"
    MakeOpcode 3, &H2E, 3, 0, "IStDargUnkFunc"
    MakeOpcode 3, &H2F, 3, 0, "IStAdFunc"
    MakeOpcode 3, &H30, 3, 0, "IStDargCopy"
    MakeOpcode 3, &H31, 3, 0, "IStStrCopy"
    MakeOpcode 3, &H32, 1, none, "PrintChan"
    MakeOpcode 3, &H33, 1, none, "WriteChan"
    MakeOpcode 3, &H34, 1, none, "PrintComma"
    MakeOpcode 3, &H35, 1, none, "PrintEos"
    MakeOpcode 3, &H36, 1, none, "PrintNL"
    MakeOpcode 3, &H37, 1, none, "PrintItemComma"
    MakeOpcode 3, &H38, 1, none, "PrintItemSemi"
    MakeOpcode 3, &H39, 1, none, "PrintItemNL"
    MakeOpcode 3, &H3A, 3, 0, "PrintObj"
    MakeOpcode 3, &H3B, 1, none, "PrintSpc"
    MakeOpcode 3, &H3C, 1, none, "PrintTab"
    MakeOpcode 3, &H3D, 1, none, "Close"
    MakeOpcode 3, &H3E, 1, none, "CloseAll"
    MakeOpcode 3, &H3F, 3, 0, "FLdZeroAd"
    MakeOpcode 3, &H40, 3, 0, "IWMemLdUI1"
    MakeOpcode 3, &H41, 3, 0, "IWMemLdI2"
    MakeOpcode 3, &H42, 3, 0, "IWMemLdI4"
    MakeOpcode 3, &H43, 3, 0, "IWMemLdI4"
    MakeOpcode 3, &H44, 3, 0, "IWMemLdCy"
    MakeOpcode 3, &H45, 3, 0, "IWMemLdCy"
    MakeOpcode 3, &H46, 5, 0, "IWMemLdRfDarg"
    MakeOpcode 3, &H47, 3, 0, "IWMemLdI4"
    MakeOpcode 3, &H48, 3, 0, "IWMemLdI4"
    MakeOpcode 3, &H49, 3, 0, "IWMemLdFPR4"
    MakeOpcode 3, &H4A, 3, 0, "IWMemLdFPR8"
    MakeOpcode 3, &H4B, 3, 0, "IWMemLdPr"
    MakeOpcode 3, &H4C, 3, 0, "IWMemLdRf"
    MakeOpcode 3, &H4D, 3, 0, "IWMemLdDarg"
    MakeOpcode 3, &H4E, 0, 0, "---"
    MakeOpcode 3, &H4F, 0, 0, "---"
    MakeOpcode 3, &H50, 3, 0, "IWMemStUI1"
    MakeOpcode 3, &H51, 3, 0, "IWMemStI2"
    MakeOpcode 3, &H52, 3, 0, "IWMemStR4"
    MakeOpcode 3, &H53, 3, 0, "IWMemStR4"
    MakeOpcode 3, &H54, 3, 0, "IWMemStCy"
    MakeOpcode 3, &H55, 3, 0, "IWMemStCy"
    MakeOpcode 3, &H56, 3, 0, "IWMemStDarg"
    MakeOpcode 3, &H57, 3, 0, "IWMemStStr"
    MakeOpcode 3, &H58, 3, 0, "IWMemStAd"
    MakeOpcode 3, &H59, 3, 0, "IWMemStFPR4"
    MakeOpcode 3, &H5A, 3, 0, "IWMemStFPR8"
    MakeOpcode 3, &H5B, 3, 0, "IWMemStDargAd"
    MakeOpcode 3, &H5C, 3, 0, "IWMemStDargAdFunc"
    MakeOpcode 3, &H5D, 3, 0, "IWMemStDargUnk"
    MakeOpcode 3, &H5E, 3, 0, "IWMemStDargUnkFunc"
    MakeOpcode 3, &H5F, 3, 0, "IWMemStAdFunc"
    MakeOpcode 3, &H60, 3, 0, "IWMemStDargCopy"
    MakeOpcode 3, &H61, 3, 0, "IWMemStStrCopy"
    MakeOpcode 3, &H62, 3, 0, "FLdZeroAd"
    MakeOpcode 3, &H63, 3, 0, "FStVarNoPop"
    MakeOpcode 3, &H64, 3, 0, "FStStrNoPop"
    MakeOpcode 3, &H65, 0, 0, "---"
    MakeOpcode 3, &H66, 0, 0, "---"
    MakeOpcode 3, &H67, 3, 0, "CVarUI1"
    MakeOpcode 3, &H68, 3, 0, "CVarI2"
    MakeOpcode 3, &H69, 3, 0, "CVarI4"
    MakeOpcode 3, &H6A, 3, 0, "CVarR4"
    MakeOpcode 3, &H6B, 3, 0, "CVarR8"
    MakeOpcode 3, &H6C, 3, 0, "CVarCy"
    MakeOpcode 3, &H6D, 0, 0, "---"
    MakeOpcode 3, &H6E, 3, 0, "CVarStr"
    MakeOpcode 3, &H6F, 3, 0, "CVarAd"
    MakeOpcode 3, &H70, 3, 0, "MemLdUI1"
    MakeOpcode 3, &H71, 3, 0, "MemLdI2"
    MakeOpcode 3, &H72, 3, 0, "MemLdR4"
    MakeOpcode 3, &H73, 3, 0, "MemLdR4"
    MakeOpcode 3, &H74, 3, 0, "MemLdR8"
    MakeOpcode 3, &H75, 3, 0, "MemLdR8"
    MakeOpcode 3, &H76, 3, 0, "MemLdRfVar"
    MakeOpcode 3, &H77, 3, 0, "MemLdR4"
    MakeOpcode 3, &H78, 3, 0, "MemLdR4"
    MakeOpcode 3, &H79, 3, 0, "MemLdFPR4"
    MakeOpcode 3, &H7A, 3, 0, "MemLdFPR8"
    MakeOpcode 3, &H7B, 3, 0, "MemLdPr"
    MakeOpcode 3, &H7C, 3, 0, "MemLdRfVar"
    MakeOpcode 3, &H7D, 3, 0, "MemLdVar"
    MakeOpcode 3, &H7E, 0, 0, "---"
    MakeOpcode 3, &H7F, 0, 0, "---"
    MakeOpcode 3, &H80, 3, 0, "MemStUI1"
    MakeOpcode 3, &H81, 3, 0, "MemStI2"
    MakeOpcode 3, &H82, 3, 0, "MemStR4"
    MakeOpcode 3, &H83, 3, 0, "MemStR4"
    MakeOpcode 3, &H84, 3, 0, "MemStCy"
    MakeOpcode 3, &H85, 3, 0, "MemStCy"
    MakeOpcode 3, &H86, 3, 0, "MemStVar"
    MakeOpcode 3, &H87, 3, 0, "MemStStr"
    MakeOpcode 3, &H88, 3, 0, "MemStAd"
    MakeOpcode 3, &H89, 3, 0, "MemStFPR4"
    MakeOpcode 3, &H8A, 3, 0, "MemStFPR8"
    MakeOpcode 3, &H8B, 3, 0, "MemStVarAd"
    MakeOpcode 3, &H8C, 3, 0, "MemStVarAdFunc"
    MakeOpcode 3, &H8D, 3, 0, "MemStVarUnk"
    MakeOpcode 3, &H8E, 3, 0, "MemStVarUnkFunc"
    MakeOpcode 3, &H8F, 3, 0, "MemStAdFunc"
    MakeOpcode 3, &H90, 3, 0, "MemStVarCopy"
    MakeOpcode 3, &H91, 3, 0, "MemStStrCopy"
    MakeOpcode 3, &H92, 0, 0, "---"
    MakeOpcode 3, &H93, 3, 0, "CDargRef"
    MakeOpcode 3, &H94, 5, 0, "CVarRef"
    MakeOpcode 3, &H95, 3, 0, "ExitProcCb"
    MakeOpcode 3, &H96, 3, 0, "ExitProcCbStack"
    MakeOpcode 3, &H97, 0, 0, "---"
    MakeOpcode 3, &H98, 0, 0, "---"
    MakeOpcode 3, &H99, 3, 0, "FFree1Var"
    MakeOpcode 3, &H9A, 3, 0, "FFree1Str"
    MakeOpcode 3, &H9B, 3, 0, "FFree1Ad"
    MakeOpcode 3, &H9C, 3, 0, "FStAdNoPop"
    MakeOpcode 3, &H9D, 3, 0, "FStAdFuncNoPop"
    MakeOpcode 3, &H9E, 1, none, "FLdPrThis"
    MakeOpcode 3, &H9F, 1, none, "LdPrVar"
    MakeOpcode 3, &HA0, 3, 0, "ImpAdLdUI1"
    MakeOpcode 3, &HA1, 3, 0, "ImpAdLdI2"
    MakeOpcode 3, &HA2, 3, 0, "ImpAdLdStr"
    MakeOpcode 3, &HA3, 3, 0, "ImpAdLdStr"
    MakeOpcode 3, &HA4, 3, 0, "ImpAdLdCy"
    MakeOpcode 3, &HA5, 3, 0, "ImpAdLdCy"
    MakeOpcode 3, &HA6, 3, 0, "ImpAdLdRf"
    MakeOpcode 3, &HA7, 3, 0, "ImpAdLdStr"
    MakeOpcode 3, &HA8, 3, 0, "ImpAdLdStr"
    MakeOpcode 3, &HA9, 3, 0, "ImpAdLdFPR4"
    MakeOpcode 3, &HAA, 3, 0, "ImpAdLdFPR8"
    MakeOpcode 3, &HAB, 3, 0, "ImpAdLdPr"
    MakeOpcode 3, &HAC, 3, 0, "ImpAdLdRf"
    MakeOpcode 3, &HAD, 3, 0, "ImpAdLdVar"
    MakeOpcode 3, &HAE, 0, 0, "---"
    MakeOpcode 3, &HAF, 0, 0, "---"
    MakeOpcode 3, &HB0, 3, 0, "ImpAdStUI1"
    MakeOpcode 3, &HB1, 3, 0, "ImpAdStI2"
    MakeOpcode 3, &HB2, 3, 0, "ImpAdStR4"
    MakeOpcode 3, &HB3, 3, 0, "ImpAdStR4"
    MakeOpcode 3, &HB4, 3, 0, "ImpAdStR8"
    MakeOpcode 3, &HB5, 3, 0, "ImpAdStR8"
    MakeOpcode 3, &HB6, 3, 0, "ImpAdStVar"
    MakeOpcode 3, &HB7, 3, 0, "ImpAdStStr"
    MakeOpcode 3, &HB8, 3, 0, "ImpAdStAd"
    MakeOpcode 3, &HB9, 3, 0, "ImpAdStFPR4"
    MakeOpcode 3, &HBA, 3, 0, "ImpAdStFPR8"
    MakeOpcode 3, &HBB, 3, 0, "ImpAdStVarAd"
    MakeOpcode 3, &HBC, 3, 0, "ImpAdStVarAdFunc"
    MakeOpcode 3, &HBD, 3, 0, "ImpAdStVarUnk"
    MakeOpcode 3, &HBE, 3, 0, "ImpAdStVarUnkFunc"
    MakeOpcode 3, &HBF, 3, 0, "ImpAdStAdFunc"
    MakeOpcode 3, &HC0, 3, 0, "ImpAdStVarCopy"
    MakeOpcode 3, &HC1, 3, 0, "ImpAdStStrCopy"
    MakeOpcode 3, &HC2, 3, 0, "PopTmpLdAd1"
    MakeOpcode 3, &HC3, 3, 0, "PopTmpLdAd2"
    MakeOpcode 3, &HC4, 3, 0, "PopTmpLdAdStr"
    MakeOpcode 3, &HC5, 3, 0, "PopTmpLdAd8"
    MakeOpcode 3, &HC6, 3, 0, "PopTmpLdAdVar"
    MakeOpcode 3, &HC7, 3, 0, "PopTmpLdAdStr"
    MakeOpcode 3, &HC8, 3, 0, "PopTmpLdAdFPR4"
    MakeOpcode 3, &HC9, 3, 0, "PopTmpLdAdFPR8"
    MakeOpcode 3, &HCA, 3, 0, "CopyBytes"
    MakeOpcode 3, &HCB, 1, none, "ExitForCollObj"
    MakeOpcode 3, &HCC, 1, none, "ExitForCollObj"
    MakeOpcode 3, &HCD, 1, none, "ExitForCollObj"
    MakeOpcode 3, &HCE, 1, none, "ExitForAryVar"
    MakeOpcode 3, &HCF, 1, none, "ExitForVar"
    MakeOpcode 3, &HD0, 5, 0, "FMemLdUI1"
    MakeOpcode 3, &HD1, 5, 0, "FMemLdI2"
    MakeOpcode 3, &HD2, 5, 0, "FMemLdR4"
    MakeOpcode 3, &HD3, 5, 0, "FMemLdR4"
    MakeOpcode 3, &HD4, 5, 0, "FMemLdR8"
    MakeOpcode 3, &HD5, 5, 0, "FMemLdR8"
    MakeOpcode 3, &HD6, 5, 0, "FMemLdRf"
    MakeOpcode 3, &HD7, 5, 0, "FMemLdR4"
    MakeOpcode 3, &HD8, 5, 0, "FMemLdR4"
    MakeOpcode 3, &HD9, 5, 0, "FMemLdFPR4"
    MakeOpcode 3, &HDA, 5, 0, "FMemLdFPR8"
    MakeOpcode 3, &HDB, 5, 0, "FMemLdPr"
    MakeOpcode 3, &HDC, 5, 0, "FMemLdRf"
    MakeOpcode 3, &HDD, 5, 0, "FMemLdVar"
    MakeOpcode 3, &HDE, 0, 0, "---"
    MakeOpcode 3, &HDF, 0, 0, "---"
    MakeOpcode 3, &HE0, 5, 0, "FMemStUI1"
    MakeOpcode 3, &HE1, 5, 0, "FMemStI2"
    MakeOpcode 3, &HE2, 5, 0, "FMemStR4"
    MakeOpcode 3, &HE3, 5, 0, "FMemStR4"
    MakeOpcode 3, &HE4, 5, 0, "FMemStR8"
    MakeOpcode 3, &HE5, 5, 0, "FMemStR8"
    MakeOpcode 3, &HE6, 5, 0, "FMemStVar"
    MakeOpcode 3, &HE7, 5, 0, "FMemStStr"
    MakeOpcode 3, &HE8, 5, 0, "FMemStAd"
    MakeOpcode 3, &HE9, 5, 0, "FMemStFPR4"
    MakeOpcode 3, &HEA, 5, 0, "FMemStFPR8"
    MakeOpcode 3, &HEB, 5, 0, "FMemStVarAd"
    MakeOpcode 3, &HEC, 5, 0, "FMemStVarAdFunc"
    MakeOpcode 3, &HED, 5, 0, "FMemStVarUnk"
    MakeOpcode 3, &HEE, 5, 0, "FMemStVarUnkFunc"
    MakeOpcode 3, &HEF, 5, 0, "FMemStAdFunc"
    MakeOpcode 3, &HF0, 5, 0, "FMemStVarCopy"
    MakeOpcode 3, &HF1, 5, 0, "FMemStStrCopy"
    MakeOpcode 3, &HF2, 3, 0, "CastAd"
    MakeOpcode 3, &HF3, 3, 0, "CastAdVar"
    MakeOpcode 3, &HF4, 3, 0, "New"
    MakeOpcode 3, &HF5, 3, 0, "NewIfNullRf"
    MakeOpcode 3, &HF6, 3, 0, "NewIfNullAd"
    MakeOpcode 3, &HF7, 3, 0, "NewIfNullPr"
    MakeOpcode 3, &HF8, 3, 0, "CVarBoolI2"
    MakeOpcode 3, &HF9, 3, 0, "CVarDateVar"
    MakeOpcode 3, &HFA, 3, 0, "CVarErrI4"
    MakeOpcode 3, &HFB, 3, 0, "CVarDate"
    MakeOpcode 3, &HFC, 3, 0, "CVarAryVarg"
    MakeOpcode 3, &HFD, 1, none, "CStrVarTmp"
    MakeOpcode 3, &HFE, 3, 0, "CStrVarVal"
    MakeOpcode 3, &HFF, 5, 0, "DestructOFrame"
End Sub

Sub Init4()
    MakeOpcode 4, &H0, 3, 0, "ThisVCallUI1"
    MakeOpcode 4, &H1, 3, 0, "ThisVCallI2"
    MakeOpcode 4, &H2, 3, 0, "ThisVCallI2"
    MakeOpcode 4, &H3, 3, 0, "ThisVCallR4"
    MakeOpcode 4, &H4, 3, 0, "ThisVCallR8"
    MakeOpcode 4, &H5, 3, 0, "ThisVCallCy"
    MakeOpcode 4, &H6, 0, 0, "---"
    MakeOpcode 4, &H7, 3, 0, "ThisVCallI2"
    MakeOpcode 4, &H8, 3, 0, "ThisVCallI2"
    MakeOpcode 4, &H9, 3, 0, "ThisVCallHidden"
    MakeOpcode 4, &HA, 3, 0, "ThisVCallHidden"
    MakeOpcode 4, &HB, 0, 0, "---"
    MakeOpcode 4, &HC, 3, 0, "ThisVCallHidden"
    MakeOpcode 4, &HD, 5, 0, "ThisVCallCbFrame"
    MakeOpcode 4, &HE, 3, 0, "StLsetFixStr"
    MakeOpcode 4, &HF, 3, 0, "StFixedStrFree"
    MakeOpcode 4, &H10, 3, 0, "VCallUI1"
    MakeOpcode 4, &H11, 3, 0, "VCallStr"
    MakeOpcode 4, &H12, 3, 0, "VCallStr"
    MakeOpcode 4, &H13, 3, 0, "VCallR4"
    MakeOpcode 4, &H14, 3, 0, "VCallR8"
    MakeOpcode 4, &H15, 3, 0, "VCallCy"
    MakeOpcode 4, &H16, 0, 0, "---"
    MakeOpcode 4, &H17, 3, 0, "VCallStr"
    MakeOpcode 4, &H18, 3, 0, "VCallStr"
    MakeOpcode 4, &H19, 3, 0, "VCallFPR8"
    MakeOpcode 4, &H1A, 3, 0, "VCallFPR8"
    MakeOpcode 4, &H1B, 0, 0, "---"
    MakeOpcode 4, &H1C, 3, 0, "VCallFPR8"
    MakeOpcode 4, &H1D, 5, 0, "VCallCbFrame"
    MakeOpcode 4, &H1E, 3, 0, "StFixedStrR"
    MakeOpcode 4, &H1F, 3, 0, "StFixedStrRFree"
    MakeOpcode 4, &H20, 5, 0, "ImpAdCallUI1"
    MakeOpcode 4, &H21, 5, 0, "ImpAdCallI4"
    MakeOpcode 4, &H22, 5, 0, "ImpAdCallI4"
    MakeOpcode 4, &H23, 5, 0, "ImpAdCallR4"
    MakeOpcode 4, &H24, 5, 0, "ImpAdCallR8"
    MakeOpcode 4, &H25, 5, 0, "ImpAdCallCy"
    MakeOpcode 4, &H26, 0, 0, "---"
    MakeOpcode 4, &H27, 5, 0, "ImpAdCallI4"
    MakeOpcode 4, &H28, 5, 0, "ImpAdCallI4"
    MakeOpcode 4, &H29, 5, 0, "ImpAdCallFPR4"
    MakeOpcode 4, &H2A, 5, 0, "ImpAdCallFPR4"
    MakeOpcode 4, &H2B, 0, 0, "---"
    MakeOpcode 4, &H2C, 5, 0, "ImpAdCallFPR4"
    MakeOpcode 4, &H2D, 9, 0, "ImpAdCallCbFrame"
    MakeOpcode 4, &H2E, 3, 0, "LdStkRf"
    MakeOpcode 4, &H2F, 3, 0, "LdFrameRf"
    MakeOpcode 4, &H30, 0, 0, "---"
    MakeOpcode 4, &H31, 0, 0, "---"
    MakeOpcode 4, &H32, 0, 0, "---"
    MakeOpcode 4, &H33, 0, 0, "---"
    MakeOpcode 4, &H34, 0, 0, "---"
    MakeOpcode 4, &H35, 0, 0, "---"
    MakeOpcode 4, &H36, 0, 0, "---"
    MakeOpcode 4, &H37, 0, 0, "---"
    MakeOpcode 4, &H38, 0, 0, "---"
    MakeOpcode 4, &H39, 0, 0, "---"
    MakeOpcode 4, &H3A, 0, 0, "---"
    MakeOpcode 4, &H3B, 0, 0, "---"
    MakeOpcode 4, &H3C, 0, 0, "---"
    MakeOpcode 4, &H3D, 5, 0, "LitVarUI1"
    MakeOpcode 4, &H3E, 0, 0, "---"
    MakeOpcode 4, &H3F, 0, 0, "---"
    MakeOpcode 4, &H40, 0, 0, "---"
    MakeOpcode 4, &H41, 0, 0, "---"
    MakeOpcode 4, &H42, 0, 0, "---"
    MakeOpcode 4, &H43, 0, 0, "---"
    MakeOpcode 4, &H44, 0, 0, "---"
    MakeOpcode 4, &H45, 0, 0, "---"
    MakeOpcode 4, &H46, 0, 0, "---"
    MakeOpcode 4, &H47, 0, 0, "---"
    MakeOpcode 4, &H48, 0, 0, "---"
    MakeOpcode 4, &H49, 0, 0, "---"
    MakeOpcode 4, &H4A, 0, 0, "---"
    MakeOpcode 4, &H4B, 0, 0, "---"
    MakeOpcode 4, &H4C, 0, 0, "---"
    MakeOpcode 4, &H4D, 1, none, "SetVarVar"
    MakeOpcode 4, &H4E, 1, none, "SetVarVarFunc"
    MakeOpcode 4, &H4F, 5, 0, "ImpAdCallHresult"
    MakeOpcode 4, &H50, 0, 0, "---"
    MakeOpcode 4, &H51, 0, 0, "---"
    MakeOpcode 4, &H52, 0, 0, "---"
    MakeOpcode 4, &H53, 0, 0, "---"
    MakeOpcode 4, &H54, 0, 0, "---"
    MakeOpcode 4, &H55, 0, 0, "---"
    MakeOpcode 4, &H56, 0, 0, "---"
    MakeOpcode 4, &H57, 0, 0, "---"
    MakeOpcode 4, &H58, 0, 0, "---"
    MakeOpcode 4, &H59, 0, 0, "---"
    MakeOpcode 4, &H5A, 0, 0, "---"
    MakeOpcode 4, &H5B, 0, 0, "---"
    MakeOpcode 4, &H5C, 0, 0, "---"
    MakeOpcode 4, &H5D, 3, 0, "OpenFile"
    MakeOpcode 4, &H5E, 3, 0, "LockFile"
    MakeOpcode 4, &H5F, 0, 0, "---"
    MakeOpcode 4, &H60, 3, 0, "EraseDestruct"
    MakeOpcode 4, &H61, 3, 0, "LdFixedStr"
    MakeOpcode 4, &H62, 5, 0, "ForUI1"
    MakeOpcode 4, &H63, 5, 0, "ForI2:"
    MakeOpcode 4, &H64, 5, 0, "ForI4:"
    MakeOpcode 4, &H65, 5, 0, "ForR4"
    MakeOpcode 4, &H66, 5, 0, "ForR8"
    MakeOpcode 4, &H67, 5, 0, "ForCy"
    MakeOpcode 4, &H68, 5, 0, "ForVar"
    MakeOpcode 4, &H69, 0, 0, "---"
    MakeOpcode 4, &H6A, 5, 0, "ForStepUI1"
    MakeOpcode 4, &H6B, 5, 0, "ForStepI2"
    MakeOpcode 4, &H6C, 5, 0, "ForStepI4"
    MakeOpcode 4, &H6D, 5, 0, "ForStepR4"
    MakeOpcode 4, &H6E, 5, 0, "ForStepR8"
    MakeOpcode 4, &H6F, 5, 0, "ForStepCy"
    MakeOpcode 4, &H70, 5, 0, "ForStepVar"
    MakeOpcode 4, &H71, 0, 0, "---"
    MakeOpcode 4, &H72, 5, 0, "ForEachCollVar"
    MakeOpcode 4, &H73, 5, 0, "NextEachCollVar"
    MakeOpcode 4, &H74, 5, 0, "ForEachCollAd"
    MakeOpcode 4, &H75, 5, 0, "NextEachCollAd"
    MakeOpcode 4, &H76, 7, 0, "ForEachAryVar"
    MakeOpcode 4, &H77, 7, 0, "NextEachAryVar"
    MakeOpcode 4, &H78, 5, 0, "NextUI1"
    MakeOpcode 4, &H79, 5, 0, "NextI2"
    MakeOpcode 4, &H7A, 5, 0, "NextI4"
    MakeOpcode 4, &H7B, 5, 0, "NextStepR4"
    MakeOpcode 4, &H7C, 5, 0, "NextR8"
    MakeOpcode 4, &H7D, 5, 0, "NextStepCy"
    MakeOpcode 4, &H7E, 5, 0, "NextStepVar"
    MakeOpcode 4, &H7F, 0, 0, "---"
    MakeOpcode 4, &H80, 5, 0, "NextStepUI1"
    MakeOpcode 4, &H81, 5, 0, "NextStepI2"
    MakeOpcode 4, &H82, 5, 0, "NextStepI4"
    MakeOpcode 4, &H83, 5, 0, "NextStepR4"
    MakeOpcode 4, &H84, 5, 0, "NextR8"
    MakeOpcode 4, &H85, 5, 0, "NextStepCy"
    MakeOpcode 4, &H86, 5, 0, "NextStepVar"
    MakeOpcode 4, &H87, 0, 0, "---"
    MakeOpcode 4, &H88, 7, 0, "ForEachCollObj"
    MakeOpcode 4, &H89, 5, 0, "ForEachVar"
    MakeOpcode 4, &H8A, 5, 0, "ForEachVarFree"
    MakeOpcode 4, &H8B, 11, none, "NextEachCollObj"
    MakeOpcode 4, &H8C, 9, 0, "NextEachVar"
    MakeOpcode 4, &H8D, 3, 0, "CheckType"
    MakeOpcode 4, &H8E, 9, 0, "Redim"
    MakeOpcode 4, &H8F, 9, 0, "RedimPreserve"
    MakeOpcode 4, &H90, 5, 0, "RedimVar"
    MakeOpcode 4, &H91, 5, 0, "RedimPreserveVar"
    MakeOpcode 4, &H92, 5, 0, "FDupVar"
    MakeOpcode 4, &H93, 5, 0, "FDupStr"
    MakeOpcode 4, &H94, 0, 0, "---"
    MakeOpcode 4, &H95, 7, 0, "OnGosub"
    MakeOpcode 4, &H96, 7, 0, "OnGoto"
    MakeOpcode 4, &H97, 1, none, "AddRef"
    MakeOpcode 4, &H98, 5, 0, "LateMemCall"
    MakeOpcode 4, &H99, 5, 0, "LateMemLdVar"
    MakeOpcode 4, &H9A, 7, 0, "LateMemCallLdVar"
    MakeOpcode 4, &H9B, 3, 0, "LateMemSt"
    MakeOpcode 4, &H9C, 5, 0, "LateMemCallSt"
    MakeOpcode 4, &H9D, 5, 0, "LateMemStAd"
    MakeOpcode 4, &H9E, 4, 0, "ExitProcFrameCb"
    MakeOpcode 4, &H9F, 4, 0, "ExitProcFrameCbStack"
    MakeOpcode 4, &HA0, 7, 0, "LateIdCall"
    MakeOpcode 4, &HA1, 7, 0, "LateIdLdVar"
    MakeOpcode 4, &HA2, 9, 0, "LateIdCallLdVar"
    MakeOpcode 4, &HA3, 5, 0, "LateIdSt"
    MakeOpcode 4, &HA4, 7, 0, "LateIdCallSt"
    MakeOpcode 4, &HA5, 7, 0, "LateIdStAd"
    MakeOpcode 4, &HA6, 7, 0, "LateMemNamedCall"
    MakeOpcode 4, &HA7, 9, 0, "LateMemNamedCallLdVar"
    MakeOpcode 4, &HA8, 7, 0, "LateMemNamedCallSt"
    MakeOpcode 4, &HA9, 7, 0, "LateMemNamedStAd"
    MakeOpcode 4, &HAA, 9, 0, "LateIdNamedCall"
    MakeOpcode 4, &HAB, 11, none, "LateIdNamedCallLdVar"
    MakeOpcode 4, &HAC, 9, 0, "LateIdNamedCallSt"
    MakeOpcode 4, &HAD, 9, 0, "LateIdNamedStAd"
    MakeOpcode 4, &HAE, 5, 0, "VarIndexLdVar"
    MakeOpcode 4, &HAF, 5, 0, "VarIndexLdRfVar"
    MakeOpcode 4, &HB0, 3, 0, "VarIndexSt"
    MakeOpcode 4, &HB1, 3, 0, "VarIndexStAd"
    MakeOpcode 4, &HB2, -1, none, "FFreeVar"
    MakeOpcode 4, &HB3, -1, none, "FFreeStr"
    MakeOpcode 4, &HB4, -1, none, "FFreeAd"
    MakeOpcode 4, &HB5, 3, 0, "LitI2"
    MakeOpcode 4, &HB6, 3, 0, "LitI2FP"
    MakeOpcode 4, &HB7, 5, 0, "LitCy4"
    MakeOpcode 4, &HB8, 5, 0, "LitI4"
    MakeOpcode 4, &HB9, 5, 0, "LitI4"
    MakeOpcode 4, &HBA, 5, 0, "LitR4FP"
    MakeOpcode 4, &HBB, 9, 0, "LitR8"
    MakeOpcode 4, &HBC, 9, 0, "LitR8"
    MakeOpcode 4, &HBD, 9, 0, "LitR8FP"
    MakeOpcode 4, &HBE, 9, 0, "LitR8FP"
    MakeOpcode 4, &HBF, 3, 0, "LitStr"
    MakeOpcode 4, &HC0, 5, 0, "LitVarI2"
    MakeOpcode 4, &HC1, 7, 0, "LitVarI4"
    MakeOpcode 4, &HC2, 7, 0, "LitVarR4"
    MakeOpcode 4, &HC3, 11, none, "LitVarCy"
    MakeOpcode 4, &HC4, 11, none, "LitVarR8"
    MakeOpcode 4, &HC5, 11, none, "LitVarDate"
    MakeOpcode 4, &HC6, 5, 0, "LitVarStr"
    MakeOpcode 4, &HC7, 1, none, "CStrBool"
    MakeOpcode 4, &HC8, 1, none, "CStrDate"
    MakeOpcode 4, &HC9, 1, none, "CDateStr"
    MakeOpcode 4, &HCA, 0, 0, "---"
    MakeOpcode 4, &HCB, 0, 0, "---"
    MakeOpcode 4, &HCC, 1, none, "FreeStrNoPop"
    MakeOpcode 4, &HCD, 1, none, "FreeVarNoPop"
    MakeOpcode 4, &HCE, 1, none, "FreeAdNoPop"
    MakeOpcode 4, &HCF, 1, none, "EraseNoPop"
    MakeOpcode 4, &HD0, 3, 0, "WMemLdUI1"
    MakeOpcode 4, &HD1, 3, 0, "WMemLdI2"
    MakeOpcode 4, &HD2, 3, 0, "WMemLdStr"
    MakeOpcode 4, &HD3, 3, 0, "WMemLdStr"
    MakeOpcode 4, &HD4, 3, 0, "WMemLdCy"
    MakeOpcode 4, &HD5, 3, 0, "WMemLdCy"
    MakeOpcode 4, &HD6, 3, 0, "WMemLdRfVar"
    MakeOpcode 4, &HD7, 3, 0, "WMemLdStr"
    MakeOpcode 4, &HD8, 3, 0, "WMemLdStr"
    MakeOpcode 4, &HD9, 3, 0, "WMemLdFPR4"
    MakeOpcode 4, &HDA, 3, 0, "WMemLdFPR8"
    MakeOpcode 4, &HDB, 3, 0, "IWMemLdPr"
    MakeOpcode 4, &HDC, 3, 0, "WMemLdRfVar"
    MakeOpcode 4, &HDD, 3, 0, "WMemLdVar"
    MakeOpcode 4, &HDE, 0, 0, "---"
    MakeOpcode 4, &HDF, 0, 0, "---"
    MakeOpcode 4, &HE0, 3, 0, "WMemStUI1"
    MakeOpcode 4, &HE1, 3, 0, "WMemStI2"
    MakeOpcode 4, &HE2, 3, 0, "WMemStR4"
    MakeOpcode 4, &HE3, 3, 0, "WMemStR4"
    MakeOpcode 4, &HE4, 3, 0, "WMemStR8"
    MakeOpcode 4, &HE5, 3, 0, "WMemStR8"
    MakeOpcode 4, &HE6, 3, 0, "WMemStVar"
    MakeOpcode 4, &HE7, 3, 0, "WMemStStr"
    MakeOpcode 4, &HE8, 3, 0, "WMemStAd"
    MakeOpcode 4, &HE9, 3, 0, "WMemStFPR4"
    MakeOpcode 4, &HEA, 3, 0, "WMemStFPR8"
    MakeOpcode 4, &HEB, 3, 0, "WMemStVarAd"
    MakeOpcode 4, &HEC, 3, 0, "WMemStVarAdFunc"
    MakeOpcode 4, &HED, 3, 0, "WMemStVarUnk"
    MakeOpcode 4, &HEE, 3, 0, "WMemStVarUnkFunc"
    MakeOpcode 4, &HEF, 3, 0, "WMemStAdFunc"
    MakeOpcode 4, &HF0, 3, 0, "WMemStVarCopy"
    MakeOpcode 4, &HF1, 3, 0, "WMemStStrCopy"
    MakeOpcode 4, &HF2, 7, 0, "VarIndexLdRfVarLock"
    MakeOpcode 4, &HF3, 0, 0, "---"
    MakeOpcode 4, &HF4, 0, 0, "---"
    MakeOpcode 4, &HF5, 3, 0, "AssignRecord"
    MakeOpcode 4, &HF6, 5, 0, "DestructAnsiOFrame"
    MakeOpcode 4, &HF7, 3, 0, "FStVarZero"
    MakeOpcode 4, &HF8, 3, 0, "FStVarCopyObj"
    MakeOpcode 4, &HF9, 1, none, "VerifyVarObj"
    MakeOpcode 4, &HFA, 1, none, "VerifyPVarObj"
    MakeOpcode 4, &HFB, 1, none, "FnInStrB4"
    MakeOpcode 4, &HFC, 3, 0, "FnInStrB4Var"
    MakeOpcode 4, &HFD, 1, none, "FnInStr4"
    MakeOpcode 4, &HFE, 3, 0, "FnInStr4Var"
    MakeOpcode 4, &HFF, 1, none, "FnStrComp3"
End Sub

Sub Init5()
    MakeOpcode 5, &H0, 3, 0, "FnStrComp3Var"
    MakeOpcode 5, &H1, 1, none, "StAryMove"
    MakeOpcode 5, &H2, 1, none, "StAryCopy"
    MakeOpcode 5, &H3, 3, 0, "StAryRecMove"
    MakeOpcode 5, &H4, 3, 0, "StAryRecCopy"
    MakeOpcode 5, &H5, 0, 0, "---"
    MakeOpcode 5, &H6, 5, 0, "AryInRecLdPr"
    MakeOpcode 5, &H7, 5, 0, "AryInRecLdRf"
    MakeOpcode 5, &H8, 1, none, "CExtInstUnk"
    MakeOpcode 5, &H9, 3, 0, "IStVarCopyObj"
    MakeOpcode 5, &HA, 1, none, "ArrayRebase1Var"
    MakeOpcode 5, &HB, 1, none, "Assert"
    MakeOpcode 5, &HC, 7, 0, "RaiseEvent"
    MakeOpcode 5, &HD, 5, 0, "PrintObject"
    MakeOpcode 5, &HE, 5, 0, "PrintFile"
    MakeOpcode 5, &HF, 5, 0, "WriteFile"
    MakeOpcode 5, &H10, 5, 0, "InputFile"
    MakeOpcode 5, &H11, 0, 0, "---"
    MakeOpcode 5, &H12, 1, none, "GetRecFxStr3"
    MakeOpcode 5, &H13, 1, none, "GetRecFxStr4"
    MakeOpcode 5, &H14, 1, none, "PutRecFxStr3"
    MakeOpcode 5, &H15, 1, none, "PutRecFxStr4"
    MakeOpcode 5, &H16, 3, 0, "GetRecOwn3"
    MakeOpcode 5, &H17, 3, 0, "GetRecOwn4"
    MakeOpcode 5, &H18, 3, 0, "PutRecOwn3"
    MakeOpcode 5, &H19, 3, 0, "PutRecOwn4"
    MakeOpcode 5, &H1A, 2, 0, "LitI2_Byte"
    MakeOpcode 5, &H1B, 1, none, "CBoolVarNull"
    MakeOpcode 5, &H1C, 2, 0, "LargeBos"
    MakeOpcode 5, &H1D, 0, 0, "Bos"
    MakeOpcode 5, &H1E, 5, 0, "ImpAdCallNonVirt"
    MakeOpcode 5, &H1F, 0, 0, "---"
    MakeOpcode 5, &H20, 0, 0, "---"
    MakeOpcode 5, &H21, 0, 0, "---"
    MakeOpcode 5, &H22, 0, 0, "---"
    MakeOpcode 5, &H23, 0, 0, "---"
    MakeOpcode 5, &H24, 0, 0, "---"
    MakeOpcode 5, &H25, 0, 0, "---"
    MakeOpcode 5, &H26, 0, 0, "---"
    MakeOpcode 5, &H27, 0, 0, "---"
    MakeOpcode 5, &H28, 0, 0, "---"
    MakeOpcode 5, &H29, 0, 0, "---"
    MakeOpcode 5, &H2A, 3, 0, "DestructRecord"
    MakeOpcode 5, &H2B, 3, 0, "VCallFPR8"
    MakeOpcode 5, &H2C, 3, 0, "ThisVCallHidden"
    MakeOpcode 5, &H2D, 1, none, "ZeroRetVal"
    MakeOpcode 5, &H2E, 1, none, "ZeroRetValVar"
    MakeOpcode 5, &H2F, 5, 0, "ExitProcCbHresult"
    MakeOpcode 5, &H30, 7, 0, "ExitProcFrameCbHresult"
    MakeOpcode 5, &H31, 3, 0, "EraseDestrKeepData"
    MakeOpcode 5, &H32, 5, 0, "CDargRefUdt"
    MakeOpcode 5, &H33, 5, 0, "CVarRefUdt"
    MakeOpcode 5, &H34, 5, 0, "CVarUdt"
    MakeOpcode 5, &H35, 3, 0, "StUdtVar"
    MakeOpcode 5, &H36, 3, 0, "StAryVar"
    MakeOpcode 5, &H37, 3, 0, "CopyBytesZero"
    MakeOpcode 5, &H38, 5, 0, "FLdZeroAry"
    MakeOpcode 5, &H39, 3, 0, "FStVarZero"
    MakeOpcode 5, &H3A, 7, 0, "CVarAryUdt"
    MakeOpcode 5, &H3B, 7, 0, "RedimVarUdt"
    MakeOpcode 5, &H3C, 7, 0, "RedimPreserveVarUdt"
    MakeOpcode 5, &H3D, 5, 0, "VarLateMemLdRfVar"
    MakeOpcode 5, &H3E, 6, 0, "VarLateMemCallLdRfVar"
    MakeOpcode 5, &H3F, 0, 0, "---"
    MakeOpcode 5, &H40, 0, 0, "---"
    MakeOpcode 5, &H41, 5, 0, "VarLateMemLdVar"
    MakeOpcode 5, &H42, 7, 0, "VarLateMemCallLdVar"
    MakeOpcode 5, &H43, 3, 0, "VarLateMemSt"
    MakeOpcode 5, &H44, 5, 0, "VarLateMemCallSt"
    MakeOpcode 5, &H45, 5, 0, "VarLateMemStAd"
    MakeOpcode 5, &H46, 0, 0, "---"
    MakeOpcode 5, &H47, 0, 0, "Unknow"
    MakeOpcode 5, &H48, 0, 0, "Unknow"
    MakeOpcode 5, &H49, 0, 0, "Unknow"
    MakeOpcode 5, &H4A, 0, 0, "Unknow"
    MakeOpcode 5, &H4B, 0, 0, "Unknow"
    MakeOpcode 5, &H4C, 0, 0, "Unknow"
    MakeOpcode 5, &H4D, 0, 0, "Unknow"
    MakeOpcode 5, &H4E, 0, 0, "Unknow"
    MakeOpcode 5, &H4F, 0, 0, "Unknow"
    MakeOpcode 5, &H50, 0, 0, "Unknow"
    MakeOpcode 5, &H51, 0, 0, "Unknow"
    MakeOpcode 5, &H52, 0, 0, "Unknow"
    MakeOpcode 5, &H53, 0, 0, "Unknow"
    MakeOpcode 5, &H54, 0, 0, "Unknow"
    MakeOpcode 5, &H55, 0, 0, "Unknow"
    MakeOpcode 5, &H56, 0, 0, "Unknow"
    MakeOpcode 5, &H57, 0, 0, "Unknow"
    MakeOpcode 5, &H58, 0, 0, "Unknow"
    MakeOpcode 5, &H59, 0, 0, "Unknow"
    MakeOpcode 5, &H5A, 0, 0, "Unknow"
    MakeOpcode 5, &H5B, 0, 0, "Unknow"
    MakeOpcode 5, &H5C, 0, 0, "Unknow"
    MakeOpcode 5, &H5D, 0, 0, "Unknow"
    MakeOpcode 5, &H5E, 0, 0, "Unknow"
    MakeOpcode 5, &H5F, 0, 0, "Unknow"
    MakeOpcode 5, &H60, 0, 0, "Unknow"
    MakeOpcode 5, &H61, 0, 0, "Unknow"
    MakeOpcode 5, &H62, 0, 0, "Unknow"
    MakeOpcode 5, &H63, 0, 0, "Unknow"
    MakeOpcode 5, &H64, 0, 0, "Unknow"
    MakeOpcode 5, &H65, 0, 0, "Unknow"
    MakeOpcode 5, &H66, 0, 0, "Unknow"
    MakeOpcode 5, &H67, 0, 0, "Unknow"
    MakeOpcode 5, &H68, 0, 0, "Unknow"
    MakeOpcode 5, &H69, 0, 0, "Unknow"
    MakeOpcode 5, &H6A, 0, 0, "Unknow"
    MakeOpcode 5, &H6B, 0, 0, "Unknow"
    MakeOpcode 5, &H6C, 0, 0, "Unknow"
    MakeOpcode 5, &H6D, 0, 0, "Unknow"
    MakeOpcode 5, &H6E, 0, 0, "Unknow"
    MakeOpcode 5, &H6F, 0, 0, "Unknow"
    MakeOpcode 5, &H70, 0, 0, "Unknow"
    MakeOpcode 5, &H71, 0, 0, "Unknow"
    MakeOpcode 5, &H72, 0, 0, "Unknow"
    MakeOpcode 5, &H73, 0, 0, "Unknow"
    MakeOpcode 5, &H74, 0, 0, "Unknow"
    MakeOpcode 5, &H75, 0, 0, "Unknow"
    MakeOpcode 5, &H76, 0, 0, "Unknow"
    MakeOpcode 5, &H77, 0, 0, "Unknow"
    MakeOpcode 5, &H78, 0, 0, "Unknow"
    MakeOpcode 5, &H79, 0, 0, "Unknow"
    MakeOpcode 5, &H7A, 0, 0, "Unknow"
    MakeOpcode 5, &H7B, 0, 0, "Unknow"
    MakeOpcode 5, &H7C, 0, 0, "Unknow"
    MakeOpcode 5, &H7D, 0, 0, "Unknow"
    MakeOpcode 5, &H7E, 0, 0, "Unknow"
    MakeOpcode 5, &H7F, 0, 0, "Unknow"
    MakeOpcode 5, &H80, 0, 0, "Unknow"
    MakeOpcode 5, &H81, 0, 0, "Unknow"
    MakeOpcode 5, &H82, 0, 0, "Unknow"
    MakeOpcode 5, &H83, 0, 0, "Unknow"
    MakeOpcode 5, &H84, 0, 0, "Unknow"
    MakeOpcode 5, &H85, 0, 0, "Unknow"
    MakeOpcode 5, &H86, 0, 0, "Unknow"
    MakeOpcode 5, &H87, 0, 0, "Unknow"
    MakeOpcode 5, &H88, 0, 0, "Unknow"
    MakeOpcode 5, &H89, 0, 0, "Unknow"
    MakeOpcode 5, &H8A, 0, 0, "Unknow"
    MakeOpcode 5, &H8B, 0, 0, "Unknow"
    MakeOpcode 5, &H8C, 0, 0, "Unknow"
    MakeOpcode 5, &H8D, 0, 0, "Unknow"
    MakeOpcode 5, &H8E, 0, 0, "Unknow"
    MakeOpcode 5, &H8F, 0, 0, "Unknow"
    MakeOpcode 5, &H90, 0, 0, "Unknow"
    MakeOpcode 5, &H91, 0, 0, "Unknow"
    MakeOpcode 5, &H92, 0, 0, "Unknow"
    MakeOpcode 5, &H93, 0, 0, "Unknow"
    MakeOpcode 5, &H94, 0, 0, "Unknow"
    MakeOpcode 5, &H95, 0, 0, "Unknow"
    MakeOpcode 5, &H96, 0, 0, "Unknow"
    MakeOpcode 5, &H97, 0, 0, "Unknow"
    MakeOpcode 5, &H98, 0, 0, "Unknow"
    MakeOpcode 5, &H99, 0, 0, "Unknow"
    MakeOpcode 5, &H9A, 0, 0, "Unknow"
    MakeOpcode 5, &H9B, 0, 0, "Unknow"
    MakeOpcode 5, &H9C, 0, 0, "Unknow"
    MakeOpcode 5, &H9D, 0, 0, "Unknow"
    MakeOpcode 5, &H9E, 0, 0, "Unknow"
    MakeOpcode 5, &H9F, 0, 0, "Unknow"
    MakeOpcode 5, &HA0, 0, 0, "Unknow"
    MakeOpcode 5, &HA1, 0, 0, "Unknow"
    MakeOpcode 5, &HA2, 0, 0, "Unknow"
    MakeOpcode 5, &HA3, 0, 0, "Unknow"
    MakeOpcode 5, &HA4, 0, 0, "Unknow"
    MakeOpcode 5, &HA5, 0, 0, "Unknow"
    MakeOpcode 5, &HA6, 0, 0, "Unknow"
    MakeOpcode 5, &HA7, 0, 0, "Unknow"
    MakeOpcode 5, &HA8, 0, 0, "Unknow"
    MakeOpcode 5, &HA9, 0, 0, "Unknow"
    MakeOpcode 5, &HAA, 0, 0, "Unknow"
    MakeOpcode 5, &HAB, 0, 0, "Unknow"
    MakeOpcode 5, &HAC, 0, 0, "Unknow"
    MakeOpcode 5, &HAD, 0, 0, "Unknow"
    MakeOpcode 5, &HAE, 0, 0, "Unknow"
    MakeOpcode 5, &HAF, 0, 0, "Unknow"
    MakeOpcode 5, &HB0, 0, 0, "Unknow"
    MakeOpcode 5, &HB1, 0, 0, "Unknow"
    MakeOpcode 5, &HB2, 0, 0, "Unknow"
    MakeOpcode 5, &HB3, 0, 0, "Unknow"
    MakeOpcode 5, &HB4, 0, 0, "Unknow"
    MakeOpcode 5, &HB5, 0, 0, "Unknow"
    MakeOpcode 5, &HB6, 0, 0, "Unknow"
    MakeOpcode 5, &HB7, 0, 0, "Unknow"
    MakeOpcode 5, &HB8, 0, 0, "Unknow"
    MakeOpcode 5, &HB9, 0, 0, "Unknow"
    MakeOpcode 5, &HBA, 0, 0, "Unknow"
    MakeOpcode 5, &HBB, 0, 0, "Unknow"
    MakeOpcode 5, &HBC, 0, 0, "Unknow"
    MakeOpcode 5, &HBD, 0, 0, "Unknow"
    MakeOpcode 5, &HBE, 0, 0, "Unknow"
    MakeOpcode 5, &HBF, 0, 0, "Unknow"
    MakeOpcode 5, &HC0, 0, 0, "Unknow"
    MakeOpcode 5, &HC1, 0, 0, "Unknow"
    MakeOpcode 5, &HC2, 0, 0, "Unknow"
    MakeOpcode 5, &HC3, 0, 0, "Unknow"
    MakeOpcode 5, &HC4, 0, 0, "Unknow"
    MakeOpcode 5, &HC5, 0, 0, "Unknow"
    MakeOpcode 5, &HC6, 0, 0, "Unknow"
    MakeOpcode 5, &HC7, 0, 0, "Unknow"
    MakeOpcode 5, &HC8, 0, 0, "Unknow"
    MakeOpcode 5, &HC9, 0, 0, "Unknow"
    MakeOpcode 5, &HCA, 0, 0, "Unknow"
    MakeOpcode 5, &HCB, 0, 0, "Unknow"
    MakeOpcode 5, &HCC, 0, 0, "Unknow"
    MakeOpcode 5, &HCD, 0, 0, "Unknow"
    MakeOpcode 5, &HCE, 0, 0, "Unknow"
    MakeOpcode 5, &HCF, 0, 0, "Unknow"
    MakeOpcode 5, &HD0, 0, 0, "Unknow"
    MakeOpcode 5, &HD1, 0, 0, "Unknow"
    MakeOpcode 5, &HD2, 0, 0, "Unknow"
    MakeOpcode 5, &HD3, 0, 0, "Unknow"
    MakeOpcode 5, &HD4, 0, 0, "Unknow"
    MakeOpcode 5, &HD5, 0, 0, "Unknow"
    MakeOpcode 5, &HD6, 0, 0, "Unknow"
    MakeOpcode 5, &HD7, 0, 0, "Unknow"
    MakeOpcode 5, &HD8, 0, 0, "Unknow"
    MakeOpcode 5, &HD9, 0, 0, "Unknow"
    MakeOpcode 5, &HDA, 0, 0, "Unknow"
    MakeOpcode 5, &HDB, 0, 0, "Unknow"
    MakeOpcode 5, &HDC, 0, 0, "Unknow"
    MakeOpcode 5, &HDD, 0, 0, "Unknow"
    MakeOpcode 5, &HDE, 0, 0, "Unknow"
    MakeOpcode 5, &HDF, 0, 0, "Unknow"
    MakeOpcode 5, &HE0, 0, 0, "Unknow"
    MakeOpcode 5, &HE1, 0, 0, "Unknow"
    MakeOpcode 5, &HE2, 0, 0, "Unknow"
    MakeOpcode 5, &HE3, 0, 0, "Unknow"
    MakeOpcode 5, &HE4, 0, 0, "Unknow"
    MakeOpcode 5, &HE5, 0, 0, "Unknow"
    MakeOpcode 5, &HE6, 0, 0, "Unknow"
    MakeOpcode 5, &HE7, 0, 0, "Unknow"
    MakeOpcode 5, &HE8, 0, 0, "Unknow"
    MakeOpcode 5, &HE9, 0, 0, "Unknow"
    MakeOpcode 5, &HEA, 0, 0, "Unknow"
    MakeOpcode 5, &HEB, 0, 0, "Unknow"
    MakeOpcode 5, &HEC, 0, 0, "Unknow"
    MakeOpcode 5, &HED, 0, 0, "Unknow"
    MakeOpcode 5, &HEE, 0, 0, "Unknow"
    MakeOpcode 5, &HEF, 0, 0, "Unknow"
    MakeOpcode 5, &HF0, 0, 0, "Unknow"
    MakeOpcode 5, &HF1, 0, 0, "Unknow"
    MakeOpcode 5, &HF2, 0, 0, "Unknow"
    MakeOpcode 5, &HF3, 0, 0, "Unknow"
    MakeOpcode 5, &HF4, 0, 0, "Unknow"
    MakeOpcode 5, &HF5, 0, 0, "Unknow"
    MakeOpcode 5, &HF6, 0, 0, "Unknow"
    MakeOpcode 5, &HF7, 0, 0, "Unknow"
    MakeOpcode 5, &HF8, 0, 0, "Unknow"
    MakeOpcode 5, &HF9, 0, 0, "Unknow"
    MakeOpcode 5, &HFA, 0, 0, "Unknow"
    MakeOpcode 5, &HFB, 0, 0, "Unknow"
    MakeOpcode 5, &HFC, 0, 0, "Unknow"
    MakeOpcode 5, &HFD, 0, 0, "Unknow"
    MakeOpcode 5, &HFE, 0, 0, "Unknow"
    MakeOpcode 5, &HFF, 0, 0, "Unknow"
    


End Sub
'&=long
'$=string
'%=integer
'Sub MakeOpcode(index&, num&, length&, flag&, mnem$)
Sub MakeOpcode(index As Long, Num As Long, Length As Long, Flag As Long, Mnem As String)
    OPCode(index, Num).Mnemonic = Mnem
    OPCode(index, Num).Flag = Flag
    OPCode(index, Num).Size = Length
End Sub

Function MakeHex$(a)
    If a < 16 Then MakeHex = "0" + Hex(a) Else MakeHex = Hex(a)
End Function

Function MakeHex16$(a1, a2)
    MakeHex16 = MakeHex(a2) + MakeHex(a1)
End Function

Function MakeHex32$(a1, a2, a3, a4)
    MakeHex32 = MakeHex(a4) + MakeHex(a3) + MakeHex(a2) + MakeHex(a1)
End Function

Function File16&(addr As Long)
    Dim t%
    CopyMemory t%, File(addr), 2
    File16 = t
End Function

Function File32&(addr As Long)
On Error Resume Next
    Dim t&
    CopyMemory t&, File(addr), 4
    File32 = t
End Function

Function FileZ$(addr As Long)
    Dim t$, a&
    a = addr
    Do While File(a) <> 0
        t = t + Chr(File(a))
        a = a + 1
    Loop
    FileZ = t
End Function

Function FileW$(addr As Long)
    Dim t$, a&
    If addr > 0 Then
        a = addr
        Do While File16(a) <> 0
            t = t + ChrW(File16(a))
            a = a + 2
        Loop
    End If
    FileW = t
End Function

Function GetByte(addr As Long) As Long
    GetByte = File(addr)
    addr = addr + 1
End Function

Function Cvl(a As String)
    Dim i&
    CopyMemory i&, a$, 4
    Cvl = i
End Function

Sub AddMap(a As Long)
On Error Resume Next
    Map(a&) = Map(a) Or 1
End Sub

Sub AddRef(a As Long, src As Long)
On Error Resume Next
    Map(a&) = Map(a) Or 2
    If InStr(RefName(a), Hex(src)) = 0 Then
        If RefName(a&) <> vbNullString Then RefName(a&) = RefName(a&) + ", "
        RefName(a&) = RefName(a&) + Hex(src&)
    End If
End Sub

Function HasMap&(a As Long)
On Error Resume Next
    If Map(a) And 1 Then HasMap = -1
End Function

Function HasRef&(a As Long)
On Error Resume Next
    If Map(a) And 2 Then HasRef = -1
End Function

Sub ReadOpcode(f$, o() As OpcodeType, index As Long)
    Dim t$, i&, c&
    Open f$ For Input As #10
        Line Input #10, t
        Do
            Line Input #10, t
            i = InStr(t$, Chr$(9))
            If i Then
                t = Mid(t, i + 1)
                i = InStr(t$, Chr$(9))
                c = Val("&h" + Mid(t, 1, i - 1))
                t = Mid(t, i + 1)
                i = InStr(t$, Chr$(9))
                o(index, c).Mnemonic = Mid(t, 1, i - 1)
                t = Mid(t, i + 1)
                i = InStr(t$, Chr$(9))
                o(index, c).Size = Val(Mid(t, 1, i - 1))
                Print #3, "MakeOpcode"; index; ",&h"; MakeHex(c); ","; o(index, c).Size; ","; "0,"; Chr(34); o(index, c).Mnemonic; (Chr(34))
            End If
        Loop Until EOF(10)
    Close #10
End Sub

Sub init()
    Init0
    Init1
    Init2
    Init3
    Init4
    Init5
    
End Sub

Sub AddProc(addr As Long)
    If HasMap(addr) = 0 Then
        ProcList(ProcCnt) = addr
        ReDim Preserve ProcList(UBound(ProcList) + 1)
        ProcCnt = ProcCnt + 1
        AddMap addr
    End If
End Sub

Function FastName$(prefix As String, addr As Long)
    If addr >= base And addr < base + PESize Then
        If SubName(addr&) > vbNullString Then
            FastName = SubName(addr)
            Exit Function
        End If
    End If
    FastName = prefix + "_" + Hex(addr)
End Function

Function GetOpp$(a As String, b As String)
    Dim i&
    If b$ <> Chr(13) Then a = Trim(a)
    i = InStr(a, b$)
    If i = 0 Then
        GetOpp = a
        a = vbNullString
    ElseIf a = b Then
        GetOpp = a
        a = vbNullString
    ElseIf i = 1 Then
        GetOpp = vbNullString
        a = Mid(a, Len(b) + 1)
    ElseIf i + Len(b) >= Len(a) + 1 Then
        GetOpp = Mid(a, 1, i - 1)
        a = vbNullString
    Else
        GetOpp = Mid(a, 1, i - 1)
        a = Mid(a, i + Len(b))
    End If
End Function

Function MakeAddr$(t As Long)
    Dim u$
    If t >= base And t < base + PESize Then
        If File(t) = &HBA And File(t + 5) = &HB9 Then
            t = File32(t + 1)
            u = FastName("proc", t)
            AddProc t
        ElseIf File16(t) = &H25FF Then
            t = File32(t + 2)
            u = FastName("ext", t) 'External import
            
            
        Else
            u = FastName("unk", t)
        End If
    Else
        u = " ???=" + Hex(t)
    End If
    MakeAddr = u
End Function


Function ConvertStrToVB(Mnem As String, addr As Long, pool As Long, origpool As Long, ProcPC As Long)
'Arguement Type
'vbgamer45
    Dim a&, c$, t&, u$, i&, j&
    'MsgBox Mnem
    i = addr
    For a = 1 To Len(Mnem)
        c = Mid(Mnem, a, 1)
        If c <> "%" Then
          '  u = u + c
        Else
            a = a + 1
            c = Mid(Mnem, a, 1)
            Select Case c
                Case "a"
                    t = File16(i)
                    u = u + MakeArg(t)
                    'u = u & "var" & t
                    i = i + 2
                Case "c"
                    t = File32(File16(i) * 4 + pool)
                    u = u + MakeAddrToVB(t)
                    i = i + 2
                Case "e"
                    t = File32(File16(i) * 4 + pool) + File16(i + 2)
                    u = u + MakeAddrToVB(t)
                    i = i + 4
                Case "s"
                    t = File32(File16(i) * 4 + pool)
                   ' u = u + FastName("v", t) + " '" + FileW(t) + "' "
                    u = u & Chr(34) & FileW(t) & Chr(34)
                    i = i + 2
                Case "l"
                    t = ProcPC + File16(i)
                    u = u + FastName("loc", t)
                    AddRef t, i - 1
                    i = i + 2
                Case "1", "2", "4"
                    For j = 1 To Val(c)
                        u = u + MakeHex(File(i + Val(c) - j))
                    Next
                    i = i + Val(c)
                Case "t"
                    t = File32(File16(i) * 4 + origpool)
                    u = u + FastName("xxx", t)
                    pool = t
                Case Else
              
                    u = u + c
            End Select
        End If
    Next
    ConvertStrToVB = u
End Function
Function FastNameToVB(prefix As String, addr As Long) As String
    If addr >= base And addr < base + PESize Then
        If SubName(addr&) > vbNullString Then
            FastNameToVB = SubName(addr)
            Exit Function
        End If
    End If
    FastNameToVB = prefix + "_" + Hex(addr)
End Function
Function MakeAddrToVB(t As Long) As String
    Dim u$
    If t >= base And t < base + PESize Then
        If File(t) = &HBA And File(t + 5) = &HB9 Then
            t = File32(t + 1)
            u = FastName("proc", t)
            AddProc t
        ElseIf File16(t) = &H25FF Then
            t = File32(t + 2)
            u = FastName("ext", t) 'External import
            Dim f As Integer
            f = FreeFile
            Dim getImportRva As Long
            Dim j As Integer
            Open SFilePath For Binary Access Read As f
                Seek f, t + 1 - OptHeader.ImageBase
                Get f, , getImportRva
                'MsgBox getImportRva
            Close f
            Dim TDs As String, ouR As String
            For j = 1 To UBound(exeIMPORT_APINAME)
                If exeIMPORT_APINAME(j).Address = getImportRva Then
                    If Left$(exeIMPORT_APINAME(j).ApiName, 8) = "!ordinal" Then
                        'via ordinal
                        TDs = VBFunction_Description(Val(Mid$(exeIMPORT_APINAME(j).ApiName, 12)), vbNullString, ouR)
                        If TDs = "undef" Then
                            'tvProject.Nodes.Add exeIMPORT_APINAME(j).ApiName, 4, , "Name : " & ouR, 18
                        Else
                            'tvProject.Nodes.Add exeIMPORT_APINAME(j).ApiName, 4, , "Name: " & ouR, 18
                          u = TDs
                            'tvProject.Nodes.Add exeIMPORT_APINAME(j).ApiName, 4, , TDs, 19
                        End If
                        u = TDs
                    Else
                        'via directname
                        TDs = VBFunction_Description(0, exeIMPORT_APINAME(j).ApiName, ouR)
                        If TDs = "undef" Then
                        Else
                            u = TDs
                           'tvProject.Nodes.Add exeIMPORT_APINAME(j).ApiName, 4, , TDs, 19
                        End If
                        u = TDs
                    End If
                     MakeAddrToVB = u
                    'MsgBox TDs
                    Exit For
                End If
            Next
        Else
            u = FastName("unk", t)
        End If
    Else
        u = " ???=" + Hex(t)
    End If
    MakeAddrToVB = u
End Function

