Attribute VB_Name = "modPCodeToVB"
'#############################################
'modPCodeToVB VisualBasicZone.com 2004-2005
'##############################################
Public Function ReturnVBCodeByPcodeToken(ByVal Address As Long, ByVal strOpcode As String, ByVal strArguments As String, ByRef RemoveOpcode As Boolean)
    Dim Temp() As String
    Dim strHold As String
    Temp = Split(strOpcode, " ")
    Select Case Temp(0)
        Case "ExitProc"
            strHold = "End Sub"
            RemoveOpcode = True
        Case "ExitProcHresult"
            strHold = "End Sub"
            RemoveOpcode = True
        Case "ExitProcCbHresult"
            strHold = "End Function"
            RemoveOpcode = True
        Case "FLdRfVar"
            strHold = strArguments & " = " & modPCode.ReturnOneVar
            RemoveOpcode = True
        Case Else
            RemoveOpcode = False
    End Select
    ReturnVBCodeByPcodeToken = strHold
End Function


Public Function ReturnVBCodeByPcodeToken2(Address As Long, strOpcode As String, strArguments As String)
    Dim Temp() As String
    'Dim strHold
    'strOpcode = strOpcode ' & " "
    Temp = Split(strOpcode, " ")
 
    'strHold = strOpcode
 
    Select Case Temp(0)
        
        Case "ExitProc"
            strHold = "End Sub"
        Case "ExitProcHresult"
            strHold = "End Sub"
        Case "ExitProcI2"
            'Return Integer
        Case "ExitProcR4"
        
        Case "ExitProcR8"
        
        Case "ExitProcCy"
        
        Case "ExitProcCbHresult"
            strHold = "End Function"
            
        Case "LitI2_Byte"
           strHold = "'Dim byte" & Address & " as Byte" & vbCrLf
           strHold = strHold & "'byte" & Address & " = " & Hex2Dec(strArguments)
            'Store the Byte
            'AddValueToRegister Hex2Dec(strArguments), "byte" & Address, Address
            TempValue = Hex2Dec(strArguments)
            TempAsmReg.varAddress = Address
            TempAsmReg.varValue = TempValue
        Case "LitI2"
            strHold = "'Dim int" & Address & " as Integer" & vbCrLf
            strHold = strHold & "'int" & Address & " = " & Hex2Dec(strArguments)
            'Store the integer
            'AddValueToRegister Hex2Dec(strArguments), "int" & Address, Address
            TempValue = Hex2Dec(strArguments)
            TempAsmReg.varAddress = Address
            TempAsmReg.varValue = TempValue
       Case "LitStr"
            strHold = "'Dim str" & Address & " as String" & vbCrLf
            strHold = strHold & "'str" & Address & " = " & strArguments
            TempValue = strArguments
        Case "LitVarStr"
            strHold = "'Dim str" & Address & " as String" & vbCrLf
            strHold = strHold & "'str" & Address & " = " & strArguments
            TempValue = strArguments
       Case "LitI4"
            strHold = "'Dim lng" & Address & " as Long" & vbCrLf
            strHold = strHold & "'lng" & Address & " = " & Hex2Dec(strArguments)
            'Store the Long
            'AddValueToRegister Hex2Dec(strArguments), "lng" & Address, Address
            TempValue = Hex2Dec(strArguments)
            TempAsmReg.varAddress = Address
            TempAsmReg.varValue = TempValue
        Case "AddI2"
            'strHold = "'var" & AsmRegister(0).varAddress & " = " & "var" & AsmRegister(0).varAddress & " + " & TempAsmReg.varValue '& AsmRegister(1).varValue
        Case "SubI2"
            'strHold = "'var" & AsmRegister(0).varAddress & " = " & "var" & AsmRegister(0).varAddress & " - " & TempAsmReg.varValue
        Case "MulI2"
            'strHold = "'var" & AsmRegister(0).varAddress & " = " & "var" & AsmRegister(0).varAddress & " * " & TempAsmReg.varValue
        Case "DivR8"
            'strHold = "'var" & AsmRegister(0).varAddress & " = " & "var" & AsmRegister(0).varAddress & " / " & TempAsmReg.varValue
        Case "ImpAdCallFPR4"
           ' strHold = "'" & strArguments
        'Case "ImpAdCallFPR4"
        
        Case "FStI2"
        'Store I2
           

        Case "FLdI2"
        'Load I2
        Case "LargeBos"
        
        Case "LitVar_Missing"
        
        Case Else
            strHold = "'" & strOpcode & strArguments
    
    End Select
    
    'Return the vbcode
    ReturnVBCodeByPcodeToken2 = strHold
    
End Function

Public Function Hex2Dec(sText As String) As Long
    On Error GoTo err
    Dim H As String
    H = sText
    Dim tmp$
    Dim lo1 As Integer, lo2 As Integer
    Dim hi1 As Long, hi2 As Long
    Const Hx = "&H"
    Const BigShift = 65536
    Const LilShift = 256, Two = 2
    tmp = H
    If UCase$(Left$(H, 2)) = "&H" Then tmp = Mid$(H, 3)
    tmp = Right$("0000000" & tmp, 8)
        If IsNumeric(Hx & tmp) Then
            lo1 = CInt(Hx & Right$(tmp, Two))
            hi1 = CLng(Hx & Mid$(tmp, 5, Two))
            lo2 = CInt(Hx & Mid$(tmp, 3, Two))
            hi2 = CLng(Hx & Left$(tmp, Two))
            Hex2Dec = CCur(hi2 * LilShift + lo2) * BigShift + (hi1 * LilShift) + lo1
        End If
Exit Function
err:
End Function

