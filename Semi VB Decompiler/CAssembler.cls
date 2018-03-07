VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAssembler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com

Private Type t_asmmodel
    code(15) As Byte
    mask(15) As Byte
    Length As Long
    jmpsize As Long
    jmpoffset As Long
    jmppos As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Private Declare Function Assemble Lib "olly.dll" ( _
        ByVal CMD As String, ByVal ip As Long, model As t_asmmodel, _
        ByVal attempt As Long, ByVal constsize As Long, ByVal errtext As String) As Long

Private Declare Sub VB_SetOptions Lib "olly.dll" ( _
        Optional ByVal isideal As Long = 0, Optional ByVal isLower As Long = 0, _
        Optional doTabs As Long = 0, Optional ByVal dispseg As Long = 1)

Private lblNames As New Collection
Private lblAddr As New Collection
Private apiNames As New Collection
Private apiAddr As New Collection

Dim sc As Object

Public ErrorMessage As String

Private Sub Class_Initialize()
    On Error Resume Next
    Set sc = CreateObject("MSScriptControl.ScriptControl") 'not required
    If Not sc Is Nothing Then sc.Language = "VBScript"
End Sub

Private Function isUserLabel(ByVal r) As Boolean
    On Error GoTo hell
    r = lblNames(r)
    isUserLabel = True
    Exit Function
hell:
End Function

Sub ClearApi()
    Set apiNames = New Collection
    Set apiAddr = New Collection
End Sub

Function AddApi(name As String, addr As Long)
    On Error Resume Next
    name = UCase$(name)
    apiNames.Add name, name
    apiAddr.Add "[ 0" & Hex$(addr) & "h]", name
End Function

 
Function AsmBlock(ByVal inst As String, base As Long) As Byte()
    Dim tmp() As String
    Dim i As Integer
    Dim b() As Byte
    Dim offset As Long
    Dim leng As Long
    Dim errLen As Long
    Dim dis() As String
    Dim mathstring As String
    Dim evalOffset As String
    Dim label
    Dim ret() As Byte
    
    Set lblNames = New Collection
    Set lblAddr = New Collection
    
    On Error GoTo hell
    
    offset = base
    tmp = Split(UCase$(inst), vbCrLf)
    
    For i = 0 To UBound(tmp)
       
        If InStr(tmp(i), "$") > 0 Then 'support $+x for currentaddr + x notation
            If sc Is Nothing Then
                MsgBox "Could not create instance of Script Control", vbExclamation
                err.Raise 1
            End If
            mathstring = Mid$(tmp(i), InStr(tmp(i), "$"))
            evalOffset = Replace(mathstring, "$", offset)
            evalOffset = sc.Eval(evalOffset)
            tmp(i) = Replace(tmp(i), mathstring, Hex(evalOffset))
        End If

        If Right$(Trim$(tmp(i)), 1) = ":" Then
            label = UCase$(Trim$(Replace(tmp(i), ":", Empty)))
            lblNames.Add label, label
            lblAddr.Add Hex(offset), label
        ElseIf Len(Trim$(tmp(i))) > 0 Then
        
            For Each label In lblNames
                If InStr(tmp(i), label) > 0 Then
                    tmp(i) = Replace(tmp(i), label, lblAddr(label))
                End If
            Next
            
            For Each label In apiNames
                If InStr(tmp(i), label) > 0 Then
                    tmp(i) = Replace(tmp(i), label, apiAddr(label))
                End If
            Next
            
            b() = AsmLine(tmp(i), offset, leng, errLen)
            
            offset = offset + leng
            
            aryJoin ret, b
            
            If errLen > 1 Then GoTo asmerr

        End If
        
    Next
    
    AsmBlock = ret()
    
hell:
Exit Function
asmerr:
        MsgBox ErrorMessage, vbInformation
        
End Function

Function AsmLine(line As String, offset As Long, asmLen As Long, errLen As Long) As Byte()
    Dim leng As Long
    Dim am As t_asmmodel
    Dim b() As Byte
    Dim i As Integer
    
    ErrorMessage = String$(256, Chr(0))
    asmLen = Assemble(line, offset, am, 0, 0, ErrorMessage)
    
    errLen = InStr(ErrorMessage, Chr(0))
    If asmLen < 0 Then Exit Function
    
    ReDim b(1 To asmLen)
    CopyMemory b(1), am.code(0), asmLen
    
    AsmLine = b()
    
End Function

Private Sub aryJoin(mainAry1, ary2)
    Dim i As Long
    
    On Error GoTo hell
    
    For i = LBound(ary2) To UBound(ary2)
        push mainAry1, ary2(i)
    Next
    
hell:
End Sub

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

