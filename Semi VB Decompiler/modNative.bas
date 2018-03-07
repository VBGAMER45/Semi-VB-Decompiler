Attribute VB_Name = "modNative"
'*********************************************
'modNative
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************
'Module for Processing Native Code
Option Explicit
Private Type API_VBDEF
    rva As Long
    Ordinal As Long
    uName As String
    uDescr As String
End Type
Public exeVB6_APIDEF() As API_VBDEF

Private Type NativeProcType
    offset As Long
    sName As String
End Type
Global gNativeProcArray() As NativeProcType
Global NativeShowOffsets As Boolean
Global NativeShowHexInformation As Boolean

Public Sub Decode(ByVal Filename As String)
'*****************************
'Purpose: To Get the procdures of a Native Exe and produce a report
'*****************************

Dim FileNum As Integer
    Dim F As Long
    F = FreeFile
    Open App.Path & "\dump\" & SFile & "\NativeOut.txt" For Output As #F
        Print #F, "Semi VB Decompiler - VisualBasicZone.com"
        Print #F, "Native Output for : " & Filename
        Print #F, "---------------------------------"
       
        Print #F, "Procedure Offsets:"
        If gProjectInfo.aNativeCode <> 0 Then
            If gVBHeader.aSubMain <> 0 Then
                 Print #F, gVBHeader.aSubMain
            End If
        End If
        Dim i As Integer
        For i = 0 To UBound(gNativeProcArray) - 1
             Print #F, gNativeProcArray(i).offset
        Next i
    Close #F
    
End Sub

Sub VBFunction_Description_Init(ByVal fRes As String)
'*****************************
'Purpose: To load the Msvbvm60.dll api list from a file
'*****************************
Dim lfp As Integer, i As Long
Dim sAdr As String, sOrd As String, sName As String, sDef As String
lfp = FreeFile
Erase exeVB6_APIDEF()

    Open fRes For Input Access Read As #lfp
        i = 0
        Do
        i = i + 1
            Input #lfp, sAdr, sOrd, sName, sDef
            If LCase$(sAdr) <> "eof" Then
                ReDim Preserve exeVB6_APIDEF(1 To i)
                exeVB6_APIDEF(i).rva = Val("&H" & sAdr)

                exeVB6_APIDEF(i).Ordinal = CLng(sOrd)
                exeVB6_APIDEF(i).uName = sName
                exeVB6_APIDEF(i).uDescr = sDef
            Else
                Exit Do
            End If
        Loop Until EOF(1)
    
    Close #lfp

End Sub
Public Function VBFunction_Description(ByVal inOrdinal As Long, ByVal inAPIname As String, ByRef outRName As String) As String
'*****************************
'Purpose: To return the description of a function
'*****************************
Dim i As Long


If inOrdinal > 0 And inAPIname = vbNullString Then
    'by ordinal :
    For i = 1 To UBound(exeVB6_APIDEF)
        If exeVB6_APIDEF(i).Ordinal = inOrdinal Then
            VBFunction_Description = exeVB6_APIDEF(i).uDescr
            outRName = exeVB6_APIDEF(i).uName
            Exit Function
        End If
    Next i

Else
    'by name:
   
    For i = 1 To UBound(exeVB6_APIDEF)
        If exeVB6_APIDEF(i).uName = inAPIname Then
            VBFunction_Description = exeVB6_APIDEF(i).uDescr
            
            Exit Function
        End If
    Next i
End If

VBFunction_Description = "Error API incorrect or not present in msvbvm60.dll"

End Function



