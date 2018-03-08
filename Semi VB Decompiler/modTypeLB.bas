Attribute VB_Name = "modTypeLB"
'=========================================================================================
' RTF API Module
'=========================================================================================
' Adapted and Modified By: Marc Cramer
' Published Date: 01/15/2001
' WebSite: MKC Computers at http://www.mkccomputers.com
'=========================================================================================
' Based On: VB Type Library Registration Utility
' WebSite: vbAccelerator at http://www.vbaccelerator.com
'=========================================================================================
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function LoadTypeLib Lib "oleaut32.dll" (pFileName As Byte, pptlib As Object) As Long
Private Declare Function RegisterTypeLib Lib "oleaut32.dll" (ByVal ptlib As Object, szFullPath As Byte, szHelpFile As Byte) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (lpsz As Byte, pclsid As GUID) As Long
Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
'============================================================================================
Public Sub RegisterOLB(strFileName As String)
' find out what to do...register or unregister
  Dim Buffer As String
  Dim Filename As String
  
  Buffer = String$(255, 0)
  GetFileTitle CStr(strFileName), Buffer, Len(Buffer)
  Buffer = Left$(Buffer, InStr(1, Buffer, Chr$(0)) - 1)
  Filename = Buffer


    RegisterMe CStr(strFileName), Filename


End Sub
'============================================================================================
Private Sub RegisterMe(FilePath As String, Filename As String)
' register the type library
  Dim ResultMessage As Long
  Dim TypeLibraryPointer As Object
  Dim TypeLibraryPath() As Byte
  
  TypeLibraryPath = FilePath & vbNullChar
  ResultMessage = LoadTypeLib(TypeLibraryPath(0), TypeLibraryPointer)
  
  If ResultMessage = 0 Then
    ResultMessage = RegisterTypeLib(TypeLibraryPointer, TypeLibraryPath(0), 0)
  End If
  If ResultMessage = 0 Then
    'MsgBox "Successful Registration of Type Library: " & Filename, vbInformation, "Registration Successful"
  Else
    MsgBox "Registration of Type Library: " & Filename & " Unsuccessful", vbInformation, "Registration Unsuccessful"
  End If
End Sub ' RegisterMe(FilePath As String, FileName As String)

