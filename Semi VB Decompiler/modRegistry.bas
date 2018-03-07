Attribute VB_Name = "modRegistry"
'*********************************************
'modRegistry
'Copyright VisualBasicZone.com 2004 - 2005
'Purpose: To get external componet information
'*********************************************
Option Explicit

Const REG_SZ = 1 ' Unicode null terminated string
Const REG_BINARY = 3 ' Free form binary
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Function RegQueryStringValue(ByVal hKey As Long, ByVal strPath, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'retrieve nformation about the key
    Dim hSubKey As Long
    lResult = RegOpenKey(hKey, strPath, hSubKey)
 
    lResult = RegQueryValueEx(hSubKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
  
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            'Create a buffer
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(hSubKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'Remove the unnecessary chr$(0)'s
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hSubKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
    RegCloseKey (hSubKey)
End Function

