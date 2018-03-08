VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************
'clsConsole
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************
Option Explicit

Private strConsoleText As String
Public Sub Clear()
    strConsoleText = ""
End Sub
Public Sub WriteLine(ByVal strText As String)
    strConsoleText = strConsoleText & strText & vbCrLf
End Sub
Public Sub WriteA(ByVal strText As String)
    strConsoleText = strConsoleText & strText
End Sub
Public Sub SaveConsoleToFile(strFileName As String)
On Error GoTo BadFile
    Dim f As Long
    f = FreeFile
    Open strFileName For Output As #f
        Print #f, strConsoleText
    Close #f
Exit Sub
BadFile:
    MsgBox "Error_clsConsole_SaveConsoleToFile: " & err.Description, vbCritical
Exit Sub
End Sub

