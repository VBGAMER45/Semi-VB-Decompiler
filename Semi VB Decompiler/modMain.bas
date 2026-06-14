Attribute VB_Name = "modMain"
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32" ()

Sub Main()
    InitCommonControls
    If HasCommandLine() Then
        'Headless mode: process the file and exit, no UI.
        Call RunCli
    Else
        frmMain.Show
    End If
End Sub
