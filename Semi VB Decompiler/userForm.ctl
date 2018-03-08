VERSION 5.00
Begin VB.UserControl userForm 
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3150
   ScaleHeight     =   1995
   ScaleWidth      =   3150
End
Attribute VB_Name = "userForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Paint()
    Dim x As Long
    Dim y As Long
    ForeColor = vbBlack
    Cls
    For x = 0 To Width Step 32
        For y = 0 To Height Step 32
            PSet (x, y), vbBlack
        Next y
        DoEvents
    Next x
    
End Sub

