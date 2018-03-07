Attribute VB_Name = "modWindowSpy"
' *********************************************************************
'  Copyright ©2003 Karl E. Peterson, All Rights Reserved
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

' Win32 APIs...
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Const PS_INSIDEFRAME As Long = 6
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const SM_CXBORDER As Long = 5
Private Const SM_CYBORDER As Long = 6
Private Const SM_CXFRAME As Long = 32
Private Const SM_CYFRAME As Long = 33
Private Const NULL_BRUSH As Long = 5
Private Const NULL_PEN As Long = 8
Private Const R2_NOT As Long = 6
Private Const HS_DIAGCROSS As Long = 5
Private Const CTLCOLOR_STATIC As Long = 6

' Region Flags
Private Const ERRORAPI As Long = 0
Private Const NULLREGION As Long = 1
Private Const SIMPLEREGION As Long = 2
Private Const COMPLEXREGION As Long = 3

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

' The following procedure was inspired by the work of Alex Feinman, and
' I'd like to thank him for offering suggestions on how to do this!
Public Sub FrameWindow(ByVal hWnd As Long, Optional PenWidth As Long = 3)
   Dim hDC As Long
   Dim hRgn As Long
   Dim hPen As Long
   Dim hOldPen As Long
   Dim hBrush As Long
   Dim hOldBrush As Long
   Dim OldMixMode As Long
   Dim cxFrame As Long
   Dim cyFrame As Long
   Dim r As RECT
   
   If IsWindow(hWnd) Then
      hDC = GetWindowDC(hWnd)
      hRgn = CreateRectRgn(0, 0, 0, 0)
      hPen = CreatePen(PS_INSIDEFRAME, _
                       GetSystemMetrics(SM_CXBORDER) * PenWidth, _
                       RGB(0, 0, 0))
      
      hOldPen = SelectObject(hDC, hPen)
      hOldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
      OldMixMode = SetROP2(hDC, R2_NOT)
      
      If GetWindowRgn(hWnd, hRgn) <> ERRORAPI Then
         hBrush = CreateHatchBrush(HS_DIAGCROSS, GetSysColor(CTLCOLOR_STATIC))
         Call FrameRgn(hDC, hRgn, hBrush, _
                       GetSystemMetrics(SM_CXBORDER) * PenWidth, _
                       GetSystemMetrics(SM_CYBORDER) * PenWidth)
   
      Else
         cxFrame = GetSystemMetrics(SM_CXFRAME)
         cyFrame = GetSystemMetrics(SM_CYFRAME)
         Call GetWindowRect(hWnd, r)
         
         If IsZoomed(hWnd) Then
            Call Rectangle(hDC, cxFrame, cyFrame, _
                           GetSystemMetrics(SM_CXSCREEN) + cxFrame, _
                           GetSystemMetrics(SM_CYSCREEN) + cyFrame)
         Else
            Call Rectangle(hDC, 0, 0, r.Right - r.Left, r.Bottom - r.Top)
         End If
      End If
      
   '   // cleanup....
      Call SelectObject(hDC, hOldPen)
      Call SelectObject(hDC, hOldBrush)
      Call SetROP2(hDC, OldMixMode)
      Call DeleteObject(hPen)
      Call DeleteObject(hBrush)
      Call DeleteObject(hRgn)
      Call ReleaseDC(hWnd, hDC)
   End If
End Sub


