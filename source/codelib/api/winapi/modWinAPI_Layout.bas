Attribute VB_Name = "modWinAPI_Layout"
Attribute VB_Description = "WinAPI-Funktionen zur Layoutgestaltung"
'---------------------------------------------------------------------------------------
' Module: modWinAPI_Layout
'---------------------------------------------------------------------------------------
'/**
' <summary>
' WinAPI-Funktionen zur Layoutgestaltung
' </summary>
' <remarks>
' </remarks>
'\ingroup WinAPI
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/modWinAPI_Layout.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Declare Function CreateSolidBrush _
      Lib "gdi32.dll" ( _
      ByVal CrColor As Long _
      ) As Long

Private Declare Function RedrawWindow _
      Lib "user32" ( _
      ByVal Hwnd As Long, _
      LprcUpdate As Any, _
      ByVal HrgnUpdate As Long, _
      ByVal FuRedraw As Long _
      ) As Long

Private Declare Function SetClassLong _
      Lib "user32.dll" _
      Alias "SetClassLongA" ( _
      ByVal Hwnd As Long, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As Long _
      ) As Long

Private Const GCL_HBRBACKGROUND As Long = -10
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4

'--------------------------------------
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Const HWND_DESKTOP As Long = 0
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Const SM_CXVSCROLL As Long = 2


'---------------------------------------------------------------------------------------
' Sub: SetBackColor (Josef P�tzl, 2010-04-19)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hintergrundfarbe eines Fensters einstellen
' </summary>
' <param name="Hwnd">Fenster-Handle</param>
' <param name="Color">Farbnummer</param>
' <returns></returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub SetBackColor(ByVal Hwnd As Long, ByVal Color As Long)
  
   Dim NewBrush As Long
   
   'Brush erzeugen
   NewBrush = CreateSolidBrush(Color)
   'Brush zuweisen
   SetClassLong Hwnd, GCL_HBRBACKGROUND, NewBrush
   'Fenster neuzeichnen (gesamtes Fenster inkl. Background)
   RedrawWindow Hwnd, ByVal 0&, ByVal 0&, RDW_INVALIDATE Or RDW_ERASE

End Sub

'---------------------------------------------------------------------------------------
' Function: TwipsPerPixelX
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Breite eines Pixels in twips
' </summary>
' <param name="Param"></param>
' <returns>Single</returns>
' <remarks>
' http://support.microsoft.com/kb/94927/de
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function TwipsPerPixelX() As Single
   Dim lngDC As Long
   lngDC = GetDC(HWND_DESKTOP)
   TwipsPerPixelX = 1440& / GetDeviceCaps(lngDC, LOGPIXELSX)
   ReleaseDC HWND_DESKTOP, lngDC
End Function

'---------------------------------------------------------------------------------------
' Function: TwipsPerPixelY
'---------------------------------------------------------------------------------------
'/**
' <summary>
' H�he eines Pixels in twips
' </summary>
' <returns>Single</returns>
' <remarks>
' http://support.microsoft.com/kb/94927/de
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function TwipsPerPixelY() As Single
   Dim lngDC As Long
   lngDC = GetDC(HWND_DESKTOP)
   TwipsPerPixelY = 1440& / GetDeviceCaps(lngDC, LOGPIXELSY)
   ReleaseDC HWND_DESKTOP, lngDC
End Function

'---------------------------------------------------------------------------------------
' Function: GetScrollbarWidth
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Breite der Bildlaufleiste
' </summary>
' <param name="Param"></param>
' <returns>Single</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetScrollbarWidth() As Single
   GetScrollbarWidth = GetSystemMetrics(SM_CXVSCROLL) * TwipsPerPixelX
End Function

'---------------------------------------------------------------------------------------
' Function: GetTwipsFromPixel
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Rechnet Pixel in Twips um
' </summary>
' <param name="pixel">Anzahl der Pixel</param>
' <returns>Long</returns>
' <remarks>
' GetTwipsFromPixel = TwipsPerPixelX * pixel
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetTwipsFromPixel(ByVal Pixel As Long) As Long
   GetTwipsFromPixel = TwipsPerPixelX * Pixel
End Function

'---------------------------------------------------------------------------------------
' Function: GetPixelFromTwips
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Rechnet twips in Pixel um
' </summary>
' <param name="twips">Anzahl twips</param>
' <returns>Long</returns>
' <remarks>
'  GetPixelFromTwips = twips / TwipsPerPixelX
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetPixelFromTwips(ByVal Twips As Long) As Long
        GetPixelFromTwips = Twips / TwipsPerPixelX
End Function
