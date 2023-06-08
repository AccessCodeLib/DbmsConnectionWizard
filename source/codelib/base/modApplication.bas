Attribute VB_Name = "modApplication"
Attribute VB_Description = "Standard-Prozeduren f�r die Arbeit mit ApplicationHandler"
'---------------------------------------------------------------------------------------
' Package: base.modApplication
'---------------------------------------------------------------------------------------
'
' Standard procedures for working with ApplicationHandler
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/modApplication.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/_config_Application.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

' Instance of the main control
Private m_ApplicationHandler As ApplicationHandler

'---------------------------------------------------------------------------------------
' Property: CurrentApplication
'---------------------------------------------------------------------------------------
'
' Property f�r ApplicationHandler-Instanz (diese Property im Code verwenden)
'
'---------------------------------------------------------------------------------------
Public Property Get CurrentApplication() As ApplicationHandler
   If m_ApplicationHandler Is Nothing Then
      InitApplication
   End If
   Set CurrentApplication = m_ApplicationHandler
End Property

'---------------------------------------------------------------------------------------
' Sub: TraceLog
'---------------------------------------------------------------------------------------
'
' TraceLog
'
' Parameters:
'     Msg
'     Args
'
'---------------------------------------------------------------------------------------
Public Sub TraceLog(ByRef Msg As String, ParamArray Args() As Variant)
   CurrentApplication.WriteLog Msg, ApplicationHandlerLogType.AppLogType_Tracing, Args
End Sub

Private Sub InitApplication()

   ' Hauptinstanz erzeugen
   Set m_ApplicationHandler = New ApplicationHandler
   
   'Einstellungen initialisieren
   Call InitConfig(m_ApplicationHandler)

End Sub


'---------------------------------------------------------------------------------------
' Sub: DisposeCurrentApplicationHandler
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Instanz von ApplicationHandler und den Erweiterungen zerst�ren
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub DisposeCurrentApplicationHandler()

   Dim CheckCnt As Long, MaxCnt As Long

On Error Resume Next
   
   If Not m_ApplicationHandler Is Nothing Then
      m_ApplicationHandler.Dispose
   End If
   
   Set m_ApplicationHandler = Nothing
   
End Sub


'---------------------------------------------------------------------------------------
'
' Hilfsprozeduren
Public Sub WriteApplicationLogEntry(ByVal Msg As String, _
           Optional LogType As ApplicationHandlerLogType, _
           Optional ByVal Args As Variant)
   CurrentApplication.WriteLog Msg, LogType, Args
End Sub

Public Property Get PublicPath() As String
   PublicPath = CurrentApplication.PublicPath
End Property
