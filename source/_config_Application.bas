Attribute VB_Name = "_config_Application"
'
'############################################################################
'##                                                                        ##
'##  Individuell gestaltete Config-Module nicht in das Repositiory laden!  ##
'##                                                                        ##
'############################################################################
'
'---------------------------------------------------------------------------------------
' Modul: _config_Application (Beispiel)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Beispiel für Anwendungskonfiguration
' </summary>
' <remarks>
' Indiviuell gestaltete Config-Module nicht in das Repositiory laden.
' </remarks>
'\ingroup base
'**/
'<codelib>
'  <license>_codelib/license.bas</license>
'  <use>base/modApplication.bas</use>
'  <use>base/ApplicationHandler.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
' Das Modul _config_Application wird vom Import-Assistenden nicht überschrieben.
' Sollte eine neues _config_Application-Modul geladen werden,
' ist das alte zuvor umzubennen oder zu löschen.
'
'
' Nicht vergessen: USELOCALIZATION = 1 als Complier-Arg. in Projekteigenschaft einstellen
'

Option Compare Text
Option Explicit
Option Private Module

'/** \addtogroup base
'@{ **/

Private Const APPLICATION_VERSION As String = DCW_Version

#Const USE_CLASS_APPLICATIONHANDLER_APPFILE = 1
#Const USE_CLASS_APPLICATIONHANDLER_VERSION = 1

Private Const APPLICATION_NAME As String = DCW_AppName
Private Const APPLICATION_FULLNAME As String = DCW_AppFullName
Private Const APPLICATION_TITLE As String = APPLICATION_FULLNAME
Private Const APPLICATION_ICONFILE As String = APPLICATION_NAME & ".ico"


Public Const APPLICATION_DOWNLOADSOURCE As String = "https://access.joposol.com/downloads/tools/download/19-tools/111-dbmsconnectionwizard"
Private Const APPLICATION_DOWNLOAD_FOLDER As String = "http://access-codelib.net/download/addins/"
Private Const APPLICATION_DOWNLOAD_VERSIONXMLFILE As String = APPLICATION_DOWNLOAD_FOLDER & "DbmsConnectionWizard.xml"

Public Const APPLICATION_FILTERCODEMODULE_USEVBCOMPONENTSIMPORT As Boolean = True

#If USE_GLOBAL_ERRORHANDLER Then
Const m_DefaultErrorHandlerMode = ACLibErrorHandlerMode.aclibErrMsgBox
#End If

#Const USE_EXTENSIONS = True
#If USE_EXTENSIONS = True Then
Private m_Extensions As ApplicationHandler_ExtensionCollection
#End If

Private Const ApplicationStartFormName As String = "frmConfig_DBMS"

'---------------------------------------------------------------------------------------
' Sub: InitConfig
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Konfigurationseinstellungen initialisieren
' </summary>
' <param name="oCurrentAppHandler">Möglichkeit einer Referenzübergabe, damit nicht CurrentApplication genutzt werden muss</param>
' <returns></returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub InitConfig(Optional oCurrentAppHandler As ApplicationHandler = Nothing)

'----------------------------------------------------------------------------
' Fehlerbehandlung
'
#If USE_GLOBAL_ERRORHANDLER Then
   modErrorHandler.DefaultErrorHandlerMode = m_DefaultErrorHandlerMode
#End If

'----------------------------------------------------------------------------
' Globale Variablen einstellen
'
   'SqlTools.SQL_DATEFORMAT = "\#yyyy-mm-dd\#" 'JET-SQL
   
'----------------------------------------------------------------------------
' Anwendungsinstanz einstellen
'
   If oCurrentAppHandler Is Nothing Then
      Set oCurrentAppHandler = CurrentApplication
   End If

   With oCurrentAppHandler
   
      'Zur Sicherheit AccDb einstellen
      Set .AppDb = CodeDb 'muss auf CodeDb zeigen,
                          'da diese Anwendung als Add-In verwendet wird
   
      'Anwendungsname
      .ApplicationName = APPLICATION_NAME
      .ApplicationFullName = APPLICATION_FULLNAME
      .ApplicationTitle = APPLICATION_TITLE

      'Version
      .Version = APPLICATION_VERSION

      ' Formular, das am Ende von CurrentApplication.Start aufgerufen wird
      .ApplicationStartFormName = ApplicationStartFormName

   End With
   
   
'----------------------------------------------------------------------------
' Erweiterung: ...
'
'----------------------------------------------------------------------------
' Erweiterung: AppFile
'
#If USE_EXTENSIONS = True Then

   Set m_Extensions = New ApplicationHandler_ExtensionCollection
   With m_Extensions
      Set .ApplicationHandler = oCurrentAppHandler
     
   ' load extensions
      .Add New ApplicationHandler_AppFile
      
      Dim AppHdlVersion As ApplicationHandler_Version
      Set AppHdlVersion = New ApplicationHandler_Version
      .Add AppHdlVersion
      AppHdlVersion.XmlVersionCheckFile = APPLICATION_DOWNLOAD_VERSIONXMLFILE
     
   End With
   
#End If

   
End Sub

'############################################################################
'
' Funktionen für die Anwendungswartung
' (werden nur im Anwendungsentwurf benötigt)
'
'----------------------------------------------------------------------------
' Hilfsfunktion zum Speichern von Dateien in die lokale AppFile-Tabelle
'----------------------------------------------------------------------------
Private Sub SetAppFiles()
   Call CurrentApplication.Extensions("AppFile").SaveAppFile("AppIcon", CodeProject.Path & "\" & APPLICATION_ICONFILE)
End Sub


'/** @} **/ '<-- Ende der Doxygen-Gruppen-Zuordnung
