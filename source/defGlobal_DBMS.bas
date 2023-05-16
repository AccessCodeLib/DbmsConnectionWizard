Attribute VB_Name = "defGlobal_DBMS"
'---------------------------------------------------------------------------------------
' Modul: defGlobal_DBMS
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Allgemeine Konstanten und Eigenschaften für DBMS-Teile
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'
Option Explicit
Option Compare Text

'---------------------------------------------------------------------------------------
'
' Konstanten
'

'Login-Formular
Public Const g_ApplicationLoginFormName As String = "frmLogin"

'DBMS-Konfiguration
Private Const m_conDbmsConfigFormName As String = vbNullString ' "frmConfig_DBMS" ... vbnullstring = deaktiviert (Kein Einblenden des Konfigurationsfesters bei fehlenden Verbindungseinstellungen)

'---------------------------------------------------------------------------------------
'
' Hilfs-Variablen
'
Public g_TempRef As Object 'Hilfsvariable zum Austausch einer Objektreferenz (Wird für Loginform wegen acDialog benötigt)


'---------------------------------------------------------------------------------------
'
' Hilfs-Prozeduren
'
Private m_DbmsConfigFormName As String


'---------------------------------------------------------------------------------------
' Property: DbmsConfigFormName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Name des Formulars zur DBMS-Konfiguration
' </summary>
' <returns>String</returns>
' <remarks>
' Wird aus
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get DbmsConfigFormName() As String
On Error Resume Next

   If Len(m_DbmsConfigFormName) = 0 Then 'Wert von Konstante
      m_DbmsConfigFormName = m_conDbmsConfigFormName
   End If

   DbmsConfigFormName = m_DbmsConfigFormName

End Property

Public Property Let DbmsConfigFormName(ByVal AppName As String)
On Error Resume Next
    m_DbmsConfigFormName = AppName
End Property
