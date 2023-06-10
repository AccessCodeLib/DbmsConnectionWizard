Attribute VB_Name = "defDbmsConnectionWizard"
'---------------------------------------------------------------------------------------
' Modul: defGlobalParam
'---------------------------------------------------------------------------------------
'/* *
' <summary>
' Allgemeine Konstanten und Eigenschaften
' </summary>
' <remarks></remarks>
'* */
'---------------------------------------------------------------------------------------
'
Option Explicit
Option Private Module

Public Const DCW_AppName As String = "DbmsConnectionWizard"
Public Const DCW_AppFullName As String = "DBMS Connection Wizard"
Public Const DCW_Version As String = "1.4.4"
Public Const DCW_VersionDate As String = "2023-06-10"

'Formulare
Public Const DCW_LoginFormName As String = "LoginForm"
Public Const DCW_DbmsConfigFormName As String = "frmConfig_DBMS"
Public Const DCW_LinkTablesFormName As String = "frmConfig_LinkTables"
Public Const DCW_TestSqlFormName As String = "frmTest_SQL"

'Datenfelder
Public Const DCW_usys_DbmsConnection_Fields As String = "CID, ActiveConnection, DBMS, dbmsConnectionMode, dbmsOleDbProvider, dbmsOdbcDriver, dbmsServer, dbmsPort, dbmsDatabase, dbmsUseTrustedConnection, dbmsUseLoginForm, dbmsUser, dbmsPwd, dbmsOptionsODBC, dbmsOptionsOLEDB, dbmsDSN, dbmsConStrODBC, dbmsConStrOLEDB, Remarks"


Public Enum DCW_RecordsetModes
   DCW_ADODB = 1  ' ADOD-Recordset
   DCW_DAOBE = 2  ' DAO-Recordset über BE-Database
   DCW_DAOPT = 3  ' DAO-Recordset über PT-Abfrage
End Enum
