Attribute VB_Name = "defDbConnection"
'---------------------------------------------------------------------------------------
' Package: data.defDbConnection
'---------------------------------------------------------------------------------------
'
' Set of global types, enums etc. for the DbConnection classes
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/defDbConnection.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
'
' Type + Enums
'

'---------------------------------------------------------------------------------------
' Type: DbmsConnectionStrings
'---------------------------------------------------------------------------------------
'
' Type for the simultaneous transfer of the connection strings
'
' <list type="table">
'   <item><term>OledbConnectionString</term><description>OLEDB connectionstring for ADODB connection</description></item>
'   <item><term>OdbcConnectionString</term><description>ODBC connectionstring for DAO connection</description></item>
'   <item><term>DatabaseFile</term><description>Database file name incl. path if a file backend is used.</description></item>
' </list>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Type DbmsConnectionStrings
   OledbConnectionString As String
   OdbcConnectionString As String
   DatabaseFile As String
End Type


'---------------------------------------------------------------------------------------
' Enum: DbmsConnectionModes
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DbmsConnectionModes
' </summary>
' <list type="table">
'   <item><term>DMBS_DSNless (1)</term><description>without DSN</description></item>
'   <item><term>DMBS_DSN (2)</term><description>with DSN</description></item>
'   <item><term>DBMS_USERDEF (128)</term><description>User-defined connection strings</description></item>
' </list>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Enum DbmsConnectionModes
   DMBS_DSNless = 1   'ohne DSN
   DMBS_DSN = 2       'mit DSN
   DBMS_USERDEF = 128 'benutzerdefinierte Connectionstrings
End Enum
