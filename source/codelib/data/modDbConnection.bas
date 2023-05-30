Attribute VB_Name = "modDbConnection"
'---------------------------------------------------------------------------------------
' Package: data.modDbConnection
'---------------------------------------------------------------------------------------
'
' Collection of procedures for the DbConnectionManager class
'
' Remarks:
'     Used to instantiate DbConnectionManager and access the main elements of DbConnectionManager.</remarks>
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/modDbConnection.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>data/DbConnectionManager.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

Private m_DbConnectionManager As DbConnectionManager ' Hauptsteuerung

'---------------------------------------------------------------------------------------
' Property: DbCon
'---------------------------------------------------------------------------------------
'
' Access to DbConnection of the main DbConnectionManager instance
'
' Returns:
'     DbConnectionHandler
'
' Remarks:
'  | Simplifies the call, as the writing of the DbConnectionManager instance variables is omitted.
'  | The first access creates the instance of the DbConnectionManager class if it does not already exist.
'
'---------------------------------------------------------------------------------------
Public Property Get DbCon() As DbConnectionHandler

   If m_DbConnectionManager Is Nothing Then
      Set m_DbConnectionManager = New DbConnectionManager
   End If
   Set DbCon = m_DbConnectionManager.DbConnection

End Property

'---------------------------------------------------------------------------------------
' Property: CurrentConnectionInfo
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Access to DbConnectionInfo of the main DbConnectionManager instance
' </summary>
' <returns>DbConnectionInfo</returns>
' <remarks>
' Simplifies the call, as the writing of the DbConnectionManager instance variables is omitted.
' The first access creates the instance of the DbConnectionManager class if it does not already exist.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CurrentConnectionInfo() As DbConnectionInfo

   If m_DbConnectionManager Is Nothing Then
      Set m_DbConnectionManager = New DbConnectionManager
   End If
   Set CurrentConnectionInfo = m_DbConnectionManager.ConnectionInfo

End Property

'---------------------------------------------------------------------------------------
' Function: CheckConnectionStatus
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Verbindung prüfen <see cref=DbConnectionManager#CheckConnectionStatus>DbConnectionManager.CheckConnectionStatus</see>
' </summary>
' <returns>Boolean: True if success</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CheckConnectionStatus() As Boolean

   If m_DbConnectionManager Is Nothing Then
      Set m_DbConnectionManager = New DbConnectionManager
   End If
   CheckConnectionStatus = m_DbConnectionManager.CheckConnectionStatus

End Function

'---------------------------------------------------------------------------------------
' Sub: DisposeDbConnection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Dispose DbConnection objects (incl. DbConnectionManager)
' </summary>
'**/
'---------------------------------------------------------------------------------------
Public Sub DisposeDbConnection()

On Error Resume Next

   If Not (m_DbConnectionManager Is Nothing) Then
      m_DbConnectionManager.Dispose
   End If
   Set m_DbConnectionManager = Nothing
   
End Sub


'############################
'
' Hilfsfunktionen

Public Function TableDefExists(ByVal sTableDefName As String, Optional ByRef dbs As DAO.Database = Nothing) As Boolean
'Schneller wäre der Zugriff auf MSysObject (select .. from MSysObject where Name = 'Tabellenname' AND Type IN (1, 4, 6)
'Eine weitere Alternative wäre das Auswerten über cnn.OpenSchema(adSchemaTables, ...) ... dann werden allerdings keine verknüpften Tabellen geprüft
   
   Dim tdf As DAO.TableDef

   If dbs Is Nothing Then
      Set dbs = CurrentDb
   End If
   
   dbs.TableDefs.Refresh
   For Each tdf In dbs.TableDefs
      If tdf.Name = sTableDefName Then
         TableDefExists = True
         Exit For
      End If
   Next

End Function

Public Function QueryDefExists(ByVal sQueryDefName As String, Optional ByVal dbs As DAO.Database = Nothing) As Boolean
   
   Dim qdf As DAO.QueryDef

   If dbs Is Nothing Then
      Set dbs = CurrentDb
   End If
   
   dbs.QueryDefs.Refresh
   For Each qdf In dbs.QueryDefs
      If qdf.Name = sQueryDefName Then
         QueryDefExists = True
         Exit For
      End If
   Next

End Function
