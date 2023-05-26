Attribute VB_Name = "_initDbConnectionWizard"
'---------------------------------------------------------------------------------------
' Modul: _initApplication
'---------------------------------------------------------------------------------------
'/* *
' <summary>
' Initialisierungsaufruf der Anwendung
' </summary>
' <remarks></remarks>
'* */
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'---------------------------
' Initialisierungsfunktion
'---------------------------
Private Function InitDbmsConnectionWizard() As Boolean
   If Application.CurrentDb Is Nothing Then
      MsgBox "Bitte öffnen Sie zuerst eine Access-Anwendung.", vbCritical
      InitDbmsConnectionWizard = False
      Exit Function
   End If
   InitDbmsConnectionWizard = StartApplication
End Function

Public Function StartDbmsConnectionWizard() As Variant
   If Not InitDbmsConnectionWizard Then Exit Function
   CheckTemplatesDb
   DoCmd.OpenForm DCW_DbmsConfigFormName, acNormal, , , , acWindowNormal
End Function

Public Function StartDbmsConnectionLinkTableWizard() As Variant
   If Not InitDbmsConnectionWizard Then Exit Function
   DoCmd.OpenForm DCW_LinkTablesFormName, acNormal, , , , acWindowNormal
End Function

Public Function CreatePassThroughQuery() As Variant
   If Not InitDbmsConnectionWizard Then Exit Function
   DoCmd.OpenForm DCW_TestSqlFormName, acNormal, , , , acWindowNormal, DCW_RecordsetModes.DCW_DAOPT
End Function

Public Function InsertOdbcConnectionString(ByRef strObjektName As String, ByRef strTextFeldname As String, ByRef strAktuellerWert As String) As Variant
   If Not InitDbmsConnectionWizard Then Exit Function
   Select Case strTextFeldname
      Case "Query"
         strAktuellerWert = CurrentConnectionInfo.OdbcConnectionString
      Case Else
         '
   End Select
   InsertOdbcConnectionString = strAktuellerWert
End Function


Public Sub DisposeWizard()
   DisposeDbConnection
End Sub

'#################################################
'
' Hilfsprozeduren

Private Sub CheckTemplatesDb()

   Dim db As DAO.Database
   Dim tdf As DAO.TableDef
   
On Error Resume Next 'falls etwas nicht klappt, einfach übergehen, dann können nur keine Vorlagen verwendet werden
   
   Set db = CodeDb
   
   If Not TableDefExists("usys_DbmsConnection", db) Then
   
      Set tdf = db.CreateTableDef("usys_DbmsConnection")
      tdf.Connect = ";Database=" & TemplatesDatabaseFile
      tdf.SourceTableName = "usys_DbmsConnection"
      db.TableDefs.Append tdf
      
   ElseIf Len(Dir$(Mid$(db.TableDefs("usys_DbmsConnection").Connect, Len(";Database=") + 1))) = 0 Then
   
      With db.TableDefs("usys_DbmsConnection")
         .Connect = ";Database=" & TemplatesDatabaseFile
         .RefreshLink
      End With
      
   End If
   
   Set db = Nothing

End Sub

Private Property Get TemplatesDatabaseFile() As String
   
   Dim db As DAO.Database
   Dim cnn As ADODB.Connection
   Dim strDbFile As String
   Dim strPath As String

   strPath = Environ("Appdata") & "\DbmsConnectionWizard"
   If Len(Dir$(strPath, vbDirectory)) = 0 Then
      MkDir strPath
   End If
   
   strDbFile = CodeDb.Name
   strDbFile = Mid$(strDbFile, InStrRev(strDbFile, "."))
   If Left$(strDbFile, 5) = ".accd" Then
      strDbFile = ".accdu"
   Else
      strDbFile = ".mdt"
   End If
   strDbFile = strPath & "\ConnectionTemplates" & strDbFile
   
   If Len(Dir$(strDbFile)) = 0 Then

      'Datenbank anlegen
      If CodeDb.Version = "4.0" Then
         Set db = DBEngine.CreateDatabase(strDbFile, dbLangGeneral, dbVersion40)
      Else
         Set db = DBEngine.CreateDatabase(strDbFile, dbLangGeneral)
      End If
      db.Close
      
      'Tabelle erstellen
      Set cnn = New ADODB.Connection
      cnn.ConnectionString = Replace(CodeProject.Connection.ConnectionString, CodeDb.Name, strDbFile)
      cnn.Open
      CreateConnectionTable cnn
      cnn.Close
      Set cnn = Nothing
      
      'Beispiel-Vorlagen einfügen
      Set db = DBEngine.OpenDatabase(strDbFile, False, False)
      db.Execute "insert into usys_DbmsConnection (" & DCW_usys_DbmsConnection_Fields & ") select " & DCW_usys_DbmsConnection_Fields & " from [" & CodeDb.Name & "].usys_DbmsConnection_CopyBase"
      db.Close
      Set db = Nothing

   End If

   TemplatesDatabaseFile = strDbFile

End Property
