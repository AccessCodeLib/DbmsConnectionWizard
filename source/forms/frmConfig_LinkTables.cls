Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Const m_Caption As String = "Config Linked Tables"
Private Const m_DoEventCounter As Long = 1 'Anzahl der zu aktualisierenden Tabellen bis Event "DbObjectRelinked" bzw. "DbObjectDeleted" ausgelöst wird

Private WithEvents m_OdbcHandler As OdbcHandler
Attribute m_OdbcHandler.VB_VarHelpID = -1
Private m_CalledByWizard As Boolean

Private m_OdbcHandlerEventBridge As OdbcHandlerEventBridge
Private WithEvents m_L10nDict As L10nDict
Attribute m_L10nDict.VB_VarHelpID = -1

Private Sub cbSelectAll_AfterUpdate()

   Dim bolSelectAll As Boolean
On Error GoTo HandleErr

   bolSelectAll = (Me.cbSelectAll.Value = True)

   ListEnableMode = Not bolSelectAll

   If Not bolSelectAll Then
      listLinkedTableNames
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub


Private Property Let ListEnableMode(ByVal bEnabled As Boolean)

   Me.lstLinkedTabNames.Enabled = bEnabled

   If bEnabled Then
      Me.labFilter.ForeColor = &H7D491F
   Else
      Me.labFilter.ForeColor = &HDED1C6
      If Nz(Me.cbxFilterObjectType.Value, 3) <> 3 Or Len(Me.txtFilterObjectName.Value) > 0 Then
         Me.cbxFilterObjectType.Value = 3
         Me.txtFilterObjectName.Value = Null
         listLinkedTableNames
      End If
   End If

   Me.cbxFilterObjectType.Enabled = bEnabled
   Me.txtFilterObjectName.Enabled = bEnabled
   Me.cmdUseFilter.Enabled = bEnabled

   '
End Property

Private Sub cbxFilterObjectType_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo HandleErr

   If KeyCode = vbKeyReturn Then
      listLinkedTableNames
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub cmdDeleteLinkedTables_Click()

On Error GoTo HandleErr

   If Me.cbSelectAll.Value = True Then
      If MsgBox("Should all linked tables be deleted?" & vbNewLine & _
                "(The tables are removed only in the frontend.)", vbYesNo) = vbYes Then
         StartDeleteOdbcTableDefs
      End If
   Else
       If MsgBox("Should the selected tables be deleted?" & vbNewLine & _
                "(The tables are removed only in the frontend.)", vbYesNo) = vbYes Then
         DeleteSelectedTabs
      End If
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub cmdLinkTables_Click()

On Error GoTo HandleErr

   Select Case ogLinkMode.Value
      Case 1 'manuell
         LinkTextboxTable
      Case 2 'listbox
         LinkSelectedTables
      Case Else
         'nicht möglich
   End Select

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub cmdReconnect_Click()
On Error GoTo HandleErr

   StartReconnectOdbc

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub cmdUseFilter_Click()
On Error GoTo HandleErr

   listLinkedTableNames

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub Form_Load()
On Error GoTo HandleErr

   Me.Section(acDetail).BackColor = dcwBackgroundForm
   listLinkedTableNames
   Me.Caption = m_Caption

   If Len(Me.OpenArgs) > 0 Then
      m_CalledByWizard = True
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub Form_Open(Cancel As Integer)
   L10n.TranslateControls Me.Controls
   Set m_L10nDict = L10n
End Sub

Private Sub Form_Timer()
On Error Resume Next
   Me.labCompletionInfo.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   If Not m_CalledByWizard Then
      DisposeWizard
   End If
   Set m_L10nDict = Nothing
End Sub

Private Sub m_OdbcHandler_DbObjectRelinked(ByVal LastDbObjectName As String, ByVal lCounter As Long)
On Error GoTo HandleErr

   Me.txtLastObjectName.Value = LastDbObjectName
   Me.txtObjectCounter.Value = lCounter
   DoEvents 'erforderlich, damit Änderungen sichtbar werden

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub


Private Sub m_OdbcHandler_DbObjectDeleted(ByVal LastDbObjectName As String, ByVal lCounter As Long)
On Error GoTo HandleErr

   Me.txtLastObjectName.Value = LastDbObjectName
   Me.txtObjectCounter.Value = lCounter
   DoEvents

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub ogLinkMode_AfterUpdate()
On Error GoTo HandleErr

   RefreshDbmsTableNamesList

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub RefreshDbmsTableNamesList()
   Dim lngLinkMode As Long

   lngLinkMode = Nz(ogLinkMode.Value, 0)
   Me.txtTabName.Enabled = (lngLinkMode = 1)
   Me.lstTabNames.Enabled = (lngLinkMode = 2)

   Select Case lngLinkMode
      Case 1 'manuell
         Me.txtTabName.SetFocus
      Case 2 'listbox
         listDbmsTableNames
      Case Else
         'nicht möglich
   End Select
End Sub

Private Sub listDbmsTableNames()

   Const conQdfNameSourceTableList As String = "usys_DbmsTableNames"
   Dim db As DAO.Database
   Dim qdf As DAO.QueryDef
   Dim strDbmsName As String
   Dim strOdbcConnectionString As String
   Dim strTableListSqlText As String
   Dim strSQL As String

   strDbmsName = CurrentConnectionInfo.DbmsName
   strTableListSqlText = CurrentConnectionInfo.TableListSqlText
   If Not CheckListDbmsTableMode() Then
      InsertTablesFromAdoSchemaTables
      Exit Sub
   End If

   strOdbcConnectionString = CurrentConnectionInfo.OdbcConnectionString

   Set db = CodeDb 'AddIn zum Speichern des QDF verwenden
   If QueryDefExists(conQdfNameSourceTableList, db) Then
      Set qdf = db.QueryDefs(conQdfNameSourceTableList)
   Else
      Set qdf = db.CreateQueryDef(conQdfNameSourceTableList)
   End If
   qdf.SQL = "select 1 as x"
   If StrComp(strDbmsName, "MSJET", vbTextCompare) = 0 Then
      qdf.Connect = vbNullString
   Else
      qdf.Connect = strOdbcConnectionString
   End If
   qdf.SQL = strTableListSqlText
   qdf.ReturnsRecords = True


   Dim strFilter As String
   strFilter = GetDbmsTableListFilter("T.")

   strSQL = "select NULL AS TABLE_CATALOG, T.TABLE_SCHEMA, T.TABLE_NAME from " & conQdfNameSourceTableList & " AS T" & _
            " left join (select ForeignName FROM [" & CurrentDb.Name & "].MSysObjects WHERE Type = 4) O ON O.ForeignName = T.TabKey" & _
            " where O.ForeignName IS NULL"

   If Len(strFilter) > 0 Then
      strSQL = strSQL & " AND " & strFilter
   End If

   Me.lstTabNames.RowSource = strSQL

End Sub

Private Sub InsertTablesFromAdoSchemaTables()
   Dim cnn As ADODB.Connection
   Dim rst As ADODB.Recordset

   Set cnn = New ADODB.Connection
   cnn.ConnectionString = CurrentConnectionInfo.OledbConnectionString
   cnn.CursorLocation = adUseClient
   cnn.Open

   'TABLE_CATALOG TABLE_SCHEMA  TABLE_NAME    TABLE_TYPE    TABLE_GUID    DESCRIPTION   TABLE_PROPID  DATE_CREATED  DATE_MODIFIED
   Set rst = cnn.OpenSchema(adSchemaTables, _
                            Array(Empty, Empty, Empty, Empty))
   Set rst.ActiveConnection = Nothing
   cnn.Close
   Set cnn = Nothing

   Dim strFilter As String
   strFilter = GetDbmsTableListFilter
   If Len(strFilter) > 0 Then
      rst.Filter = strFilter
   End If

   Set Me.lstTabNames.Recordset = rst

End Sub

Private Function GetDbmsTableListFilter(Optional SourcePrefix As String = vbNullString) As String

   Dim strFilter As String, strTempFilter As String
   strTempFilter = Nz(Me.txtFiltSchema.Value, vbNullString)
   If Len(strTempFilter) > 0 Then
      If InStr(1, strTempFilter, "*") > 0 Or InStr(1, strTempFilter, "%") > 0 Then
         strFilter = " " & SourcePrefix & "TABLE_SCHEMA like " & SqlTools.TextToSqlText(Replace(strTempFilter, "%", "*"))
      Else
         strFilter = " " & SourcePrefix & "TABLE_SCHEMA = " & SqlTools.TextToSqlText(strTempFilter)
      End If
   End If

   strTempFilter = Nz(Me.txtFiltTabname.Value, vbNullString)
   If Len(strTempFilter) > 0 Then
      If Len(strFilter) > 0 Then
         strFilter = strFilter & " AND "
      End If
      If InStr(1, strTempFilter, "*") > 0 Or InStr(1, strTempFilter, "%") > 0 Then
         strFilter = strFilter & " " & SourcePrefix & "TABLE_NAME like " & SqlTools.TextToSqlText(Replace(strTempFilter, "%", "*"))
      Else
         strFilter = strFilter & " " & SourcePrefix & "TABLE_NAME = " & SqlTools.TextToSqlText(strTempFilter)
      End If
   End If

   GetDbmsTableListFilter = strFilter

End Function

Private Sub listLinkedTableNames()

   Dim strSQL As String
   Dim strFilter As String
   Dim lngObjectFilter As Long

   If Len(Me.txtFilterObjectName.Value) > 0 Then
      strFilter = " AND O.Name like " & SqlTools.TextToSqlText(Me.txtFilterObjectName.Value)
   End If

   lngObjectFilter = Nz(Me.cbxFilterObjectType.Value, 3)

   If (lngObjectFilter And 1) = 1 Then

      'Tabellen
      strSQL = "SELECT 'T' AS [*], O.Name AS TABLE_NAME, O.ForeignName" & _
               " FROM  [" & CurrentDb.Name & "].MSysObjects O" & _
               " WHERE O.ForeignName > '' AND O.Type IN (4, 6)"

      If Len(strFilter) > 0 Then
         strSQL = strSQL & strFilter
      End If
   End If

   If (lngObjectFilter And 2) = 2 Then

      If lngObjectFilter = 3 Then
         strSQL = strSQL & " UNION ALL "
      End If

      '+PT-Abfragen
      strSQL = strSQL & _
               "SELECT 'Q' AS [*], O.Name AS TABLE_NAME, Q.Expression AS ForeignName" & _
               " FROM [" & CurrentDb.Name & "].MSysObjects O INNER JOIN [" & CurrentDb.Name & "].MSysQueries Q ON O.Id = Q.ObjectId" & _
               " WHERE O.Type = 5 And Q.Attribute = 1 And Q.Name1 Like 'ODBC;*'"

      If Len(strFilter) > 0 Then
         strSQL = strSQL & strFilter
      End If

   End If

   strSQL = strSQL & " ORDER BY 1 desc, 2 asc"

   Me.lstLinkedTabNames.RowSource = strSQL

End Sub

Private Sub LinkTextboxTable()

   Dim strSourceTabName As String
   Dim strTdfName As String

   strSourceTabName = Nz(Me.txtTabName, vbNullString)
   If Len(strSourceTabName) = 0 Then
      MsgBox "Please enter a table name."
      Exit Sub
   End If

   strTdfName = strSourceTabName
   DbCon.ODBC.LinkTable strSourceTabName, strTdfName, CurrentConnectionInfo.SavePassword
   Application.RefreshDatabaseWindow

End Sub

Private Sub LinkSelectedTables()

   Dim varItm As Variant
   Dim lb As ListBox
   Dim strSourceTabName As String
   Dim strSourceTabSchema As String
   Dim strTdfName As String

   DbCon.CurrentDb.TableDefs.Refresh 'AnwendungsDB verwenden

   Set lb = Me.lstTabNames
   For Each varItm In lb.ItemsSelected
      strSourceTabName = Nz(lb.Column(2, varItm), vbNullString)
      strSourceTabSchema = Nz(lb.Column(1, varItm), vbNullString)
      If Len(strSourceTabSchema) > 0 Then strSourceTabSchema = strSourceTabSchema & "."
      strTdfName = strSourceTabName
      DbCon.ODBC.LinkTable strSourceTabSchema & strSourceTabName, strTdfName, CurrentConnectionInfo.SavePassword
      lb.Selected(varItm) = False
   Next

   lb.Requery

   DbCon.CurrentDb.TableDefs.Refresh
   Application.RefreshDatabaseWindow

End Sub

Private Sub StartReconnectOdbc()
   If Me.cbSelectAll.Value = True Then
      reconnectAllTab
   Else
      ReconnectSelectedTab
   End If
End Sub

Private Sub reconnectAllTab()
   Me.labObjectCounter.Caption = L10n.Text("Tables/queries were linked")
   Me.labLastObjectName.Caption = L10n.Text("last linked:")
   Me.txtObjectCounter.Value = Null
   Set m_OdbcHandler = DbCon.ODBC

   TryConnectClientOdbcHandlerEventBridge m_OdbcHandler

   m_OdbcHandler.RelinkTablesAndQueries CurrentConnectionInfo.SavePassword, m_DoEventCounter ' 10 statt 1 würde nur jede 10. Aktualisierung anzeigen und den Ablauf beschleunigen
   Set m_OdbcHandler = Nothing
   ShowCompleteMessage
End Sub

Private Sub TryConnectClientOdbcHandlerEventBridge(ByVal OdbcHandlerRef As OdbcHandler)

   Dim HookObserver As Object

With New ErrorTrappingObserver

   .ErrorTrapping = BreakOnUnhandledErrors

On Error Resume Next

   Set HookObserver = Application.Run("GetDbmsConnectionWizardHookObserver")

On Error GoTo 0

End With

   If HookObserver Is Nothing Then
      Exit Sub
   End If

   Set m_OdbcHandlerEventBridge = New OdbcHandlerEventBridge
   m_OdbcHandlerEventBridge.Init OdbcHandlerRef, HookObserver


End Sub

Private Sub ReconnectSelectedTab()

   Dim varItm As Variant
   Dim lb As ListBox
   Dim strTdfQdfName As String
   Dim lngCnt As Long
   Dim strOdbcConnectionString As String
   Dim bolSavePassword As Boolean

   DbCon.CurrentDb.TableDefs.Refresh 'AnwendungsDB verwenden

   Me.labObjectCounter.Caption = L10n.Text("Tables/queries were linked")
   Me.labLastObjectName.Caption = L10n.Text("last linked:")
   Me.txtObjectCounter.Value = Null

   bolSavePassword = CurrentConnectionInfo.SavePassword
   strOdbcConnectionString = CurrentConnectionInfo.OdbcConnectionString(Not bolSavePassword)

   Set lb = Me.lstLinkedTabNames
   For Each varItm In lb.ItemsSelected
      strTdfQdfName = Nz(lb.Column(1, varItm), vbNullString)
   'On Error Resume Next
      If Nz(lb.Column(0, varItm), vbNullString) = "Q" Then
         DbCon.ODBC.RelinkPassThroughQuery strTdfQdfName, bolSavePassword
      Else
         DbCon.ODBC.RelinkTable strTdfQdfName, bolSavePassword
      End If
   'On Error GoTo 0
      lngCnt = lngCnt + 1
      Me.txtObjectCounter.Value = lngCnt
      Me.txtLastObjectName.Value = strTdfQdfName
      lb.Selected(varItm) = False
      DoEvents
   Next

   lb.Requery

   DbCon.CurrentDb.TableDefs.Refresh
   Application.RefreshDatabaseWindow

   ShowCompleteMessage

End Sub

Private Sub StartDeleteOdbcTableDefs()
   Me.txtObjectCounter.Value = Null
   Me.labObjectCounter.Caption = L10n.Text("linked tables were removed")
   Me.labLastObjectName.Caption = L10n.Text("last removed:")
   Set m_OdbcHandler = DbCon.ODBC
   m_OdbcHandler.DeleteOdbcTableDefs m_DoEventCounter
   Set m_OdbcHandler = Nothing
   Application.RefreshDatabaseWindow
   listLinkedTableNames
End Sub

Private Sub DeleteSelectedTabs()

   Dim db As DAO.Database
   Dim varItm As Variant
   Dim lb As ListBox
   Dim strTdfQdfName As String
   Dim lngCnt As Long

   Set db = CurrentDb
   db.TableDefs.Refresh 'AnwendungsDB verwenden

   Me.labObjectCounter.Caption = L10n.Text("linked tables were removed")
   Me.labLastObjectName.Caption = L10n.Text("last removed:")
   Me.txtObjectCounter.Value = Null

   Set lb = Me.lstLinkedTabNames
   For Each varItm In lb.ItemsSelected
      strTdfQdfName = Nz(lb.Column(1, varItm), vbNullString)
      If Nz(lb.Column(1, varItm), vbNullString) = "Q" Then
         db.QueryDefs.Delete strTdfQdfName
      Else
         db.TableDefs.Delete strTdfQdfName
      End If

      lngCnt = lngCnt + 1
      Me.txtObjectCounter.Value = lngCnt
      Me.txtLastObjectName.Value = strTdfQdfName
      lb.Selected(varItm) = False
   Next

   lb.Requery

   db.TableDefs.Refresh '
   Application.RefreshDatabaseWindow

   listLinkedTableNames

End Sub


Private Sub tabLinkTables_Change()
On Error GoTo HandleErr

   Select Case Me.tabLinkTables.Value
      Case 0
         listLinkedTableNames
      Case 1
         If Me.ogLinkMode = 2 Then
            listDbmsTableNames
         End If
      Case Else
         '
   End Select

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Function CheckListDbmsTableMode() As Boolean
   If Len(CurrentConnectionInfo.TableListSqlText) = 0 Then
      CheckListDbmsTableMode = False
   Else
      CheckListDbmsTableMode = True
   End If
End Function

Private Sub txtFilterObjectName_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo HandleErr

   If KeyCode = vbKeyReturn Then
      listLinkedTableNames
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub txtFiltSchema_AfterUpdate()
On Error GoTo HandleErr

   RefreshDbmsTableNamesList

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub txtFiltTabname_AfterUpdate()
On Error GoTo HandleErr

   RefreshDbmsTableNamesList

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub ShowCompleteMessage()
   Me.txtLastObjectName.Value = Me.txtLastObjectName.Value & "."
   Me.labCompletionInfo.Visible = True
   Me.TimerInterval = 2000
End Sub

Private Sub m_L10nDict_LanguageChanged()
   L10n.TranslateControls Me.Controls
End Sub
