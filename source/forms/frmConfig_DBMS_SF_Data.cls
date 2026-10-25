Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Basis-Datenherkunft
Private Const m_DATASOURCE As String = _
   "SELECT C.CID, C.ActiveConnection, C.DBMS, C.dbmsConnectionMode, C.dbmsOleDbProvider, C.dbmsOdbcDriver, C.dbmsServer, C.dbmsPort, C.dbmsDatabase, X.dbmsEnabled, C.dbmsUseTrustedConnection, C.dbmsUseLoginForm, C.dbmsUser, C.dbmsPwd, C.dbmsOptionsODBC, C.dbmsOptionsOLEDB, C.dbmsDSN, C.dbmsConStrODBC, C.dbmsConStrOLEDB, C.Remarks, X.XdbmsConnectionModes, X.XdbmsServer, X.XdbmsPort, X.XdbmsDatabase, X.XdbmsUseLoginForm, X.XdbmsUseTrustedConnection, X.XdbmsUser, X.XdbmsPwd, X.XdbmsOptionsODBC, X.XdbmsOptionsOLEDB, X.XdbmsDSN, X.XdbmsConStrODBC, X.XdbmsConStrOLEDB, X.XdbmsOledbProvider, X.XdbmsOdbcDriver, X.DriverODBC, X.DriverOLEDB" & _
   " FROM usys_DbmsConnection_X AS X RIGHT JOIN usys_DbmsConnection AS C ON X.DBMS = C.DBMS"

Public Event CidValueChanged(ByVal newCID As String)
Public Event CidStateChanged(ByVal currentCID As String)

Private m_OpenModules As Long
Private m_CodeModulImporter As AppFileCodeModulTransfer

Private m_TemplatesForm As Form_frmConfig_DBMS

Private m_TemplateMode As Boolean

Private WithEvents m_L10nDict As L10nDict
Attribute m_L10nDict.VB_VarHelpID = -1

Public Property Let TemplateMode(ByVal tm As Boolean)
   m_TemplateMode = tm
   InitForm
End Property

'Datenquellen für Formular (je nachdem ob Formular für FE-Verbindung bzw. Vorlage verwendet wird):
Private Property Get SourceDb() As DAO.Database
   If m_TemplateMode Then 'DB des Add-In verwenden
      Set SourceDb = CodeDb
   Else 'FE-Databaseinstanz einsetzen
      Set SourceDb = CurrentDb
   End If
End Property

Private Property Get DataSource() As String
   If m_TemplateMode Then 'unveränderte Basis-Datenherkunft verwenden
      DataSource = m_DATASOURCE
   Else 'Basis-Datenherkunft anpassen, damit usys_DbmsConnection_X aus Add-In verwendet wird
      DataSource = Replace(m_DATASOURCE, "usys_DbmsConnection_X", "[" & CodeDb.Name & "].usys_DbmsConnection_X")
   End If
End Property

'Datenquelle für Formular einstellen
Public Sub SetData(ByVal sCID As String)
'sCID: Übergabe der Verbindungskennung

   Dim strSQL As String

   If Len(sCID) > 0 Then
      strSQL = DataSource & " where CID = " & SqlTools.TextToSqlText(sCID)
      Set Me.Recordset = SourceDb.OpenRecordset(strSQL, dbOpenDynaset)
   Else
      SetNoData
   End If

End Sub

Public Sub SetNoData()
   Me.RecordSource = "SELECT NULL AS CID, NULL AS ActiveConnection, NULL AS DBMS, NULL AS dbmsConnectionMode, NULL AS dbmsOleDbProvider, NULL AS dbmsOdbcDriver, NULL AS dbmsServer, NULL AS dbmsPort, NULL AS dbmsDatabase, NULL AS dbmsEnabled, NULL AS dbmsUseTrustedConnection, NULL AS dbmsUseLoginForm, NULL AS dbmsUser, NULL AS dbmsPwd, NULL AS dbmsOptionsODBC, NULL AS dbmsOptionsOLEDB, NULL AS dbmsDSN, NULL AS dbmsConStrODBC, NULL AS dbmsConStrOLEDB, NULL AS Remarks, NULL AS XdbmsConnectionModes, NULL AS XdbmsServer, NULL AS XdbmsPort, NULL AS XdbmsDatabase, NULL AS XdbmsUseLoginForm, NULL AS XdbmsUseTrustedConnection, NULL AS XdbmsUser, NULL AS XdbmsPwd, NULL AS XdbmsOptionsODBC, NULL AS XdbmsOptionsOLEDB, NULL AS XdbmsDSN, NULL AS XdbmsConStrODBC, NULL AS XdbmsConStrOLEDB, NULL AS XdbmsOledbProvider, NULL AS XdbmsOdbcDriver, NULL AS DriverODBC, NULL AS DriverOLEDB FROM tab1DS where 1=0"
End Sub

Public Sub SaveRecord()
   If Me.Dirty Then Me.Dirty = False
End Sub

Public Sub CreateNewRecord()

   Set Me.Recordset = SourceDb.OpenRecordset(DataSource & " where 1=0", dbOpenDynaset)

   Me.AllowAdditions = True

   Me.txtCID.Locked = False
   Me.txtCID.SetFocus

   Me.txtCID.BackColor = dcwBackgroundTextboxEdit
   Me.txtCID.BorderStyle = 1
   Me.txtCID.BorderColor = vbRed

   Me.cbxDBMS.Visible = True
   Me.cbxDBMS.BackColor = dcwBackgroundTextboxEdit
   'Me.cbxDBMS.BorderStyle = 1
   Me.cbxDBMS.BorderColor = vbRed

   Me.txtDBMS.Visible = False
   SetLayoutMode

End Sub

Private Sub cbTrustedConnection_AfterUpdate()

On Error GoTo HandleErr

   If (Me.cbTrustedConnection.Value = True) And (Me!XdbmsUseTrustedConnection > 0) Then
      Me.cbUseLoginForm.Value = False
      Me.txtDbmsUser.Value = Null
      Me.txtDbmsPwd.Value = Null
   End If

   SetControlMode

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub


Private Sub cbActiveConnection_AfterUpdate()

On Error GoTo HandleErr

   Me.Dirty = False
   If Me.cbActiveConnection.Value = True Then
      SourceDb.Execute "update usys_DbmsConnection set ActiveConnection = 0 where CID <> " & SqlTools.TextToSqlText(Me!CID), dbFailOnError
      ClearConnectionInfos
      CurrentConnectionInfo.ConnectionIdentifier = Nz(Me!CID, vbNullString)
   Else
      ClearApplicationConnectionInfo
      DisposeDbConnection
   End If

   Me.cmdLinkTables.Enabled = (Me.cbActiveConnection.Value = True)
   Me.cmdTestSQL.Enabled = (Me.cbActiveConnection.Value = True)
   RaiseEvent CidStateChanged(Me!CID)

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub SetControlMode()

   Dim bolTrustedConnection As Boolean
   Dim bolUseLoginForm As Boolean

   bolUseLoginForm = Me.cbUseLoginForm.Value
   bolTrustedConnection = Me.cbTrustedConnection.Value And (Me!XdbmsUseTrustedConnection > 0)

   Me.txtDbmsUser.Enabled = Not bolTrustedConnection
   Me.txtDbmsPwd.Enabled = Not (bolTrustedConnection Or bolUseLoginForm)
   If bolTrustedConnection Then
      Me.txtDbmsUser.Value = Null
      Me.txtDbmsPwd.Value = Null
   ElseIf bolUseLoginForm Then
      Me.txtDbmsPwd.Value = Null
   End If

   If Me!XdbmsUseTrustedConnection = 0 And Me.cbTrustedConnection.Value = True Then
      Me.labDbmsUser.Caption = "Loginadmin:"
   Else
      Me.labDbmsUser.Caption = "User:"
   End If

End Sub

Private Sub SetLayoutMode()

   Dim bolTrustedConnection As Boolean
   Dim bolUseLoginForm As Boolean
   Dim bolIsMSJET As Boolean
   Dim lngConnectionMode As Long
   Dim lngAllowedConnectionModes As Long
   Dim strDbmsName As String
   Dim backgroundHeader As Long

   If m_TemplateMode Then
      backgroundHeader = dcwBackgroundLightHeaderTemplates
   Else
      backgroundHeader = dcwBackgroundLightHeader
   End If

   Me!sysFirst.SetFocus

   Me.cbxDBMS.Visible = (Len(Me.cbxDBMS & "") = 0)
   Me.txtDBMS.Visible = Not Me.cbxDBMS.Visible
   Me.txtCID.Locked = (Len(Me.txtCID & "") > 0)
   Me.txtCID.BorderStyle = Abs(Not Me.txtCID.Locked)
   If Me.txtCID.Locked Then
      Me.txtCID.BackColor = backgroundHeader
   Else
      Me.txtCID.BackColor = dcwBackgroundTextboxEdit
   End If

   lngConnectionMode = Nz(Me!dbmsConnectionMode, -1)
   If lngConnectionMode = 0 Then lngConnectionMode = -1 'damit alle Parameterfelder deaktiviert werden
   lngAllowedConnectionModes = Nz(Me!XdbmsConnectionModes, 0)

   Me.cbActiveConnection.Enabled = (lngConnectionMode > 0)
   Me.cmdCheckConnection.Enabled = (lngConnectionMode > 0)

   Me.ogDbmsConnectionMode.Enabled = (lngAllowedConnectionModes > 0)
   Me.txtRemarks.Enabled = (lngAllowedConnectionModes > 0)

   Me.optDbmsConnectionMode1.Enabled = (lngAllowedConnectionModes And 1) = 1
   Me.optDbmsConnectionMode2.Enabled = (lngAllowedConnectionModes And 2) = 2
   Me.optDbmsConnectionMode4.Enabled = (lngAllowedConnectionModes And 4) = 4

   strDbmsName = Nz(Me!DBMS, vbNullString)
   bolIsMSJET = (strDbmsName = "MSJET")
   Me.cmdSelectFile.Visible = bolIsMSJET
   If bolIsMSJET Then
      Me.labDatabase.Caption = "mdb-&Datei:"
      Me.txtDbmsDatabase.Width = Me.cmdSelectFile.Left - Me.txtDbmsDatabase.Left - 100
   Else
      If strDbmsName = "DB2" Then
         Me.labDatabase.Caption = "&DBALIAS:"
      Else
         Me.labDatabase.Caption = "&Database:"
      End If
      Me.txtDbmsDatabase.Width = Me.txtDbmsPort.Left + Me.txtDbmsPort.Width - Me.txtDbmsDatabase.Left
   End If

   Me.txtDbmsServer.Enabled = ((Nz(Me!XdbmsServer, 255) And lngConnectionMode) = lngConnectionMode)
   Me.txtDbmsPort.Enabled = ((Nz(Me!XdbmsPort, 255) And lngConnectionMode) = lngConnectionMode)
   Me.txtDbmsDatabase.Enabled = ((Nz(Me!XdbmsDatabase, 255) And lngConnectionMode) = lngConnectionMode)
   Me.txtDbmsDSN.Enabled = ((Nz(Me!XdbmsDSN, 255) And lngConnectionMode) = lngConnectionMode)
   Me.txtDbmsOptionsODBC.Enabled = ((Nz(Me!XdbmsOptionsODBC, 255) And lngConnectionMode) = lngConnectionMode)
   Me.txtDbmsOptionsOLEDB.Enabled = ((Nz(Me!XdbmsOptionsOLEDB, 255) And lngConnectionMode) = lngConnectionMode)
   Me.cbTrustedConnection.Enabled = ((Nz(Me!XdbmsUseTrustedConnection, 255) And lngConnectionMode) = lngConnectionMode)
   Me.cbUseLoginForm.Enabled = ((Nz(Me!XdbmsUseLoginForm, 255) And lngConnectionMode) = lngConnectionMode)
   Me.txtDbmsOleDbProvider.Enabled = ((Nz(Me!XdbmsOledbProvider, 255) And lngConnectionMode) = lngConnectionMode)
   Me.txtDbmsOdbcDriver.Enabled = ((Nz(Me!XdbmsOdbcDriver, 255) And lngConnectionMode) = lngConnectionMode)

   Me.txtDbmsOleDbProvider.Format = "@;;""" & Me!DriverOLEDB
   Me.txtDbmsOdbcDriver.Format = "@;;""" & Me!DriverODBC

   Me.txtDbmsConStrODBC.Enabled = ((Nz(Me!XdbmsConStrODBC, 255) And lngConnectionMode) = lngConnectionMode)
   Me.txtDbmsConStrOLEDB.Enabled = ((Nz(Me!XdbmsConStrOLEDB, 255) And lngConnectionMode) = lngConnectionMode)

   bolUseLoginForm = Nz(Me.cbUseLoginForm.Value, False)
   bolTrustedConnection = Nz(Me.cbTrustedConnection.Value, False) And (Nz(Me!XdbmsUseTrustedConnection, 0) = 255)

   Me.txtDbmsUser.Enabled = (Not bolTrustedConnection) And ((Nz(Me!XdbmsUser, 255) And lngConnectionMode) = lngConnectionMode)
   Me.txtDbmsPwd.Enabled = (Not bolUseLoginForm) And (Not bolTrustedConnection) And ((Nz(Me!XdbmsPwd, 255) And lngConnectionMode) = lngConnectionMode)

   If Me!XdbmsUseTrustedConnection = 0 And Me.cbTrustedConnection.Value = True Then
      Me.labDbmsUser.Caption = "Loginadmin:"
   Else
      Me.labDbmsUser.Caption = "User:"
   End If

End Sub

Private Sub cbUseLoginForm_AfterUpdate()

On Error GoTo HandleErr

   If Me.cbTrustedConnection = True And Me.cbUseLoginForm = True Then
      Me.cbTrustedConnection = False
   End If
   SetControlMode

   CheckConditionalCompilationArgumentsString Me.cbUseLoginForm

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub cbxDBMS_AfterUpdate()
   RefreshDbmsControl
End Sub

Private Sub RefreshDbmsControl()

   If Me.NewRecord Then
      If Len(Me.cbxDBMS) > 0 Then
         Me.cbxDBMS.BorderColor = dcwControlBorderColor
      Else
         Me.cbxDBMS.BorderColor = vbRed
      End If

      If (Len(Me.cbxDBMS.Value) * Len(Me.txtCID.Value)) > 0 Then
         SaveRecord
         SetLayoutMode
         RaiseEvent CidValueChanged(Nz(Me.txtCID.Value, vbNullString))
      ElseIf Len(Me.txtCID & vbNullString) = 0 Then
         Me.txtCID.SetFocus
      End If
   Else
      SetLayoutMode
   End If

   SetDbmsDriverSource

End Sub

Private Sub cbxDBMS_Enter()
On Error GoTo HandleErr

   If Len(Me.cbxDBMS.RowSource) = 0 Then
      Me.cbxDBMS.RowSource = "SELECT X.DBMS FROM usys_DbmsConnection_X AS X ORDER BY X.seq"
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub cbxDBMS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo HandleErr

   If Button = 2 Then
      OpenContextMenu Me.cbxDBMS, X, Y
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub cmdCheckConnection_Click()

   Dim strCID As String
   Dim oDbConnectionInfo As DbConnectionInfo
   Dim ErrMsg As String

On Error GoTo HandleErr

   If Me.Dirty Then Me.Dirty = False

   strCID = Nz(Me!CID, vbNullString)
   If Len(strCID) = 0 Then Exit Sub

   Set oDbConnectionInfo = New DbConnectionInfo

   If Me.cbUseLoginForm Then
      oDbConnectionInfo.LoginFormName = "LoginForm"
   End If

   oDbConnectionInfo.ConnectionIdentifier = strCID 'Übergabe der Verbindungskennung
   If oDbConnectionInfo.CheckConnection(ErrMsg) = True Then
      MsgBox "Connection was successful"
   Else
      If Len(ErrMsg) > 0 Then ErrMsg = vbNewLine & vbNewLine & ErrMsg
      MsgBox "Connection could not be established." & ErrMsg
   End If
   Set oDbConnectionInfo = Nothing

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub cmdHelp_Click()
On Error GoTo HandleErr

   OpenHelpFile

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub cmdLinkTables_Click()
On Error GoTo HandleErr

   DoCmd.OpenForm DCW_LinkTablesFormName, , , , , , Me.Name

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub cmdSelectFile_Click()

   Dim strFile As String
   Dim strSearchDir As String
   Dim lngPos As Long

On Error GoTo HandleErr

   strSearchDir = Nz(Me.txtDbmsDatabase, vbNullString)
   If Len(strSearchDir) > 0 Then
      lngPos = InStrRev(strSearchDir, "\")
      If lngPos > 0 Then
         strSearchDir = Left$(strSearchDir, lngPos - 1)
      Else
         strSearchDir = vbNullString
      End If
   End If

   WizHook.Key = 51488399
   WizHook.GetFileName Me.Hwnd, CurrentApplication.ApplicationName, L10n.Text("Select database file"), "Select", _
                       strFile, CurrentProject.Path, "*.mdb", 0, 0, 0, True
   If Len(strFile) > 0 Then
      Me.txtDbmsDatabase.Value = strFile
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub Form_AfterUpdate()
On Error GoTo HandleErr

   ClearConnectionInfos

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub ClearConnectionInfos()
   If Me.cbActiveConnection.Value = True Then
      CurrentConnectionInfo.ClearConnectionInfo
      ClearApplicationConnectionInfo
   End If
End Sub

Private Sub Form_Current()

On Error GoTo HandleErr

   SetLayoutMode

   If Me.RecordsetClone.RecordCount = 0 Then Exit Sub

   Me.cmdLinkTables.Enabled = (Nz(Me.cbActiveConnection.Value, False) = True)
   Me.cmdTestSQL.Enabled = (Nz(Me.cbActiveConnection.Value, False) = True)

   If Not m_TemplateMode Then
      CheckModules
   End If

   SetDbmsDriverSource

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo HandleErr

   If KeyCode = vbKeyF1 Then
      OpenHelpFile
      KeyCode = 0
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub Form_Load()
On Error Resume Next
   InitForm
End Sub

Private Sub InitForm()

   Dim backgroundForm As Long
   Dim backgroundHeader As Long

   Dim lngMargin As Long

   lngMargin = Me.labDatabase.Left

   If m_TemplateMode Then
      backgroundForm = dcwBackgroundFormTemplates
      backgroundHeader = dcwBackgroundLightHeaderTemplates
   Else
      backgroundForm = dcwBackgroundForm
      backgroundHeader = dcwBackgroundLightHeader
   End If

   Me.Section(acDetail).BackColor = backgroundForm
   Me.rectHeader.BackColor = backgroundHeader
   Me.txtCID.BackColor = backgroundHeader
   Me.txtDBMS.BackColor = backgroundHeader
   Me.cbxDBMS.BackColor = backgroundHeader

   Me.cbActiveConnection.Visible = Not m_TemplateMode

   Me.cmdLinkTables.Visible = Not m_TemplateMode
   Me.cmdTestSQL.Visible = Not m_TemplateMode
   Me.cmdHelp.Visible = Not m_TemplateMode

   Me.labCopyModulAccUtil.Visible = Not m_TemplateMode
   Me.labCopyModulDbConnectionHandler.Visible = Not m_TemplateMode
   Me.labCopyModulDbConnectionInfo.Visible = Not m_TemplateMode
   Me.labCopyModules.Visible = Not m_TemplateMode

   Me.labCopyCaption.Visible = Not m_TemplateMode
   Me.labCaptionCopyModulAccUtil.Visible = Not m_TemplateMode
   Me.labCaptionCopyModulDbConnectionHandler.Visible = Not m_TemplateMode
   Me.labCaptionCopyModulDbConnectionInfo.Visible = Not m_TemplateMode
   Me.labCopyCaption.Visible = Not m_TemplateMode

   Me.txtRemarks.Height = 0
   If m_TemplateMode Then
      Me.labCaptionRemarks.Top = Me.ogDbmsConnectionMode.Top 'Me.txtDbmsDatabase.Top - Me.labCaptionRemarks.Height
   Else
      Me.labCaptionRemarks.Top = Me.labProviderDriver.Top
   End If
   Me.txtRemarks.Top = Me.labCaptionRemarks.Top + Me.labCaptionRemarks.Height
   Me.txtRemarks.Height = Me.txtDbmsConStrODBC.Top + Me.txtDbmsConStrODBC.Height - Me.txtRemarks.Top

   L10n.TranslateControls Me.Controls
   Set m_L10nDict = L10n

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_L10nDict = Nothing
End Sub

Private Sub m_L10nDict_LanguageChanged()
   L10n.TranslateControls Me.Controls
End Sub

Private Sub labCopyCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   modWinApi_Mouse.MouseCursor IDC_HAND
End Sub

Private Sub labDbmsCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo HandleErr

   If Button = 2 Then
      OpenContextMenu Me.labDbmsCaption, X, Y
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub ogDbmsConnectionMode_AfterUpdate()
On Error GoTo HandleErr

   SetLayoutMode

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub rectHeader_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo HandleErr

   If Button = 2 Then
      OpenContextMenu Me.rectHeader, X, Y
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub txtCID_AfterUpdate()
On Error GoTo HandleErr

   RefreshCidControl

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub RefreshCidControl()

   If Me.NewRecord Then
      If Len(Me.txtCID) > 0 Then
         Me.txtCID.BorderColor = 8421504
      Else
         Me.txtCID.BorderColor = vbRed
      End If

      If (Len(Me.cbxDBMS.Value) * Len(Me.txtCID.Value)) > 0 Then
         SaveRecord
         SetLayoutMode
         RaiseEvent CidValueChanged(Nz(Me.txtCID.Value, vbNullString))
      End If

   Else
      Me.txtCID.Locked = True
      Me.txtCID.BorderStyle = 0
      SaveRecord
      RaiseEvent CidValueChanged(Nz(Me.txtCID.Value, vbNullString))
   End If

End Sub

Private Sub txtCID_DblClick(Cancel As Integer)

On Error GoTo HandleErr

   Me.txtCID.Locked = False
   Me.txtCID.BorderStyle = 1

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub txtCID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo HandleErr

   If Button = 2 Then
      OpenContextMenu Me.txtCID, X, Y
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub txtDBMS_GotFocus()

On Error GoTo HandleErr

   Me.cbxDBMS.Visible = True
   Me.cbxDBMS.SetFocus

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub OpenContextMenu(ByRef ctlMenuControl As Control, ByVal X As Single, ByVal Y As Single)

   Dim mnu As WinApiShortcutMenu
   Dim i As Long
   Dim rst As DAO.Recordset
   Dim cidArray As Variant
   Dim lngSelectedMenuItem As Long
   Dim lngCidArraySize As Long

   If m_TemplateMode Then 'die Vorlagenansicht ist bereits aktiviert
      Exit Sub
   End If

   Set mnu = New WinApiShortcutMenu

   With mnu
      .ControlSection = acDetail
      Set .AccessForm = Me
      Set .MenuControl = ctlMenuControl

      .AddMenuItem -1, "Vorlagen bearbeiten", MenuItemStyle.MF_STRING
      .AddMenuItem -2, "---", MF_SEPARATOR

      Set rst = CodeDb.OpenRecordset("select CID FROM usys_DbmsConnection order by CID", dbOpenSnapshot, dbReadOnly)
      If Not rst.EOF Then
         cidArray = rst.GetRows(rst.RecordCount)
         lngCidArraySize = UBound(cidArray, 2)
      Else
         lngCidArraySize = -1
      End If
      rst.Close
      Set rst = Nothing

      For i = 0 To lngCidArraySize
         .AddMenuItem i + 1, cidArray(0, i), MF_STRING, 1
      Next
      If i > 0 Then
         .AddMenuItem 999, "Parameter von Vorlage übernehmen", MF_STRING Or MF_POPUP, 1
      Else
         .AddMenuItem 999, "Parameter von Vorlage übernehmen", MF_STRING + MF_GRAYED
      End If

   End With

   lngSelectedMenuItem = mnu.OpenMenu(X, Y)
   If lngSelectedMenuItem = -1 Then
      OpenTemplates
   ElseIf lngSelectedMenuItem > 0 Then
      CopyTemplate cidArray(0, lngSelectedMenuItem - 1)
   End If

End Sub

Private Sub CopyTemplate(ByVal sCID As String)

   Dim rst As DAO.Recordset
   Dim fld As DAO.Field
   Dim strSQL As String

   strSQL = "select DBMS, dbmsConnectionMode, dbmsOleDbProvider, dbmsOdbcDriver," & _
            " dbmsServer, dbmsPort, dbmsDatabase, dbmsUseTrustedConnection, dbmsUseLoginForm," & _
            " dbmsUser, dbmsPwd, dbmsOptionsODBC, dbmsOptionsOLEDB, dbmsDSN, dbmsConStrODBC, dbmsConStrOLEDB" & _
            " FROM usys_DbmsConnection where CID = " & SqlTools.TextToSqlText(sCID)

   Set rst = CodeDb.OpenRecordset(strSQL, dbOpenForwardOnly, dbReadOnly)
   With rst
      For Each fld In rst.Fields
         Me.Controls(fld.Name).Value = fld.Value
      Next
      .Close
   End With
   Set rst = Nothing

   If IsNull(Me!CID) Then
      Set rst = SourceDb.OpenRecordset("select CID from usys_DbmsConnection where CID = " & SqlTools.TextToSqlText(sCID), dbOpenForwardOnly, dbReadOnly)
      If rst.EOF Then
         Me!CID = sCID
         RefreshCidControl
      End If
      rst.Close
      Set rst = Nothing
   End If
   RefreshDbmsControl

End Sub

Private Sub OpenTemplates()
   Set m_TemplatesForm = New Form_frmConfig_DBMS
   m_TemplatesForm.TemplateMode = True
   m_TemplatesForm.Visible = True
End Sub

Private Sub txtDBMS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo HandleErr

   If Button = 2 Then
      OpenContextMenu Me.txtDBMS, X, Y
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub SetDbmsDriverSource()
   Me.txtDbmsOdbcDriver.RowSource = "SELECT Driver FROM usys_DbmsConnection_X_Driver WHERE DBMS = " & SqlTools.TextToSqlText(Me!txtDBMS.Value) & " AND ConMode='ODBC'"
   Me.txtDbmsOleDbProvider.RowSource = "SELECT Driver FROM usys_DbmsConnection_X_Driver WHERE DBMS = " & SqlTools.TextToSqlText(Me!txtDBMS.Value) & " AND ConMode='OLEDB'"
End Sub

Private Sub txtDbmsOdbcDriver_Enter()
On Error GoTo HandleErr

   Me.txtDbmsOdbcDriver.ForeColor = 0

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub txtDbmsOdbcDriver_Exit(Cancel As Integer)
On Error GoTo HandleErr

   Me.txtDbmsOdbcDriver.ForeColor = &H808080

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub txtDbmsOleDbProvider_DblClick(Cancel As Integer)
   DoCmd.OpenTable "usys_DbmsConnection_X_Driver", acViewNormal, acEdit
End Sub

Private Sub txtDbmsOleDbProvider_Enter()
On Error GoTo HandleErr

   Me.txtDbmsOleDbProvider.ForeColor = 0

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub txtDbmsOleDbProvider_Exit(Cancel As Integer)
On Error GoTo HandleErr

   Me.txtDbmsOleDbProvider.ForeColor = &H808080

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub cmdTestSQL_Click()
On Error GoTo HandleErr

   DoCmd.OpenForm DCW_TestSqlFormName

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub OpenReplaceModulesContextMenu(ByRef MenuControlRef As Control, ByVal X As Single, ByVal Y As Single)

   Dim mnu As WinApiShortcutMenu
   Dim SelectedMenuItem As Long

   If m_TemplateMode Then
      Exit Sub
   End If

   Set mnu = New WinApiShortcutMenu

   With mnu
      .ControlSection = acDetail
      Set .AccessForm = Me
      Set .MenuControl = MenuControlRef

      .AddMenuItem 3, L10n.Text("Update code modules in add-in from SCC repository (master).")
      .AddMenuItem 4, L10n.Text("Update code modules in add-in from SCC repository (draft).")
      .AddMenuItem 0, "", MF_SEPARATOR
      .AddMenuItem 1, L10n.Text("Renew all code modules (existing modules are overwritten)")

   End With

   SelectedMenuItem = mnu.OpenMenu(X, Y)
   Select Case SelectedMenuItem
      Case 1
         ReplaceAllModules
      Case 3
         RefreshCodeModulesFromSccRepo False
      Case 4
         RefreshCodeModulesFromSccRepo True
      Case Else
         '
   End Select

   Set mnu = Nothing

End Sub

Private Sub ReplaceAllModules()

   Dim bolReplaced As Boolean

On Error Resume Next

   bolReplaced = ReplaceCodeModules("SqlTools", "defDbConnection", "DbConnectionInfo", "AdodbHandler", "DaoHandler", "OdbcHandler", _
                                    "DbConnectionHandler", "DbConnectionSettings", "DbConnectionManager", "modDbConnection")

   If bolReplaced Then

      SetModulCopyControls "CopyModulAccUtil", True
      SetModulCopyControls "CopyModulDbConnectionHandler", True
      SetModulCopyControls "CopyModulDbConnectionInfo", True
      SetModulCopyControls "CopyModules", True

      If m_OpenModules = 0 Then
         MsgBox "Modules and classes updated"
      End If

      m_OpenModules = 0

   End If

End Sub

Private Sub RefreshCodeModulesFromSccRepo(ByVal UseDraft As Boolean)

   With New ACLibGitHubImporter

      .UseDraftBranch = UseDraft

      If .LastCommit(True) = SccRevMin Then
         If MsgBox("All modules are already up to date." & vbNewLine & _
                   "Do you still want to re-import?", vbYesNo + vbDefaultButton2, CurrentApplication.ApplicationName) <> vbYes Then
            Exit Sub
         End If
      End If

      .UpdateCodeModules

   End With

   MsgBox "Codemodules have been updated.", , CurrentApplication.ApplicationName

   SetFormCaption

End Sub

Private Sub labCopyCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo HandleErr

   If Button = 1 Or Button = 2 Then
      OpenReplaceModulesContextMenu Me.labCopyCaption, X, Y
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub CheckModules()

   Dim checkOk As Boolean

On Error Resume Next

   m_OpenModules = 0

   With CodeModulImporter
      checkOk = .CheckCodeModules("SqlTools", "defDbConnection", "DbConnectionInfo")
      m_OpenModules = m_OpenModules + 1 + checkOk
      SetModulCopyControls "CopyModulDbConnectionInfo", checkOk

      checkOk = .CheckCodeModules("AdodbHandler", "DaoHandler", "OdbcHandler", "DbConnectionHandler")
      m_OpenModules = m_OpenModules + 1 + checkOk
      SetModulCopyControls "CopyModulDbConnectionHandler", checkOk

      checkOk = .CheckCodeModules("DbConnectionSettings", "DbConnectionManager", "modDbConnection")
      m_OpenModules = m_OpenModules + 1 + checkOk
      SetModulCopyControls "CopyModulAccUtil", checkOk

      checkOk = (m_OpenModules = 0)
      SetModulCopyControls "CopyModules", checkOk
   End With

End Sub

Private Property Get CodeModulImporter() As AppFileCodeModulTransfer
   If m_CodeModulImporter Is Nothing Then
      Set m_CodeModulImporter = New AppFileCodeModulTransfer
      m_CodeModulImporter.UseVbComponentsImport = APPLICATION_FILTERCODEMODULE_USEVBCOMPONENTSIMPORT
   End If
   Set CodeModulImporter = m_CodeModulImporter
End Property

Private Sub cmdCopyModulAccUtil_Click()
On Error GoTo HandleErr

   CodeModulImporter.TransferCodeModules "DbConnectionSettings", "DbConnectionManager", "modDbConnection"
   m_OpenModules = m_OpenModules - 1
   If Not CheckCodeModules("AdodbHandler", "DaoHandler", "OdbcHandler", "DbConnectionHandler") Then
      cmdCopyModulDbConnectionHandler_Click
   End If
   SetModulCopyControls "CopyModulAccUtil", True

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub cmdCopyModulDbConnectionHandler_Click()
On Error GoTo HandleErr

   CodeModulImporter.TransferCodeModules "AdodbHandler", "DaoHandler", "OdbcHandler", "DbConnectionHandler"
   m_OpenModules = m_OpenModules - 1
   If Not CheckCodeModules("defDbConnection", "DbConnectionInfo") Then
      cmdCopyModulDbConnectionInfo_Click
   End If
   SetModulCopyControls "CopyModulDbConnectionHandler", True

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub cmdCopyModulDbConnectionInfo_Click()
On Error GoTo HandleErr

   CheckConditionalCompilationArgumentsString Me.cbUseLoginForm
   CheckLoginForm
   CodeModulImporter.TransferCodeModules "SqlTools", "defDbConnection", "DbConnectionInfo"
   m_OpenModules = m_OpenModules - 1
   SetModulCopyControls "CopyModulDbConnectionInfo", True

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub cmdCopyModules_Click()
On Error GoTo HandleErr

   CheckConditionalCompilationArgumentsString Me.cbUseLoginForm
   CheckLoginForm

   CodeModulImporter.TransferCodeModules "SqlTools", "defDbConnection", "DbConnectionInfo", "AdodbHandler", "DaoHandler", "OdbcHandler", "DbConnectionHandler", "DbConnectionSettings", "DbConnectionManager", "modDbConnection"
   SetModulCopyControls "CopyModulAccUtil", True
   SetModulCopyControls "CopyModulDbConnectionHandler", True
   SetModulCopyControls "CopyModulDbConnectionInfo", True
   SetModulCopyControls "CopyModules", True
   m_OpenModules = 0

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub SetModulCopyControls(ByVal sModulCode As String, ByVal bExists As Boolean)

   Dim sCaption As String
   Dim lBorderStyle As Long

On Error Resume Next

   If bExists Then Me.sysFirst.SetFocus
   Me.Controls("cmd" & sModulCode).Enabled = Not bExists
   If bExists Then
      sCaption = "b"
      lBorderStyle = 0
   Else
      sCaption = vbNullString
      lBorderStyle = 1
   End If
   With Me.Controls("lab" & sModulCode)
      .Caption = sCaption
      .BorderStyle = lBorderStyle
   End With

   If sModulCode <> "CopyModules" And m_OpenModules = 0 Then
      SetModulCopyControls "CopyModules", True
   End If

End Sub

Private Sub SetFormCaption()
On Error Resume Next
   Me.Parent.SetFormCaption
End Sub
