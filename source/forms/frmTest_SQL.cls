Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const m_Caption As String = "SQL testen"

Private WithEvents m_DbConnectionInfo As DbConnectionInfo
Attribute m_DbConnectionInfo.VB_VarHelpID = -1
Private m_CalledByWizard As Boolean

Private Property Get CurrentDbConnectionInfo() As DbConnectionInfo

   If m_DbConnectionInfo Is Nothing Then
      Set m_DbConnectionInfo = CurrentConnectionInfo
      If Len(m_DbConnectionInfo.CID) = 0 Then
         m_DbConnectionInfo.CheckConnection
      End If
   End If
   Set CurrentDbConnectionInfo = m_DbConnectionInfo

End Property

Private Function getCID() As String
   getCID = CurrentDbConnectionInfo.CID
End Function


Private Sub setCmdStatus(ByVal bCmdStart As Boolean)

   Dim bolEnableCmdBtn As Boolean

   Me!sysFirst.SetFocus

   DoCmd.Hourglass bCmdStart
   bolEnableCmdBtn = Not bCmdStart

   Me!cmdExecuteSQL.Enabled = bolEnableCmdBtn

End Sub

Private Function checkSelection() As Boolean

   If Len(CurrentDbConnectionInfo.DbmsName) = 0 Then
      checkSelection = False
      MsgBox "Please select the DBMS first!"
      Exit Function
   End If

   If IsNull(Me.txtSQL.Value) Then
      checkSelection = False
      MsgBox "Please specify SQL statement."
      Exit Function
   End If

   checkSelection = True

End Function

Private Sub ExecuteSQL()

   Select Case Me.optRS
      Case 1 'ADO
         ExecuteSQL_ADO
      Case 2 'DAO/DB
         ExecuteSQL_DAO
      Case 3 'DAO/PT
         ExecuteSQL_DAOPT
      Case Else
         '
   End Select

End Sub


Private Sub ExecuteSQL_ADO()

   Dim frmDS As Form_frmTest_SQL_Datasheet

   Dim cnn As ADODB.Connection
   Dim rst As ADODB.Recordset
   Dim strErr As String
   Dim errCnn As ADODB.Error
   Dim strSQL As String

   Dim T As QpcStopwatch

   Set T = New QpcStopwatch

   If checkSelection = False Then
      Exit Sub
   End If

   strSQL = Me.txtSQL
   Me.labMsg.Caption = L10n.Text("Statement is executed")
   Me.labTime.Caption = vbNullString
   Me.Repaint

   Set cnn = DbCon.CurrentConnection

On Error Resume Next
   T.StartT
   Set rst = DbCon.ADODB.OpenRecordset(strSQL, adOpenKeyset, adLockOptimistic, adUseClient)
   T.StopT
   If Err.Number <> 0 Then
      strErr = Err.Description
   End If

   If cnn.Errors.Count > 0 Then
      For Each errCnn In cnn.Errors
         strErr = strErr & errCnn.Description & vbNewLine
      Next
      cnn.Errors.Clear
   End If

On Error GoTo 0
   If Len(strErr) > 0 Then
      If (rst.State = adStateOpen And adStateOpen) = adStateOpen Then rst.Close
      Me.labMsg.Caption = strErr
      MsgBox strErr
   Else
      If (rst.State = adStateOpen And adStateOpen) = adStateOpen Then 'rst ist offen => in Form anzeigen
         DoCmd.OpenForm "frmTest_SQL_Datasheet", acFormDS
         Set frmDS = Forms("frmTest_SQL_Datasheet")
         frmDS.SetFormData rst
      End If
      Me.labMsg.Caption = L10n.Text("SQL statement was executed.")
      Me.labTime.Caption = Format$(T.TimeMSec, "0.00 \m\s")
      setCmdStatus False
   End If

   Set T = Nothing
   Set rst = Nothing
   Set cnn = Nothing

End Sub

Private Sub ExecuteSQL_DAO()

   Dim frmDS As Form_frmTest_SQL_Datasheet

   Dim rst As DAO.Recordset
   Dim strErr As String
   Dim bolOpenRS As Boolean

   Dim T As QpcStopwatch

   Set T = New QpcStopwatch

   If checkSelection = False Then
      Exit Sub
   End If

   'setCmdStatus True
   Me.labMsg.Caption = L10n.Text("Statement is executed")
   Me.labTime.Caption = vbNullString
   Me.Repaint

On Error Resume Next
   'einfach einmal versuchen, ob SQL-String per openrecordset geöffnet werden kann
   bolOpenRS = True
   T.StartT
   Set rst = DbCon.ODBC.OpenRecordset(Me.txtSQL.Value, dbOpenDynaset, dbSeeChanges)
   T.StopT
   If Err.Number = 3219 Or Err.Number = 3324 Then 'keine DS zurückgegeben = Aktionsabfrage?
      bolOpenRS = False
      Err.Clear
      T.StartT
      DbCon.ODBC.Execute Me.txtSQL.Value
      T.StopT
   End If

   If Err.Number <> 0 Then
      strErr = "Fehler " & Err.Number & ": " & Err.Description
      Err.Clear
   End If

   If Len(strErr) > 0 Then
      Me.labMsg.Caption = strErr
      MsgBox strErr
   Else
On Error GoTo 0
      If bolOpenRS Then  'rst öffnen => in Form anzeigen
         DoCmd.OpenForm "frmTest_SQL_Datasheet", acFormDS
         Set frmDS = Forms("frmTest_SQL_Datasheet")
         frmDS.SetFormData rst
      End If
      Me.labMsg.Caption = L10n.Text("SQL statement was executed.")
      Me.labTime.Caption = Format$(T.TimeMSec, "0.00 \m\s")
      setCmdStatus False
   End If

End Sub

Private Sub ExecuteSQL_DAOPT()

   Dim frmDS As Form_frmTest_SQL_Datasheet

   Dim rst As DAO.Recordset
   Dim strErr As String
   Dim bolOpenRS As Boolean

   Dim T As QpcStopwatch

   Set T = New QpcStopwatch

   If checkSelection = False Then
      Exit Sub
   End If

   'setCmdStatus True
   Me.labMsg.Caption = L10n.Text("Statement is executed")
   Me.labTime.Caption = vbNullString
   Me.Repaint

On Error Resume Next
   bolOpenRS = True

   T.StartT
   Set rst = DbCon.ODBC.OpenRecordsetPT(Me.txtSQL.Value, dbOpenDynaset)
   T.StopT
   If Err.Number = 3325 Then 'keine DS zurückgegeben = Aktionsabfrage?
      bolOpenRS = False
      Err.Clear
      ' .. Abfrage müsste trotzem ausgeführt worden sein => keine weiteres Execute erforderlich
      'T.StartT
      'DbConnection.ODBC.ExecutePT Me.txtSQL.Value, dbSQLPassThrough
      'T.StopT
   End If

   If Err.Number <> 0 Then 'es ist ein Fehler aufgetreten
      strErr = "Fehler " & Err.Number & ": " & Err.Description
      Err.Clear
   End If

   If Len(strErr) > 0 Then
      Me.labMsg.Caption = strErr
      MsgBox strErr
   Else
On Error GoTo 0
      If bolOpenRS Then  'rst öffnen => in Form anzeigen
         DoCmd.OpenForm "frmTest_SQL_Datasheet", acFormDS
         Set frmDS = Forms("frmTest_SQL_Datasheet")
         frmDS.SetFormData rst
      End If
      Me.labMsg.Caption = L10n.Text("SQL statement was executed.")
      Me.labTime.Caption = Format$(T.TimeMSec, "0.00 \m\s")
      setCmdStatus False
   End If

End Sub

Private Sub cmdExecuteSQL_Click()
On Error Resume Next
   ExecuteSQL
End Sub

Private Sub cmdSavePT_Click()
On Error Resume Next
   savePassThroughQueryDef
End Sub

Private Sub savePassThroughQueryDef()

   Dim QueryDefName As String
   Dim queryExists As Boolean
   Dim qdf As DAO.QueryDef
   Dim db As DAO.Database
   Dim Cancel As Boolean

   QueryDefName = vbNullString
   Do Until (getPtQueryDefName(QueryDefName, queryExists, Cancel) = True)
      If Cancel Then
         Exit Sub
      End If
   Loop

   Set db = CurrentDb

   If queryExists Then
      Set qdf = db.QueryDefs(QueryDefName)
   Else
      Set qdf = db.CreateQueryDef(QueryDefName)
   End If

   qdf.SQL = Me.txtSQL.Value
   qdf.Connect = DbCon.ODBC.ConnectionString
   qdf.Close

   Set qdf = Nothing
   Set db = Nothing

   RefreshDatabaseWindow

End Sub

Private Function getPtQueryDefName(ByRef QueryDefName As String, ByRef bQueryDefExists As Boolean, ByRef Cancelled As Boolean) As Boolean

   QueryDefName = InputBox("Under which name should the query be saved?", "Enter query name", QueryDefName)

   If Len(QueryDefName) = 0 Then
      Cancelled = True
      getPtQueryDefName = False
      Exit Function
   End If

   If TableDefExists(QueryDefName) Then
      If VBA.MsgBox(FormatText(L10n.Text("There is already a table named '{0}'"), QueryDefName) & vbNewLine & _
                L10n.Text("Please choose another name."), vbExclamation + vbDefaultButton2 + vbOKCancel) = vbOK Then
         getPtQueryDefName = False
         Exit Function
      Else
         Cancelled = True
         getPtQueryDefName = False
         Exit Function
      End If
   End If

   bQueryDefExists = QueryDefExists(QueryDefName)
   If bQueryDefExists Then
      Select Case VBA.MsgBox(FormatText(L10n.Text("Es existiert bereits eine Abfrage mit dem Namen '{0}'"), QueryDefName) & vbNewLine & _
                         L10n.Text("Should this be changed?"), vbExclamation + vbDefaultButton2 + vbYesNoCancel)

         Case vbYes
            '
         Case vbCancel
            Cancelled = True
            getPtQueryDefName = False
            Exit Function
         Case Else
            getPtQueryDefName = False
            Exit Function

      End Select

   End If

   getPtQueryDefName = True

End Function

Private Sub Form_Close()
On Error Resume Next
   DoCmd.Hourglass False 'falls es einmal ein Problem gab
End Sub

Private Sub Form_Load()
On Error Resume Next

   setDbmsInfo

   Me.Section(acDetail).BackColor = dcwBackgroundForm
   Me.labDbmsName.BackColor = dcwInfoLabelBackColor
   Me.labTimeL.BackColor = dcwInfoLabelBackColor
   Me.labTime.BackColor = dcwInfoLabelBackColor
   Me.labMsg.BackColor = dcwInfoLabelBackColor

   Me.Caption = m_Caption

   If Len(Me.OpenArgs) > 0 Then
      m_CalledByWizard = True
   End If

End Sub

Private Sub Form_Open(Cancel As Integer)
On Error GoTo HandleErr

   If Len(Me.OpenArgs) > 0 Then
      Dim lngRecordsetMode As Long
      lngRecordsetMode = Val(Me.OpenArgs)
      If lngRecordsetMode > 0 And lngRecordsetMode < 4 Then
         Me.optRS.Value = lngRecordsetMode
      End If
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub Form_Resize()

   Dim lngH As Long

On Error Resume Next

   lngH = Me.InsideHeight
   Me.Section(0).Height = lngH
   Me.txtSQL.Height = lngH - Me.txtSQL.Top - Me.labMsg.Height - 200
   Me.labMsg.Top = lngH - Me.labMsg.Height - 100
   Me.labTime.Top = lngH - Me.labTime.Height - 100
   Me.labTimeL.Top = Me.labTime.Top - Me.labTimeL.Height

   Me.txtSQL.Width = Me.InsideWidth - Me.txtSQL.Left - 100
   Me.labMsg.Width = Me.InsideWidth - Me.labMsg.Left - 100

End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   Set m_DbConnectionInfo = Nothing
   If Not m_CalledByWizard Then
      DisposeWizard
   End If
End Sub

Private Sub m_DbConnectionInfo_DbmsConnectionChanged(ByVal newCID As String)
On Error GoTo HandleErr

   setDbmsInfo

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub m_DbConnectionInfo_DbmsInfoMessage(ByVal Source As String, ByVal Msg As String, ByVal MsgStyle As VbMsgBoxStyle)
On Error GoTo HandleErr

On Error Resume Next
   Me.labMsg.Caption = Msg

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub setDbmsInfo()

   Dim strCID As String
   Dim bolEnabled As Boolean

   strCID = getCID
   bolEnabled = (Len(strCID) > 0)
   If Not bolEnabled Then
      strCID = L10n.Text("Select connection")
   End If
   Me.labDbmsName.Caption = strCID
   Me.cmdExecuteSQL.Enabled = bolEnabled
   Me.optDAPPT.Enabled = (CurrentConnectionInfo.DbmsName <> "MSJET")

End Sub
