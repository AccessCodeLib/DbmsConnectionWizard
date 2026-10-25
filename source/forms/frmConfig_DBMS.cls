Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' VBA Document: Form_frmConfig_DBMS
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Form_frmConfig_DBMS
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Formular: frmConfig_DBMS
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private WithEvents m_DetailForm As Form_frmConfig_DBMS_SF_Data
Attribute m_DetailForm.VB_VarHelpID = -1

Private WithEvents m_L10nDict As L10nDict
Attribute m_L10nDict.VB_VarHelpID = -1

Private m_TemplateMode As Boolean

Public Property Let TemplateMode(ByVal tm As Boolean)
   m_TemplateMode = tm
   Me.lstCID.Value = Null
   InitForm
   Me.lstCID.Value = Me.lstCID.Column(0, 0)
   CidDetailForm.SetData Nz(Me.lstCID.Value, vbNullString)
End Property

Private Property Get SourceDb() As DAO.Database
   If m_TemplateMode Then
      Set SourceDb = CodeDb
   Else
      Set SourceDb = CurrentDb
   End If
End Property

Private Property Get SourceConnection() As ADODB.Connection
   If m_TemplateMode Then
      Set SourceConnection = CodeProject.Connection
   Else
      Set SourceConnection = CurrentProject.Connection
   End If
End Property


'---------------------------------------------------------------------------------------
' Property: CidDetailForm (16.03.2009)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' CidDetailForm
' </summary>
' <returns>Form</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Private Property Get CidDetailForm() As Form
   If m_DetailForm Is Nothing Then
      InitDetailForm
   End If
   Set CidDetailForm = m_DetailForm
End Property

Private Sub InitDetailForm()
   With Me.sfrDetailForm
      If Len(.SourceObject) = 0 Then
         .SourceObject = "frmConfig_DBMS_SF_Data"
      End If
      Set m_DetailForm = .Form
   End With
   m_DetailForm.TemplateMode = m_TemplateMode
End Sub

Private Sub InitConnectionList()

   Dim rst As DAO.Recordset
   Dim strRowSource As String

   Set rst = SourceDb.OpenRecordset("SELECT C.CID FROM usys_DbmsConnection AS C ORDER BY C.ActiveConnection, C.CID", dbOpenForwardOnly, dbReadOnly)
   'Set Me.lstCID.Recordset = rst
   'Schleifendurchlauf statt Recordset-Zuweisung, da die Recordsetzuweisung erst aber Ac02 funktioniert
   With rst
      If Not .EOF Then
         Do
            strRowSource = strRowSource & ";" & .Fields(0).Value
            .MoveNext
         Loop Until .EOF
         strRowSource = Mid$(strRowSource, 2)
      End If
   End With

   Me.lstCID.RowSource = strRowSource

End Sub

Private Function CheckConnectionTable() As Boolean

On Error Resume Next

   If Not TableDefExists("usys_DbmsConnection", SourceDb) Then
      If MsgBox("The required table 'usys_DbmsConnection' does not exist, should it be created?", vbOKCancel + vbQuestion) = vbOK Then
         CheckConnectionTable = CreateConnectionTable(SourceConnection)
      Else
         CheckConnectionTable = False
         Exit Function
      End If
   Else
      CheckConnectionTable = True
   End If

End Function

Private Sub cbxLangCode_AfterUpdate()
   With Me.cbxLangCode
      If IsNull(.Value) Then
         .Value = "EN"
      End If
      L10n.LangCode = .Value
   End With
End Sub

Private Sub cmdNewCID_Click()
On Error GoTo HandleErr

   CreateNewCID

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub Form_Current()
On Error Resume Next
   Me.lstCID.Value = Me.lstCID.Column(0, 0)
   CidDetailForm.SetData Nz(Me.lstCID.Value, vbNullString)
End Sub

Private Sub Form_Deactivate()
On Error Resume Next
    'Zur Sicherheit speichern, damit in anderen Formularen die neuesten Einstellungen gelten
    CidDetailForm.SaveRecord
End Sub

Private Sub InitForm()

   Dim backgroundForm As Long

   If Not CheckConnectionTable Then
      DoCmd.Close acForm, Me.Name
      Exit Sub
   End If

   InitConnectionList
   InitDetailForm

   If m_TemplateMode Then
      backgroundForm = dcwBackgroundFormTemplates
   Else
      backgroundForm = dcwBackgroundForm
   End If

   Me.Section(acDetail).BackColor = backgroundForm
   Me.lstCID.BackColor = backgroundForm

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
On Error GoTo HandleErr

   InitForm

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub Form_Open(Cancel As Integer)
   Me.cbxLangCode.Value = L10n.LangCode
   L10n.TranslateControls Me.Controls
   Set m_L10nDict = L10n
   SetFormCaption
End Sub

Public Sub SetFormCaption()

   Dim FormCaption As String

   If m_TemplateMode Then
      FormCaption = L10n.Text("Edit Template")
      Me.cmdNewCID.Caption = L10n.Text("&New Template")
   Else
      FormCaption = "ACLib DBMS Connection Wizard  " & ChrW(&H25AA) & "  Version " & CurrentApplication.Version & " (" & L10n.Text("Classes & Modules") & ": SCC-Rev. " & SccRev & ")"
      Me.cmdNewCID.Caption = L10n.Text("&New Connection")
   End If

   Me.Caption = FormCaption

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   Set m_DetailForm = Nothing
   DisposeWizard
End Sub

Private Sub lstCID_AfterUpdate()
On Error GoTo HandleErr

   CidDetailForm.SetData Nz(Me.lstCID.Value, vbNullString)

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub lstCID_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo HandleErr

   If KeyCode = vbKeyDelete Then
      deleteSelectedItem
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox

End Sub

Private Sub deleteSelectedItem()

   Dim strCID As String

   If Me.lstCID.Column(1) = True Then
      MsgBox "Eine aktivierte Verbindung kann nicht gelöscht werden!", vbCritical + vbOKOnly
      Exit Sub
   End If
   strCID = Me.lstCID.Value
   If MsgBox("Sollen die Daten von '" & strCID & "' gelöscht werden?", vbYesNo + vbDefaultButton2) = vbYes Then
      CidDetailForm.SetNoData
      Me.lstCID.Value = vbNullString
      SourceDb.Execute "delete from usys_DbmsConnection where CID = " & SqlTools.TextToSqlText(strCID), dbFailOnError
      InitConnectionList
   End If

End Sub

Private Sub CopySelectedItem()

   Dim strCID As String
   Dim strCIDnew As String
   Dim lngCounter As Long
   Dim db As DAO.Database

   strCID = Me.lstCID.Value

   Set db = SourceDb
   lngCounter = 1
   strCIDnew = strCID & "_1"
   Do While db.OpenRecordset("select count(*) from usys_DbmsConnection where CID='" & strCIDnew & "'")(0) > 0
      lngCounter = lngCounter + 1
      strCIDnew = strCID & "_" & lngCounter
   Loop

   CidDetailForm.SetNoData
   Me.lstCID.Value = vbNullString

   db.Execute "insert into usys_DbmsConnection (" & DCW_usys_DbmsConnection_Fields & ")" & _
              " select " & Replace(Replace(DCW_usys_DbmsConnection_Fields, "CID", SqlTools.TextToSqlText(strCIDnew)), "ActiveConnection", "false") & " from usys_DbmsConnection" & _
              " where CID = " & SqlTools.TextToSqlText(strCID), dbFailOnError
   Set db = Nothing

   InitConnectionList
   Me.lstCID.Value = strCIDnew
   CidDetailForm.SetData strCIDnew

End Sub

Private Sub CreateNewCID()
   Me.lstCID.Value = Null
   CidDetailForm.CreateNewRecord
End Sub

Private Sub OpenListboxContextMenu(ByVal X As Single, ByVal Y As Single)

   Dim mnu As WinApiShortcutMenu

   Set mnu = New WinApiShortcutMenu
   With mnu
      .ControlSection = acDetail
      Set .AccessForm = Me
      Set .MenuControl = Me.lstCID
      If Len(Me.lstCID.Value) > 0 Then
         .AddMenuItem 1, "markierten Eintrag löschen", MF_STRING
         .AddMenuItem 2, "markierten Eintrag kopieren", MF_STRING
         .AddMenuItem 0, "---", MF_SEPARATOR
      End If
      .AddMenuItem 3, "Neuen Eintrag erstellen", MF_STRING
   End With

   Select Case mnu.OpenMenu(X, Y)
      Case 1
         deleteSelectedItem
      Case 2
         CopySelectedItem
      Case 3
         CreateNewCID
      Case Else
         '
   End Select

End Sub

Private Sub lstCID_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo HandleErr

   If Button = 2 Then
      OpenListboxContextMenu X, Y
   End If

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub m_DetailForm_CidStateChanged(ByVal currentCID As String)
On Error GoTo HandleErr

   InitConnectionList

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub m_DetailForm_CidValueChanged(ByVal newCID As String)
On Error GoTo HandleErr

   With Me.lstCID
      .Value = Null
      InitConnectionList
      .Value = newCID
   End With

ExitHere:
   Exit Sub

HandleErr:
   HandleError Err.Number, Err.Source, Err.Description, aclibErrMsgBox
End Sub

Private Sub m_L10nDict_LanguageChanged()
   If Nz(Me.cbxLangCode.Value, vbNullString) <> L10n.LangCode Then
      Me.cbxLangCode.Value = L10n.LangCode
   End If
   SetFormCaption
   L10n.TranslateControls Me.Controls
End Sub

Private Sub labCheckVersion_Click()

   Select Case L10n.ControlTag(Me.labCheckVersion.Tag)
      Case "unchecked"
         SetNewVersionInfo
      Case "NewVersionExists"
         OpenDownloadSource
      Case Else
         '
   End Select

End Sub

Private Sub SetNewVersionInfo()

   Dim AhProcParams(1) As Variant
   Dim AhResumeMode As ApplicationHandlerResumeModes
   Dim AhResumeMessage As Variant

   Call CurrentApplication.CallExtensionProcedure("AppVersion", "NewVersionExists", AhResumeMode, AhResumeMessage, AhProcParams)

   If AhResumeMode <> AppResumeMode_Completed Then
      If AhResumeMode = AppResumeMode_Error Then
      If Len(AhResumeMessage) > 0 Then
         MsgBox AhResumeMessage, vbExclamation, "Error during version check"
      End If
      End If
      Exit Sub
   End If

   With Me.labCheckVersion
   If AhResumeMessage = False Then
      .Caption = L10n.Text("Latest version is installed")
      .Tag = "VersionUptodate" & L10n.LangSeparatorChar & "LANG:Latest version is installed"
      .ForeColor = 0
      .FontUnderline = False
   ElseIf AhResumeMessage = True Then
      .Caption = FormatText("New Version ({0}) available", AhProcParams(0))
      .ControlTipText = AhProcParams(1)
      .Tag = "NewVersionExists"
   Else

   End If
   End With

End Sub

Private Sub OpenDownloadSource()
   Application.FollowHyperlink APPLICATION_DOWNLOADSOURCE, , True, False
End Sub

Private Function GetCheckedControlName(ByVal ControlNameToCheck As String) As String
   GetCheckedControlName = Replace(ControlNameToCheck, " ", "_")
End Function
