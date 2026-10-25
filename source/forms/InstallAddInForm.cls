Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Form: _codelib.addins.shared.InstallAddInForm
'---------------------------------------------------------------------------------------
'
' Start form to install Access add-in
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/InstallAddInForm.frm</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/shared/AddInConfiguration.cls</use>
'  <use>_codelib/addins/shared/AddInInstaller.cls</use>
'  <use>base/modApplication.bas</use>
'  <use>base/modErrorHandler.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_Configuration As AddInConfiguration

Private Property Get Configuration() As AddInConfiguration
   If m_Configuration Is Nothing Then
      Set m_Configuration = New AddInConfiguration
   End If
   Set Configuration = m_Configuration
End Property

Private Sub Form_Load()

   With CurrentApplication
      Me.lblVersionInfo.Caption = .ApplicationFullName & " " & ChrW(&H2022) & " Version " & CurrentApplication.Version
   End With

   LoadDataFromConfiguration True
   CheckInstalled

End Sub

Private Sub CheckInstalled()

   Dim TargetAddInPath As String

   With Configuration
      If .AddInRegFilePath Like "|ACCDIR\*" Then
         With New AddInInstaller
            TargetAddInPath = .GetAddInLocation
         End With
         TargetAddInPath = TargetAddInPath & .FileName
         If CurrentDb.Name = TargetAddInPath Then
            Me.cbCompileAddIn.Visible = False
            Me.cmdInstallAddIn.Visible = False
            Me.Caption = "Installed Add-In"
         End If
      End If
   End With

End Sub

Private Sub LoadDataFromConfiguration(ByVal LoadFromFile As Boolean)

   If LoadFromFile Then
      Configuration.LoadFromCurrentFile
   End If

   With Configuration

      Me.txtFileName.Value = .FileName
      Me.txtAppTitle.Value = .AppTitle

      Me.txtAddInTitle.Value = .Title
      Me.txtAddInAuthor.Value = .Author
      Me.txtAddInCompany.Value = .Company
      Me.txtAddInComment.Value = .Comments

      Me.txtAddInRegPathName.Value = .AddInRegPathName
      Me.txtAddInStartFunction.Value = .AddInStartFunction

   End With

End Sub

Private Sub cmdInstallAddIn_Click()

   Dim Success As Boolean
   Dim InstallMsg As String

   Me.sysFirst.SetFocus
   Me.cmdInstallAddIn.Enabled = False

   With New AddInInstaller
      Success = .InstallAddIn(m_Configuration, Nz(Me.cbCompileAddIn.Value, False), InstallMsg)
   End With

   If Len(InstallMsg) > 0 Then
      VBA.MsgBox InstallMsg, vbInformation, m_Configuration.AddInRegPathName
   End If

   If Success Then
      Application.Quit
   End If

End Sub
