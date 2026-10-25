Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Form: LoginForm
'---------------------------------------------------------------------------------------
'/**
' User login form
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>user/login/LoginForm.frm</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'
Option Compare Database
Option Explicit

#If USELOCALIZATION = 1 Then
Private WithEvents m_L10nDict As L10nDict
Attribute m_L10nDict.VB_VarHelpID = -1
#End If

Public Event Logon(ByVal LoginName As String, ByVal Password As String, ByRef Cancel As Boolean)
Public Event Cancelled()

Public m_CloseFormOK As Boolean

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal Length As Long)

Private Sub cmdCancel_Click()
On Error Resume Next
   RaiseEvent Cancelled
   m_CloseFormOK = True
   DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdLogin_Click()

   Dim strLoginName As String
   Dim strLoginPassword As String
   Dim bolCancel As Boolean

   strLoginName = Me.txtUID & vbNullString
   strLoginPassword = Me.txtPwd & vbNullString

   If Len(strLoginName) * Len(strLoginPassword) = 0 Then
      MsgBox "Please enter login name and password"
      If Len(strLoginName) = 0 Then
         Me.txtUID.SetFocus
      Else
         Me.txtPwd.SetFocus
      End If
      Exit Sub
   End If

   RaiseEvent Logon(strLoginName, strLoginPassword, bolCancel)
   If bolCancel Then
      Me.txtPwd.SetFocus
      Exit Sub
   End If

   m_CloseFormOK = True
   DoCmd.Close acForm, Me.Name

End Sub

Private Sub Form_Load()

   Dim Args() As String
   Dim LoginFormBridgePtr As LongPtr

#If USELOCALIZATION = 1 Then
   L10n.TranslateControls Me.Controls
   Set m_L10nDict = L10n
#End If

   If Len(Me.OpenArgs) > 0 Then 'OpenArgs= "ObjPtr@Caption@StandardUser"
      Args = Split(Me.OpenArgs, "@")

      LoginFormBridgePtr = Val(Args(0))

      Me.Caption = Args(1)
      If UBound(Args) > 1 Then
         Me.txtUID.Value = Args(2)
      Else
         Me.txtUID.Value = vbNullString
      End If

      If Len(Me.txtUID.Value) > 0 Then
         Me.txtPwd.SetFocus
      End If

   End If

   Dim tempObj As Object
   Set tempObj = GetObjectFromObjPtr(LoginFormBridgePtr)
   If Not (tempObj Is Nothing) Then
      Set tempObj.LoginForm = Me 'This makes the events possible despite dialogue mode
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   If Not m_CloseFormOK Then RaiseEvent Cancelled
#If USELOCALIZATION = 1 Then
   Set m_L10nDict = Nothing
#End If
End Sub

Private Function GetObjectFromObjPtr(lPtr As LongPtr) As Object

   Dim obj As Object
   CopyMemory obj, lPtr, 4
   Set GetObjectFromObjPtr = obj
   CopyMemory obj, 0&, 4

End Function

#If USELOCALIZATION = 1 Then
Private Sub m_L10nDict_LanguageChanged()
   L10n.TranslateControls Me.Controls
End Sub
#End If
