Attribute VB_Name = "modAccessFunctions"
'---------------------------------------------------------------------------------------
' Package: modAccessFunctionReDefine
'---------------------------------------------------------------------------------------
'
' Replacements of Access/VBA functions
'
' Remarks:
'     Simplifies programming by specifying procedure parameters, etc.
'
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

Public Function MsgBox(ByVal Prompt As Variant, _
              Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
              Optional ByVal Title As Variant, _
              Optional ByVal HelpFile As Variant, _
              Optional ByVal Context As Variant) As VbMsgBoxResult
   
   Dim strTitle As String

   If IsMissing(Title) Then
      Title = CurrentApplication.ApplicationName
   End If
   
   MsgBox = L10n.MsgBox(Prompt, Buttons, Title, HelpFile, Context)
   
End Function

Public Function InputBox(ByVal Prompt As Variant, _
                Optional ByVal Title As Variant, _
                Optional ByVal Default As Variant, _
                Optional ByVal XPos As Variant, Optional ByVal YPos As Variant, _
                Optional ByVal HelpFile As Variant, _
                Optional ByVal Context As Variant) As String
   
   Dim strTitle As String

   If IsMissing(Title) Then
      Title = CurrentApplication.ApplicationName
   End If
   
   InputBox = L10n.InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context)
   
End Function
