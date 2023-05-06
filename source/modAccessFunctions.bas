Attribute VB_Name = "modAccessFunctions"
'---------------------------------------------------------------------------------------
' Modul: modAccessFunctionReDefine
'---------------------------------------------------------------------------------------
'/* *
' <summary>
' Überschreibungen von Access/VBA-Funktionen
' </summary>
' <remarks>
' Erleichtert das Programmieren, durch Vorgabe von Prozedurparametern usw.
' </remarks>
'* */
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

'MsgBox-Überschreibung um den Titel einfacher gestalten zu können
Public Function MsgBox(ByVal Prompt As Variant, _
              Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
              Optional ByVal Title As Variant, _
              Optional ByVal HelpFile As Variant, _
              Optional ByVal Context As Variant) As VbMsgBoxResult
   
   Dim strTitle As String

   If Len(Title) > 0 Then
      strTitle = Title
   Else
      strTitle = CurrentApplicationName
   End If
   
   MsgBox = L10nMsgBox(Prompt, Buttons, strTitle, HelpFile, Context)
   
End Function
