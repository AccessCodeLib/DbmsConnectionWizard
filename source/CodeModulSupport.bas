Attribute VB_Name = "CodeModulSupport"
Option Compare Database
Option Explicit

Private Const ConditionalCompilationArgumentsOptionName As String = "Conditional Compilation Arguments"
Private Const CONFIG_USELOGINFORM_ArgName As String = "CONFIG_USELOGINFORM"

Public Sub CheckConditionalCompilationArgumentsString(ByVal UseLoginForm As Boolean)

   Dim ConditionalCompilationArgumentsString As String
   Dim ConditionalCompilationArguments() As String
   Dim TestString As String
   Dim ArgExists As Boolean
   Dim i As Long
   
   ConditionalCompilationArgumentsString = Trim(Nz(Application.GetOption(ConditionalCompilationArgumentsOptionName), vbNullString))
   
   If Len(ConditionalCompilationArgumentsString) = 0 And UseLoginForm Then
      Application.SetOption ConditionalCompilationArgumentsOptionName, CONFIG_USELOGINFORM_ArgName & " = " & Abs(UseLoginForm)
      Exit Sub
   End If
   
   ConditionalCompilationArguments = GetConditionalCompilationArgumentsArray(ConditionalCompilationArgumentsString)
   For i = LBound(ConditionalCompilationArguments) To UBound(ConditionalCompilationArguments)
      If Replace(ConditionalCompilationArguments(i), " ", vbNullString) Like CONFIG_USELOGINFORM_ArgName & "=*" Then
         ConditionalCompilationArguments(i) = CONFIG_USELOGINFORM_ArgName & " = " & Abs(UseLoginForm)
         ArgExists = True
         Exit For
      End If
   Next
   
   If ArgExists Then
      ConditionalCompilationArgumentsString = Join(ConditionalCompilationArguments, ":")
   Else
      ConditionalCompilationArgumentsString = ConditionalCompilationArgumentsString & " : " & CONFIG_USELOGINFORM_ArgName & " = " & Abs(UseLoginForm)
   End If
   
   ConditionalCompilationArgumentsString = Trim(ConditionalCompilationArgumentsString)
   If Left(ConditionalCompilationArgumentsString, 1) = ":" Then
      ConditionalCompilationArgumentsString = Trim(Mid(ConditionalCompilationArgumentsString, 2))
   End If
   If Right(ConditionalCompilationArgumentsString, 1) = ":" Then
      ConditionalCompilationArgumentsString = Trim(Left(ConditionalCompilationArgumentsString, Len(ConditionalCompilationArgumentsString) - 1))
   End If

   Application.SetOption ConditionalCompilationArgumentsOptionName, ConditionalCompilationArgumentsString
   
End Sub

Private Function GetConditionalCompilationArgumentsArray(ByVal FullString As String) As String()
   GetConditionalCompilationArgumentsArray = Split(FullString, ":")
End Function
