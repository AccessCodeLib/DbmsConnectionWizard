const AddInName = "DBMS Connection Wizard"
const AddInFileName = "DbmsConnectionWizard.accda"
const MsgBoxTitle = "Install/Update ACLib-Import-Wizard"

MsgBox "Before updating the add-in file, the add-in must not be loaded!" & chr(13) & _
       "Close all access instances for safety.", , MsgBoxTitle & ": Information"

Select Case MsgBox("Should the add-in be used as ACCDE?" + chr(13) & _
                   "(The compiled Add-In is copied into the Add-In directory.)", 3, MsgBoxTitle)
   case 6 ' vbYes
      CreateMde GetSourceFileFullName, GetDestFileFullName
	MsgBox "Compiled add-in created"
   case 7 ' vbNo
      FileCopy GetSourceFileFullName, GetDestFileFullName
	MsgBox "Add-In file was copied"
   case else
      
End Select


'##################################################
' Hilfsfunktionen:

Function GetSourceFileFullName()
   GetSourceFileFullName = GetScriptLocation & AddInFileName 
End Function

Function GetDestFileFullName()
   GetDestFileFullName = GetAddInLocation & AddInFileName 
End Function

Function GetScriptLocation()

   With WScript
      GetScriptLocation = Replace(.ScriptFullName & ":", .ScriptName & ":", "") 
   End With

End Function

Function GetAddInLocation()

   GetAddInLocation = GetAppDataLocation & "Microsoft\AddIns\"

End Function

Function GetAppDataLocation()

   Set wsShell = CreateObject("WScript.Shell")
   GetAppDataLocation = wsShell.ExpandEnvironmentStrings("%APPDATA%") & "\"

End Function

Function FileCopy(SourceFilePath, DestFilePath)

   set fso = CreateObject("Scripting.FileSystemObject") 
   fso.CopyFile SourceFilePath, DestFilePath

End Function

Function CreateMde(SourceFilePath, DestFilePath)

   Set AccessApp = CreateObject("Access.Application")
   AccessApp.SysCmd 603, (SourceFilePath), (DestFilePath)
   
End Function
