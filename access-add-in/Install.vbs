const AddInName = "DBMS Connection Wizard"
const AddInFileName = "DbmsConnectionWizard.accda"
const MsgBoxTitle = "Install/Update ACLib-DBMS-Connection-Wizard"

Dim AddInFileInstalled, CompletedMsg

MsgBox "Before updating the add-in file, the add-in must not be loaded!" & chr(13) & _
       "Close all access instances for safety.", , MsgBoxTitle & ": Information"

Select Case MsgBox("Should the add-in be used as a compiled file (accde)?" + chr(13) & _
                   "(Add-In is compiled and copied to the Add-In directory.)", 3, MsgBoxTitle)
   case 6 ' vbYes
      AddInFileInstalled = CreateMde(GetSourceFileFullName, GetDestFileFullName)
      If AddInFileInstalled Then
	      CompletedMsg = "Add-In was compiled and saved in '" + GetAddInLocation + "'."
      Else
         CompletedMsg = "Error! Compiled file was not created."
      End If
   case 7 ' vbNo
      AddInFileInstalled = FileCopy(GetSourceFileFullName, GetDestFileFullName)
      If AddInFileInstalled Then
	      CompletedMsg = "Add-In was saved in '" + GetAddInLocation + "'."
      Else
	      CompletedMsg = "Error! File was not copied."
      End If
   case else
      
End Select

If AddInFileInstalled = True Then
   RegisterAddIn GetDestFileFullName()
End If

If CompletedMsg > "" Then
   MsgBox CompletedMsg, , MsgBoxTitle
End If


'##################################################
' Functions

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
   FileCopy = True
End Function

Function CreateMde(SourceFilePath, DestFilePath)
   Set AccessApp = CreateObject("Access.Application")
   AccessApp.SysCmd 603, (SourceFilePath), (DestFilePath)
   CreateMde = True
End Function


'##################################################
' Register Menu Add-In

Function RegisterAddIn(AddInFile)

   Dim AddInDb, AccessApp, rst, ItemValue, wsh

   Set AccessApp = CreateObject("Access.Application") 
   Set AddInDb = AccessApp.DBEngine.OpenDatabase(AddInFile)
    
   set wsh = CreateObject("WScript.Shell")
   Set rst = AddInDb.OpenRecordset("select Subkey, ValName, Type, Value from USysRegInfo where ValName > '' Order By ValName", 8) 'dbOpenForwardOnly=8
   Do While Not rst.EOF
        ItemValue = rst.Fields("Value").Value
        If Len(ItemValue) > 0 Then
        If InStr(1, ItemValue, "|ACCDIR") > 0 Then
            ItemValue = AddInDb.Name
        End If
        End If
        RegisterMenuAddInItem AccessApp, wsh, rst.Fields("Subkey").Value, rst.Fields("ValName").Value, rst.Fields("Type").Value, ItemValue
        rst.MoveNext
   Loop
   rst.Close

   AddInDb.Close
	
End Function

Function RegisterMenuAddInItem(AccessApp, wsh, ByVal SubKey, ByVal ItemValName, ByVal RegType, ByVal ItemValue)
    Dim RegName
    RegName = GetRegistryPath(SubKey, AccessApp)
    With wsh
        If Len(ItemValName) > 0 Then
            RegName = RegName & "\" & ItemValName
        End If
        .RegWrite RegName, ItemValue, GetRegTypeString(RegType)
    End With
End Function

Function GetRegTypeString(ByVal RegType)
    Select Case RegType
        Case 1
            GetRegTypeString = "REG_SZ"
        Case 4
            GetRegTypeString = "REG_DWORD"
        Case 0
            GetRegTypeString = vbNullString
        Case Else
            Err.Raise vbObjectError, "GetRegTypeString", "RegType not supported"
    End Select
End Function

Function GetRegistryPath(SubKey, AccessApp)
    GetRegistryPath = Replace(SubKey, "HKEY_CURRENT_ACCESS_PROFILE", HkeyCurrentAccessProfileRegistryPath(AccessApp))
End Function

Function HkeyCurrentAccessProfileRegistryPath(AccessApp)
    HkeyCurrentAccessProfileRegistryPath = "HKCU\SOFTWARE\Microsoft\Office\" & AccessApp.Version & "\Access"
End Function
