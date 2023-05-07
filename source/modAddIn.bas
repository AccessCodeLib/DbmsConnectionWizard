Attribute VB_Name = "modAddIn"
Option Compare Text
Option Explicit
Option Private Module

'Tabelle usys_DbmsConnection erstellen
Public Function CreateConnectionTable(ByRef cnn As ADODB.Connection) As Boolean

   Dim strSQL As String

   strSQL = "CREATE TABLE usys_DbmsConnection " & _
            "([CID] varchar(20) WITH COMPRESSION NOT NULL," & _
            " [ActiveConnection] bit NOT NULL DEFAULT 0," & _
            " [DBMS] varchar(20) WITH COMPRESSION NOT NULL," & _
            " [dbmsConnectionMode] byte DEFAULT 1," & _
            " [dbmsOleDbProvider] varchar(255) WITH COMPRESSION," & _
            " [dbmsOdbcDriver] varchar(255) WITH COMPRESSION," & _
            " [dbmsServer] varchar(255) WITH COMPRESSION," & _
            " [dbmsPort] varchar(20) WITH COMPRESSION," & _
            " [dbmsDatabase] varchar(255) WITH COMPRESSION," & _
            " [dbmsUseTrustedConnection] bit DEFAULT 0," & _
            " [dbmsUseLoginForm] bit DEFAULT 0," & _
            " [dbmsUser] varchar(255) WITH COMPRESSION," & _
            " [dbmsPwd] varchar(255) WITH COMPRESSION," & _
            " [dbmsOptionsODBC] varchar(255) WITH COMPRESSION," & _
            " [dbmsOptionsOLEDB] varchar(255) WITH COMPRESSION," & _
            " [dbmsDSN] varchar(255) WITH COMPRESSION," & _
            " [dbmsConStrODBC] varchar(255) WITH COMPRESSION," & _
            " [dbmsConStrOLEDB] varchar(255) WITH COMPRESSION," & _
            " [Remarks] text WITH COMPRESSION," & _
            " CONSTRAINT PK_usys_DbmsConnection PRIMARY KEY ([CID]))"
            
   cnn.Execute strSQL
   
   Application.RefreshDatabaseWindow
   
   CreateConnectionTable = True

End Function

Public Sub ClearApplicationConnectionInfo()

On Error Resume Next
#If DemoVersion = 1 Then
#Else

   If CheckCodeModule("modDbConnection", False) Then
      VBA.CallByName Application.Run(CurrentVbProject.Name & ".CurrentConnectionInfo"), "ClearConnectionInfo", VbMethod
   End If

#End If

End Sub

'Prüfen ob Login-Formular vorhanden ist
Public Sub CheckLoginForm()
   
   Dim rst As DAO.Recordset
   Dim bolMissingForm As Boolean

   Set rst = CurrentDb.OpenRecordset("Select O.Name from MSysObjects O where O.Name = '" & DCW_LoginFormName & "' and O.Type=-32768", dbOpenForwardOnly, dbReadOnly)
   bolMissingForm = rst.EOF
   rst.Close
   Set rst = Nothing
      
   If bolMissingForm Then
      'DoCmd.CopyObject CurrentDb.Name, DCW_LoginFormName, acForm, DCW_LoginFormName
      TransferCodeModul CurrentProject, acForm, DCW_LoginFormName
   End If

End Sub

'Module u. Klassen übertragen
Public Sub TransferCodeModules(ParamArray sModulName() As Variant)
   
   Dim i As Long
   Dim ArrSize As Long

   ArrSize = UBound(sModulName)
   For i = 0 To ArrSize
      CheckCodeModule sModulName(i), True
   Next

End Sub

'Module u. Klassen erneuern
Public Function ReplaceCodeModules(ParamArray sModulName() As Variant) As Boolean
   
   Dim i As Long
   Dim ArrSize As Long
   Dim vbp As Object 'VBProject
 
   'VBProject der Anwendung:
   Set vbp = CurrentVbProject

   'Module erneuern:
   If Not (vbp Is Nothing) Then
      ArrSize = UBound(sModulName)
      For i = 0 To ArrSize
      
         If CheckCodeModule(sModulName(i)) Then
            'Modul löschen
            vbp.VBComponents.Remove vbp.VBComponents(sModulName(i))
         End If
         
         'Module kopieren:
         CheckCodeModule sModulName(i), True
         
      Next
      ReplaceCodeModules = True
   End If
   
   Set vbp = Nothing

End Function


'Module u. Klassen auf Existenz prüfen. Es erfolgt keine inhaltliche Prüfung!
Public Function CheckCodeModules(ParamArray sModulName() As Variant) As Boolean
   
   Dim i As Long
   Dim bolModulesExists As Boolean
   Dim ArrSize As Long

   ArrSize = UBound(sModulName)
   bolModulesExists = True
   For i = 0 To ArrSize
      bolModulesExists = bolModulesExists And CheckCodeModule(sModulName(i), False)
   Next
   
   CheckCodeModules = bolModulesExists

End Function

'Modul oder Klasse auf Existenz prüfen. Es erfolgt keine inhaltliche Prüfung!
Public Function CheckCodeModule(ByVal sModulName As String, _
                       Optional ByVal TransferMissingModule As Boolean = False) As Boolean
   
   Dim rst As DAO.Recordset
   Dim bolMissingModule As Boolean

   Set rst = CurrentDb.OpenRecordset("Select O.Name from MSysObjects O where O.Name = '" & sModulName & "' and O.Type=-32761", dbOpenForwardOnly, dbReadOnly)
   bolMissingModule = rst.EOF
   rst.Close
   Set rst = Nothing
   
   If bolMissingModule And TransferMissingModule Then
'      If Left(sModulName, 3) = "def" Or Left(sModulName, 3) = "mod" Then
'         DoCmd.CopyObject CurrentDb.Name, "DCW_" & sModulName, acModule, sModulName
'      Else
         'DoCmd.CopyObject CurrentDb.Name, sModulName, acModule, sModulName
         TransferCodeModul CurrentProject, acModule, sModulName
'      End If
      bolMissingModule = False
   End If
   
   CheckCodeModule = Not bolMissingModule

End Function


Private Sub TransferCodeModul(ByRef TargetProject As Access.CurrentProject, ByVal ObjType As AcObjectType, ByVal sModulName As String)

   Dim strFileName As String
   
   strFileName = GetTempFileName
   CreateModuleFile sModulName, strFileName
   TargetProject.Application.LoadFromText ObjType, sModulName, strFileName
   Kill strFileName
   
End Sub
