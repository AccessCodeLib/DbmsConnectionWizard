Attribute VB_Name = "modAddIn"
Option Compare Text
Option Explicit
Option Private Module

Private m_CodeModulImporter As AppFileCodeModulTransfer

Private Property Get CodeModulImporter() As AppFileCodeModulTransfer
   If m_CodeModulImporter Is Nothing Then
      Set m_CodeModulImporter = New AppFileCodeModulTransfer
      m_CodeModulImporter.UseVbComponentsImport = APPLICATION_FILTERCODEMODULE_USEVBCOMPONENTSIMPORT
   End If
   Set CodeModulImporter = m_CodeModulImporter
End Property

'Create tabelle usys_DbmsConnection
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

'Check if login form is available
Public Sub CheckLoginForm()
   
   Dim rst As DAO.Recordset
   Dim bolMissingForm As Boolean

   Set rst = CurrentDb.OpenRecordset("Select O.Name from MSysObjects O where O.Name = '" & DCW_LoginFormName & "' and O.Type=-32768", dbOpenForwardOnly, dbReadOnly)
   bolMissingForm = rst.EOF
   rst.Close
   Set rst = Nothing
      
   If bolMissingForm Then
      CodeModulImporter.TransferCodeModul CurrentProject, acForm, DCW_LoginFormName
   End If

End Sub

'Renew modules + classes
Public Function ReplaceCodeModules(ParamArray sModulName() As Variant) As Boolean
   
   Dim i As Long
   Dim ArrSize As Long
   Dim vbp As Object 'VBProject
 
   Set vbp = CurrentVbProject

   If Not (vbp Is Nothing) Then
      ArrSize = UBound(sModulName)
      For i = 0 To ArrSize
      
         ' 1. delete
         If CheckCodeModule(sModulName(i)) Then
            vbp.VBComponents.Remove vbp.VBComponents(sModulName(i))
         End If
         
         ' 2. check => insert
         CheckCodeModule sModulName(i), True
         
      Next
      ReplaceCodeModules = True
   End If
   
   Set vbp = Nothing

End Function


'Check modules and classes for existence. There is no content check!
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

'Check module or class for existence. There is no content check!
Public Function CheckCodeModule(ByVal sModulName As String, _
                       Optional ByVal TransferMissingModule As Boolean = False) As Boolean
   
   Dim rst As DAO.Recordset
   Dim bolMissingModule As Boolean

   Set rst = CurrentDb.OpenRecordset("Select O.Name from MSysObjects O where O.Name = '" & sModulName & "' and O.Type=-32761", dbOpenForwardOnly, dbReadOnly)
   bolMissingModule = rst.EOF
   rst.Close
   Set rst = Nothing
   
   If bolMissingModule And TransferMissingModule Then
      CodeModulImporter.TransferCodeModul CurrentProject, acModule, sModulName
      bolMissingModule = False
   End If
   
   CheckCodeModule = Not bolMissingModule

End Function
