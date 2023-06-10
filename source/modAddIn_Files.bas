Attribute VB_Name = "modAddIn_Files"
Option Compare Text
Option Explicit
Option Private Module

#Const UseZip = 0

Private Declare PtrSafe Function API_GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" ( _
         ByVal lpszPath As String, _
         ByVal lpPrefixString As String, _
         ByVal wUnique As Long, _
         ByVal lpTempFileName As String) As Long
         
Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
         ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long

Private Const m_conMaxPathLen As Long = 255

Private Sub dateienEinstellen()
   SaveModulesInTable
End Sub

'General directory for user settings and changeable data:
Public Property Get DbmsConnectionWizardAppDataDirectory() As String
   
   Dim strPath As String

   strPath = Environ("Appdata") & "\DbmsConnectionWizard"
   If Len(Dir$(strPath, vbDirectory)) = 0 Then
      MkDir strPath
   End If
   DbmsConnectionWizardAppDataDirectory = strPath

End Property

Public Property Get AddInHelpFile() As String

   Dim strHelpFile As String
   Dim rst As DAO.Recordset
   Dim strVersion As String

   strHelpFile = DbmsConnectionWizardAppDataDirectory & "\DbmsConnectionWizard.chm"
   
   Set rst = CodeDb.OpenRecordset("select Version from usys_AppFiles where ID='ChmFile'")
   If Not rst.EOF Then
      strVersion = Nz(rst.Fields(0), vbNullString)
   End If
   rst.Close
   Set rst = Nothing
   
   CheckLibFile "ChmFile", strHelpFile, strVersion
   
   AddInHelpFile = strHelpFile
   
End Property

Public Sub OpenHelpFile(Optional ByVal ContextID As Long = 0)
' @todo: create new help file
   If ContextID = 0 Then
      ShellExecuteOpenFile AddInHelpFile
   Else
      WizHook.Key = 51488399
      WizHook.WizHelp AddInHelpFile, 1, ContextID
   End If
   
End Sub

Public Function CreateAppFile(ByVal sFileID As String, ByVal sFileName As String, _
                     Optional ByVal decode As Boolean = False) As Boolean

   Dim f As Integer
   Dim Binfile() As Byte
   Dim lngFieldSize As Long
   Dim rst As DAO.Recordset
   Dim fld As DAO.Field

   Dim ZipFile As String
   Dim TempFile As String
   Dim tempDir As String

   Set rst = CodeDb.OpenRecordset("select File from usys_AppFiles where ID='" & sFileID & "'")
   If rst.EOF Then
      CreateAppFile = False
   Else
      Set fld = rst.Fields(0)
      lngFieldSize = fld.FieldSize
      If lngFieldSize > 0 Then

         ReDim Binfile(lngFieldSize - 1)
         Binfile = fld.GetChunk(0, lngFieldSize)
         
         If decode Then
            CodeDecodeByteArray Binfile, sFileID
         End If
                  
#If UseZip Then
         
         ZipFile = GetTempFileName(, , "zip")
         f = FreeFile
         Open ZipFile For Binary As #f
         Put #f, , Binfile()
         Close #f
         
         tempDir = Left$(ZipFile, InStrRev(ZipFile, "\"))
         TempFile = tempDir & ExtractFromZipFile(ZipFile, tempDir)
         
         FileCopy TempFile, sFileName

         Kill TempFile
         Kill ZipFile
         
#Else
         
         f = FreeFile
         Open sFileName For Binary As #f
         Put #f, , Binfile()
         Close #f
         
#End If
         
         CreateAppFile = True
      Else
         CreateAppFile = False
      End If

   End If

End Function

Private Function CheckLibFile(ByVal sFileID As String, ByVal sFileName As String, _
                     Optional ByVal sSavedFileVersion As String = vbNullString) As Boolean

   Dim installedVersionString As String
   Dim installedVersion() As String
   Dim savedVersion() As String
   Dim i As Long
   Dim NewVersion As Boolean
   Dim check As Boolean
   Dim saveFileName As String
   Dim RetVal As Long

   If Len(Dir$(sFileName)) > 0 Then
   
      check = True
   
      If Len(sSavedFileVersion) > 0 Then

         installedVersionString = modWinAPI_FileInfo.GetFileVersion(sFileName)
         If Len(installedVersionString) = 0 Then
            installedVersionString = Format$(Nz(FileDateTime(sFileName), vbNullString), "yyyy.mm.dd")
         End If
         
         If sSavedFileVersion <> installedVersionString Then
            installedVersion = Split(installedVersionString, ".")
            savedVersion = Split(sSavedFileVersion, ".")
            For i = 0 To UBound(installedVersion)
               If Val(savedVersion(i)) > Val(installedVersion(i)) Then
                  NewVersion = True
                  Exit For
               End If
            Next
            If NewVersion Then

               RetVal = vbYes
               
               If RetVal = vbYes Then
                         
                  'Rename file only (delete only after successful recreation)
                  saveFileName = sFileName & ".saved"
                  If Len(Dir$(saveFileName)) > 0 Then Kill saveFileName
                  Name sFileName As saveFileName
                  
                  'Create new file
                  check = CreateAppFile(sFileID, sFileName)
                  If check Then
                     Kill sFileName & ".saved"
                  Else
                     Name sFileName & ".saved" As sFileName
                  End If

               End If
            End If
         End If
      End If

   Else
      check = CreateAppFile(sFileID, sFileName)
   End If

   CheckLibFile = check

End Function

Public Function CreateModuleFile(ByVal sFileID As String, ByVal sFileName As String) As Boolean

On Error Resume Next

   CreateModuleFile = CreateAppFile(sFileID, sFileName, False)

End Function

'###########################################
'
' Help function for attaching the app files
'
Private Function SaveAppFile(ByVal sFileID As String, ByVal sFileName As String, _
                    Optional ByVal SaveVersion As Boolean = False, _
                    Optional ByVal encode As Boolean = False) As String

   Dim f As Integer
   Dim Binfile() As Byte
   Dim strVersion As String
   Dim lngFileSize As Long

#If UseZip Then

   Dim ZipFile As String
   ZipFile = GetTempFileName(, , "zip")
   AddToZipFile ZipFile, sFileName
   
   f = FreeFile
   Open ZipFile For Binary As #f
   lngFileSize = LOF(f)
   If lngFileSize > 0 Then
      ReDim Binfile(lngFileSize - 1)
      Get #f, , Binfile()
   End If
   Close #f
   
   Kill ZipFile
   
#Else

   f = FreeFile
   Open sFileName For Binary As #f
   lngFileSize = LOF(f)
   If lngFileSize > 0 Then
      ReDim Binfile(LOF(f) - 1)
      Get #f, , Binfile()
   End If
   Close #f

#End If
   
   If encode Then
      CodeDecodeByteArray Binfile, sFileID
   End If

   Dim rst As DAO.Recordset
   Set rst = CodeDb.OpenRecordset("select ID, File, Version from usys_AppFiles where ID='" & sFileID & "'")
   If rst.EOF Then
      rst.AddNew
      rst.Fields("id") = sFileID
   Else
      rst.Edit
   End If
   rst.Fields("file").AppendChunk Binfile
   If SaveVersion Then
      strVersion = modWinAPI_FileInfo.GetFileVersion(sFileName)
      If Len(strVersion) = 0 Then
         strVersion = Format$(Nz(FileDateTime(sFileName), vbNullString), "yyyy.mm.dd")
      End If
      rst.Fields("version") = strVersion
   End If
   rst.Update
   rst.Close
   Set rst = Nothing

End Function

Private Sub SaveModulesInTable()

   Dim X As Variant
   Dim i As Long
   
   X = Array("SqlTools", "defDbConnection", "DbConnectionInfo", "AdodbHandler", "DaoHandler", "OdbcHandler", "DbConnectionHandler", "DbConnectionManager", "modDbConnection")
   For i = 0 To UBound(X)
      SaveCodeModulInTable acModule, X(i)
   Next
   
   X = Array("LoginForm")
   For i = 0 To UBound(X)
      SaveCodeModulInTable acForm, X(i)
   Next
   
End Sub

Private Sub SaveCodeModulInTable(ByVal ObjType As AcObjectType, ByVal sModulName As String, _
                        Optional ByVal encode As Boolean = False)
   
   Dim strFileName As String

   strFileName = GetTempFileName
   Application.SaveAsText ObjType, sModulName, strFileName
   SaveAppFile sModulName, strFileName, True, encode
   
   Kill strFileName
   
End Sub

Public Function GetTempFileName(Optional ByVal TempPath As String = "", _
                         Optional ByVal FilePrefix As String = "", _
                         Optional ByVal FileExtension As String = "") As String

   Dim strTempFileName As String
   Dim strTempPath As String
   Dim lngRet As Long
   
   If Len(TempPath) = 0 Then
      strTempFileName = String$(m_conMaxPathLen, 0)
      lngRet = GetTempPath(m_conMaxPathLen, strTempFileName)
      strTempFileName = Left$(strTempFileName, InStr(strTempFileName, Chr$(0)) - 1)
      strTempPath = strTempFileName
   Else
      strTempPath = TempPath
   End If
   
   strTempFileName = String$(m_conMaxPathLen, 0)
   lngRet = API_GetTempFilename(strTempPath, FilePrefix, 0&, strTempFileName)
   
   strTempFileName = Left$(strTempFileName, InStr(strTempFileName, Chr$(0)) - 1)
   
   'Delete file again, as only name is needed
   Call Kill(strTempFileName)
   
   If Len(FileExtension) > 0 Then
     strTempFileName = Left$(strTempFileName, Len(strTempFileName) - 3) & FileExtension
   End If
   
   GetTempFileName = strTempFileName
  
End Function

Private Sub CodeDecodeByteArray(ByRef TextByteArray() As Byte, ByVal Password As String)

   Dim Key As Byte
   Dim lPos As Long
   Dim lSize As Long
   Dim lenPwd As Long
   Dim i As Long
      
   lSize = UBound(TextByteArray)
   lenPwd = Len(Password)

   For i = 0 To lSize
      lPos = (i + 1) Mod lenPwd
      If lPos = 0 Then lPos = lenPwd
      Key = Asc(Mid$(Password, lPos, 1))
      TextByteArray(i) = (TextByteArray(i) Xor Key)
   Next i
   
End Sub
