Attribute VB_Name = "modWinAPI_FileInfo"
Option Compare Text
Option Explicit
Option Private Module

'Code basiert auf http://support.microsoft.com/kb/509493/


Private Declare PtrSafe Function GetFileVersionInfoSize _
  Lib "version.dll" _
  Alias "GetFileVersionInfoSizeA" _
  ( _
  ByVal lptstrFilename As String, _
  lpdwHandle As LongPtr _
  ) _
As Long


Private Declare PtrSafe Function GetFileVersionInfo _
  Lib "version.dll" _
  Alias "GetFileVersionInfoA" _
  ( _
  ByVal lptstrFilename As String, _
  ByVal dwHandle As LongPtr, _
  ByVal dwLen As Long, _
  lpData As Any _
  ) _
As Long

      
Private Declare PtrSafe Function VerQueryValue _
  Lib "version.dll" _
  Alias "VerQueryValueA" _
  ( _
  pBlock As Any, _
  ByVal lpSubBlock As String, _
  lplpBuffer As Any, _
  puLen As Long _
  ) _
As Long

Private Declare PtrSafe Sub MoveMemory _
  Lib "kernel32" _
  Alias "RtlMoveMemory" _
  ( _
  Dest As Any, _
  ByVal Source As Long, _
  ByVal Length As Long _
  )

Private Type VS_FIXEDFILEINFO
  dwSignature As Long
  dwStrucVersion As Long
  dwFileVersionMS As Long
  dwFileVersionLS As Long
  dwProductVersionMS As Long
  dwProductVersionLS As Long
  dwFileFlagsMask As Long
  dwFileFlags As Long
  dwFileOS As Long
  dwFileType As Long
  dwFileSubtype As Long
  dwFileDateMS As Long
  dwFileDateLS As Long
End Type

Private Type FILEINFOOUT
  FileVersion As String
  ProductVersion As String
End Type

Private Function GetVersion(ByVal sPath As String, _
                           ByRef FInfo As FILEINFOOUT) As Boolean

  Dim lRet As Long, lSize As Long, lHandle As LongPtr
  Dim lVerBufLen As Long, lVerPointer As Long
  Dim FileInfo As VS_FIXEDFILEINFO
  Dim sBuffer() As Byte

  lSize = GetFileVersionInfoSize(sPath, lHandle)
  If lSize = 0 Then
    GetVersion = False
    Exit Function
  End If
  
  ReDim sBuffer(lSize)
  lRet = GetFileVersionInfo(sPath, 0&, lSize, sBuffer(0))
  If lSize = 0 Then
    GetVersion = False
    Exit Function
  End If
  
  lRet = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerBufLen)
  If lSize = 0 Then
    GetVersion = False
    Exit Function
  End If
  
  Call MoveMemory(FileInfo, lVerPointer, Len(FileInfo))
  
  With FileInfo
  
    FInfo.FileVersion = _
      Trim$(Str$((.dwFileVersionMS And &HFFFF0000) \ &H10000)) & "." & _
      Trim$(Str$(.dwFileVersionMS And &HFFFF&)) & "." & _
      Trim$(Str$((.dwFileVersionLS And &HFFFF0000) \ &H10000)) & "." & _
      Trim$(Str$(.dwFileVersionLS And &HFFFF&))
    
    FInfo.ProductVersion = _
      Trim$(Str$((.dwProductVersionMS And &HFFFF0000) \ &H10000)) & "." & _
      Trim$(Str$(.dwProductVersionMS And &HFFFF&)) & "." & _
      Trim$(Str$((.dwProductVersionLS And &HFFFF0000) \ &H10000)) & "." & _
      Trim$(Str$(.dwProductVersionLS And &HFFFF&))
      
  End With
  
  GetVersion = True

End Function


'#####################################################
'
' Ergänzung:
'
Public Function GetFileVersion(ByVal sFile As String) As String
   
   Dim VerInfo As FILEINFOOUT

   If GetVersion(sFile, VerInfo) Then
      GetFileVersion = VerInfo.FileVersion
   Else
      GetFileVersion = vbNullString
   End If

End Function
