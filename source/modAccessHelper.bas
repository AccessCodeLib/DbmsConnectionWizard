Attribute VB_Name = "modAccessHelper"
'---------------------------------------------------------------------------------------
' Modul: modAccessHelper
'---------------------------------------------------------------------------------------
'/*
' <summary>
' Allgemeine Hilfsfunktionen
' </summary>
' <remarks></remarks>
'*/
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

Public Property Get CurrentVbProject() As Object

   Dim vbp As Object 'VBProject
   Dim strCurrentdbFile As String
   
On Error Resume Next
   
   strCurrentdbFile = UncPath(CurrentDb.Name)

   'VBProject der Anwendung suchen:
   For Each vbp In CurrentProject.Application.VBE.VBProjects
      If vbp.FileName = strCurrentdbFile Then
         Set CurrentVbProject = vbp
         Exit For
      End If
   Next
   
End Property

Private Function CheckReferenzes() As Boolean

   Dim bolCurRefOK As Boolean
   Dim bolRefBroken As Boolean
   Dim R As Reference

   For Each R In Application.References
      bolCurRefOK = False
      bolCurRefOK = R.IsBroken
      If bolCurRefOK Then
        bolRefBroken = True
      End If
   Next

   If bolRefBroken Then
      MsgBox "Es gibt einen Fehler in den Verweisen."
      CheckReferenzes = False
   Else
      CheckReferenzes = True
   End If

End Function

Public Function NullIf(ByVal Value As Variant, ByVal Expression As Variant) As Variant

   If Value = Expression Then
      NullIf = Null
   Else
      NullIf = Value
   End If

End Function

Public Function CurrentAccessVersion() As Long

   Static lngVersion As Long

On Error Resume Next

   If lngVersion = 0 Then
      lngVersion = Val(SysCmd(acSysCmdAccessVer))
   End If
   CurrentAccessVersion = lngVersion
   
End Function
