Attribute VB_Name = "_ADF_Helper"
Option Compare Text
Option Explicit

Public Const ADF_GUID As String = "{789DDE9D-EFBA-4F22-8387-3D826590F302}"

Public Function GetAddIn() As [_ADFInterface]
   GetAddIn = New [_ADFInterface]
End Function

Public Sub StartWizard()
   StartDbmsConnectionWizard
End Sub
