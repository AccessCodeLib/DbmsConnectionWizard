VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "_ADFInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit

Public Property Get GUID() As String
   GUID = ADF_GUID
End Property

Public Sub Start()
   StartWizard
End Sub
