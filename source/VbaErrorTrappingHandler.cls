VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaErrorTrappingHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const ErrorTrappingOptionName As String = "Error Trapping"

Public Enum VbaErrorTrappingOptions
   BreakOnAllErrors = 0
   BreakInClassModule = 1
   BreakOnUnhandledErrors = 2
End Enum

Private m_ValueOnClassInit As VbaErrorTrappingOptions
Private m_KeepNewValueOnTerminate As Boolean

Private Sub Class_Initialize()
   m_ValueOnClassInit = Application.GetOption(ErrorTrappingOptionName)
End Sub

Private Sub Class_Terminate()
   If Not m_KeepNewValueOnTerminate Then
   If ErrorTrapping <> m_ValueOnClassInit Then
      Application.SetOption ErrorTrappingOptionName, m_ValueOnClassInit
   End If
   End If
End Sub

Public Property Get KeepNewValueOnTerminate() As Boolean
   KeepNewValueOnTerminate = m_KeepNewValueOnTerminate
End Property

Public Property Let KeepNewValueOnTerminate(ByVal NewValue As Boolean)
   m_KeepNewValueOnTerminate = NewValue
End Property

Public Property Get ErrorTrapping() As VbaErrorTrappingOptions
   ErrorTrapping = Application.GetOption(ErrorTrappingOptionName)
End Property

Public Property Let ErrorTrapping(ByVal NewValue As VbaErrorTrappingOptions)
   If ErrorTrapping <> NewValue Then
      Application.SetOption ErrorTrappingOptionName, NewValue
   End If
End Property
