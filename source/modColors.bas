Attribute VB_Name = "modColors"
Option Compare Text
Option Explicit
Option Private Module

Private Const m_Ac07_BackgroundForm As Long = -2147483613
Private Const m_Ac07_BackgroundLightHeader As Long = -2147483612
Private Const m_Ac07_BackgroundDarkHeader As Long = -2147483611
Private Const m_Ac07_BordersGridlines As Long = -2147483609
'

' Farben

Public Property Get dcwBackgroundForm() As Long
   
   If CurrentAccessVersion >= 12 Then
      dcwBackgroundForm = m_Ac07_BackgroundForm
   Else
      dcwBackgroundForm = VBA.SystemColorConstants.vbButtonFace
   End If
   
End Property

Public Property Get dcwBackgroundTextboxEdit() As Long
   
   dcwBackgroundTextboxEdit = VBA.SystemColorConstants.vbWindowBackground

End Property

'
Public Property Get dcwBackgroundLightHeader() As Long
   
   If CurrentAccessVersion >= 12 Then
      dcwBackgroundLightHeader = m_Ac07_BackgroundLightHeader
   Else
      dcwBackgroundLightHeader = VBA.SystemColorConstants.vbButtonFace
   End If
   
End Property

Public Property Get dcwBackgroundDarkHeader() As Long
   
   If CurrentAccessVersion >= 12 Then
      dcwBackgroundDarkHeader = m_Ac07_BackgroundDarkHeader
   Else
      dcwBackgroundDarkHeader = VBA.SystemColorConstants.vbButtonFace
   End If
   
End Property

Public Property Get dcwInfoLabelBackColor() As Long
   
   If CurrentAccessVersion >= 12 Then
      dcwInfoLabelBackColor = m_Ac07_BackgroundLightHeader
   Else
      dcwInfoLabelBackColor = &HE6E6E6
   End If
   
End Property

Public Property Get dcwControlBorderColor() As Long
   
   If CurrentAccessVersion >= 12 Then
      dcwControlBorderColor = m_Ac07_BordersGridlines
   Else
      dcwControlBorderColor = 8421504
   End If
   
End Property


' Farben für Vorlagen-Anzeige
Public Property Get dcwBackgroundFormTemplates() As Long
   
   dcwBackgroundFormTemplates = &HEDEDF9

End Property

Public Property Get dcwBackgroundLightHeaderTemplates() As Long
   
   dcwBackgroundLightHeaderTemplates = &HEDEDF9
   
End Property
