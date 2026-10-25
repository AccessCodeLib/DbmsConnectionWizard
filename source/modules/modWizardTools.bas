Attribute VB_Name = "modWizardTools"
Option Compare Database
Option Explicit

Public Function CheckApplicationStartUpMethod()
   If CurrentDb.Name Like "*.accda" Then
      MsgBox "The add-in must be installed into the Access add-in directory using the add-in manager. Afterwards it has to be started via the menu entry '" & APPLICATION_NAME & "'.", _
             vbExclamation, APPLICATION_NAME & ": Incorrect start"

      Application.Quit
   End If
End Function
