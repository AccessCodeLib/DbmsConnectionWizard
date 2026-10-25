Attribute VB_Name = "modMSAccessVcsSupport"
Option Compare Database
Option Explicit

Public Sub VcsRunBeforeExport()
   CurrentApplication.SetApplicationProperty "StartUpForm", vbNullString
End Sub
