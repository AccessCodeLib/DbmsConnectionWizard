VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdbcHandlerEventBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents m_OdbcHandler As OdbcHandler
Attribute m_OdbcHandler.VB_VarHelpID = -1
Private m_HookObserver As Object

Public Sub Init(ByVal OdbcHandlerRef As OdbcHandler, ByVal HookObserver As Object)
   Set m_OdbcHandler = OdbcHandlerRef
   Set m_HookObserver = HookObserver
End Sub

Private Sub m_OdbcHandler_RelinkTableDefSkipped(ByVal TdfRef As DAO.TableDef, ByVal SkipMessage As String)
   m_HookObserver.RelinkTableDefSkipped TdfRef, SkipMessage
End Sub

Private Sub m_OdbcHandler_TableDefLinked(ByVal TdfRef As DAO.TableDef)
   m_HookObserver.TableDefLinked TdfRef
End Sub
