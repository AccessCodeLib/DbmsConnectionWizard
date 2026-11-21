Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const m_conMaxCtlCnt As Long = 30

Public Sub SetFormData(ByRef rst As Object) 'rst kann ADODB oder DAO-Recordset sein!

    Dim lngRstFldCnt As Long
    Dim lngMaxCnt As Long
    Dim i As Long

    lngRstFldCnt = rst.Fields.Count

    If m_conMaxCtlCnt > lngRstFldCnt Then
        lngMaxCnt = lngRstFldCnt - 1 'da von 0 gezählt wird
    Else
        lngMaxCnt = m_conMaxCtlCnt - 1 'es werden nicht alle Felder des Recordset übernommen
    End If

    Me!Text0.Enabled = True 'nur zur Sicherheit
    Me!Text0.SetFocus

    With rst
    For i = 0 To lngMaxCnt
        Me.Controls("lab" & i).Caption = .Fields(i).Name
        Me.Controls("Text" & i).ControlSource = .Fields(i).Name
        Me.Controls("Text" & i).Enabled = True
    Next i
    End With

    For i = lngMaxCnt + 1 To m_conMaxCtlCnt - 1
        Me.Controls("lab" & i).Caption = vbNullString
        Me.Controls("Text" & i).ControlSource = vbNullString
        Me.Controls("Text" & i).Enabled = False
    Next i

    Set Me.Recordset = rst

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   Set Me.Recordset = Nothing
End Sub
