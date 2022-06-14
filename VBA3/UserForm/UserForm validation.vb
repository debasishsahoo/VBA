Public IsCancelled As Boolean

Private Sub UserForm_Initialize()
    IsCancelled = True
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnOk_Click()
    If IsInputOk Then
        IsCancelled = False
        Me.Hide
    End If
End Sub