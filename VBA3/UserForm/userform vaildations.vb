Public Sub SetValues(udtOrder As Order)
    With udtOrder
        SetValue Me.txtClient, .Client
        SetValue Me.txtEntryDate, .EntryDate
        SetValue Me.cboProduct, .Product
        SetValue Me.cbxAttention, .Attention
    End With
End Sub

Public Sub GetValues(ByRef udtOrder As Order)
    With udtOrder
        .Client = GetValue(Me.txtClient, TypeName(.Client))
        .EntryDate = GetValue(Me.txtEntryDate, TypeName(.EntryDate))
        .Product = GetValue(Me.cboProduct, TypeName(.Product))
        .Attention = GetValue(Me.cbxAttention, TypeName(.Attention))
    End With
End Sub