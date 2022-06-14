Sub Demo()
Dim udtOrder As Order
With udtOrder
    .Client = ""
    .EntryDate = Date
    .Product = ""
    .Attention = True
End With
ufmOrder.FillList "cboProduct", Array("v1", "v2", "v3")
ufmOrder.SetValues udtOrder
ufmOrder.Show
If Not ufmOrder.IsCancelled Then
    ufmOrder.GetValues udtOrder
    ''continue process after OK here
    With udtOrder

    End With
End If
Unload ufmOrder
End Sub