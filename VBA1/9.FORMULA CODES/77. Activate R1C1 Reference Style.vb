Sub ActivateR1C1()
If Application.ReferenceStyle = xlA1 Then
Application.ReferenceStyle = xlR1C1
Else
Application.ReferenceStyle = xlR1C1
End If
End Sub