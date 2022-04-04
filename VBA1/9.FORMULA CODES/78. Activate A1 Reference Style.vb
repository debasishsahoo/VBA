Sub ActivateA1()
If Application.ReferenceStyle = xlR1C1 Then
Application.ReferenceStyle = xlA1
Else
Application.ReferenceStyle = xlA1
End If
End Sub