Private Sub SetValue(ctl As MSForms.Control, value As Variant)
On Error GoTo HandleError
    ctl.value = value
HandleExit:
    Exit Sub
HandleError:
    Resume HandleExit
End Sub

Private Function GetValue(ctl As MSForms.Control, strTypeName As String) As Variant
On Error GoTo HandleError
    Dim value As Variant
    value = ctl.value
    If IsNull(value) And strTypeName <> "Variant" Then
        Select Case strTypeName
        Case "String"
            value = ""
        Case Else
            value = 0
        End Select
    End If
HandleExit:
    GetValue = value
    Exit Function
HandleError:
    Resume HandleExit
End Function



Private Function IsInputOk() As Boolean
Dim ctl As MSForms.Control
Dim strMessage As String
    IsInputOk = False
    For Each ctl In Me.Controls
        If IsInputControl(ctl) Then
            If IsRequired(ctl) Then
                If Not HasValue(ctl) Then
                    strMessage = ControlName(ctl) & " must have value"
                End If
            End If
            If Not IsCorrectType(ctl) Then
                strMessage = ControlName(ctl) & " is not correct"
            End If
        End If
        If Len(strMessage) > 0 Then
            ctl.SetFocus
            GoTo HandleMessage
        End If
    Next
    IsInputOk = True
HandleExit:
    Exit Function
HandleMessage:
    MsgBox strMessage
    GoTo HandleExit
End Function