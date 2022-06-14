Private Function IsCorrectType(ctl As MSForms.Control) As Boolean
Dim strControlDataType As String, strMessage As String
Dim dummy As Variant
    strControlDataType = ControlDataType(ctl)
On Error GoTo HandleError
    Select Case strControlDataType
    Case "Boolean"
        dummy = CBool(GetValue(ctl, strControlDataType))
    Case "Byte"
        dummy = CByte(GetValue(ctl, strControlDataType))
    Case "Currency"
        dummy = CCur(GetValue(ctl, strControlDataType))
    Case "Date"
        dummy = CDate(GetValue(ctl, strControlDataType))
    Case "Double"
        dummy = CDbl(GetValue(ctl, strControlDataType))
    Case "Decimal"
        dummy = CDec(GetValue(ctl, strControlDataType))
    Case "Integer"
        dummy = CInt(GetValue(ctl, strControlDataType))
    Case "Long"
        dummy = CLng(GetValue(ctl, strControlDataType))
    Case "Single"
        dummy = CSng(GetValue(ctl, strControlDataType))
    Case "String"
        dummy = CStr(GetValue(ctl, strControlDataType))
    Case "Variant"
        dummy = CVar(GetValue(ctl, strControlDataType))
    End Select
    IsCorrectType = True
HandleExit:
    Exit Function
HandleError:
    IsCorrectType = False
    Resume HandleExit
End Function

Private Function ControlDataType(ctl As MSForms.Control) As String
    Select Case ctl.Name
    Case "txtClient": ControlDataType = "String"
    Case "txtEntryDate": ControlDataType = "Date"
    Case "cboProduct": ControlDataType = "String"
    Case "cbxAttention": ControlDataType = "Boolean"
    End Select
End Function
Private Function IsRequired(ctl As MSForms.Control) As Boolean
    Select Case ctl.Name
    Case "txtClient", "txtEntryDate", "cboProduct", "cbxAttention"
        IsRequired = True
    Case Else
        IsRequired = False
    End Select
End Function