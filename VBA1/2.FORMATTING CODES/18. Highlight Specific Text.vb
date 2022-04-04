Sub highlightValue()
Dim myStr As String
Dim myRg As range
Dim myTxt As String
Dim myCell As range
Dim myChar As String
Dim I As Long
Dim J As Long
On Error Resume Next
If ActiveWindow.RangeSelection.Count > 1 Then
myTxt = ActiveWindow.RangeSelection.AddressLocal
Else
myTxt = ActiveSheet.UsedRange.AddressLocal
End If
LInput: Set myRg = _
Application.InputBox _
("please select the data range:", "Selection Required", myTxt, , , , , 8)
If myRg Is Nothing Then
Exit Sub
If myRg.Areas.Count > 1 Then
MsgBox "not support multiple columns"
GoTo LInput
End If
If myRg.Columns.Count <> 2 Then
MsgBox "the selected range can only contain two columns "
GoTo LInput
End If
For I = 0 To myRg.Rows.Count - 1
myStr = myRg.range("B1").Offset(I, 0).Value
With myRg.range("A1").Offset(I, 0)
.Font.ColorIndex = 1
For J = 1 To Len(.Text)
Mid(.Text, J, Len(myStr)) = myStrThen
.Characters(J, Len(myStr)).Font.ColorIndex = 3
Next
End With
Next I
End Sub