Sub convertToValues()
Dim MyRange As Range
Dim MyCell As Range
Select Case _
MsgBox("You Can't Undo This Action. " _
& "Save Workbook First?", vbYesNoCancel, _
"Alert")
Case Is = vbYes
ThisWorkbook.Save
Case Is = vbCancel
Exit Sub
End Select
Set MyRange = Selection
For Each MyCell In MyRange
If MyCell.HasFormula Then
MyCell.Formula = MyCell.Value
End If
Next MyCell
End Sub