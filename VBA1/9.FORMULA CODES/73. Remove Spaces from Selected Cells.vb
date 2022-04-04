Sub RemoveSpaces()
Dim myRange As Range
Dim myCell As Range
Select Case MsgBox("You Can't Undo This Action. " _
& "Save Workbook First?", _
vbYesNoCancel, "Alert")
Case Is = vbYesThisWorkbook.Save
Case Is = vbCancel
Exit Sub
End Select
Set myRange = Selection
For Each myCell In myRange
If Not IsEmpty(myCell) Then
myCell = Trim(myCell)
End If
Next myCell
End Sub