Sub HideSubtotals()
Dim pt As PivotTable
Dim pf As PivotField
On Error Resume Next
Set pt = ActiveSheet.PivotTables(ActiveCell.PivotTable.name)
If pt Is Nothing Then
MsgBox "You must place your cursor inside of a PivotTable."
Exit Sub
End If
For Each pf In pt.PivotFields
pf.Subtotals(1) = True
pf.Subtotals(1) = False
Next pf
End Sub