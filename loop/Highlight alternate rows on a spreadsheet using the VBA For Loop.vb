Sub HighlightAlternateRows()
Dim loop_ctr As Integer
Dim Max As Integer
Dim clm As Integer
Max = ActiveSheet.UsedRange.Rows.Count
clm = ActiveSheet.UsedRange.Columns.Count

For loop_ctr = 1 To Max
If loop_ctr Mod 2 = 0 Then
ActiveSheet.Range(Cells(loop_ctr, 1), Cells(loop_ctr, clm)).Interior.ColorIndex = 28
End If
Next loop_ctr

MsgBox "For Loop Completed!"
End Sub