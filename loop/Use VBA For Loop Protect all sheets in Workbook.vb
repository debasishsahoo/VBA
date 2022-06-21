Sub ProtectWorksheets()
Dim loop_ctr As Integer
For loop_ctr = 1 To ActiveWorkbook.Worksheets.Count
Worksheets(loop_ctr).Protect
Next loop_ctr
End Sub