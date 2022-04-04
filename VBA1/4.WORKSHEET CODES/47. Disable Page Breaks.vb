Sub DisablePageBreaks()
Dim wb As Workbook
Dim wks As Worksheet
Application.ScreenUpdating = False
For Each wb In Application.Workbooks
For Each Sht In wb.Worksheets
Sht.DisplayPageBreaks = False
Next Sht
Next wb
Application.ScreenUpdating = True
End Sub