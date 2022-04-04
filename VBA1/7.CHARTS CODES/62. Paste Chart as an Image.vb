Sub ConvertChartToPicture()
ActiveChart.ChartArea.Copy
ActiveSheet.Range("A1").Select
ActiveSheet.Pictures.Paste.Select
End Sub