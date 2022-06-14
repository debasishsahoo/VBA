Sub AddChart_Excel()
    Dim objShape As Shape
    
    ' Create a chart and return a Shape object reference.
    ' The Shape object reference contains the chart.
    Set objShape = ActiveSheet.Shapes.AddChart(XlChartType.xlColumnStacked100)
    
    ' Ensure the Shape object contains a chart. If so,
    ' set the source data for the chart to the range A1:C3.
    If objShape.HasChart Then
        objShape.Chart.SetSourceData Source:=Range("'Sheet1'!$A$1:$C$3")
    End If
End Sub


Sub AddChart_Word()
    Dim objShape As InlineShape
    
    ' Create a chart and return a Shape object reference.
    ' The Shape object reference contains the chart.
    Set objShape = ActiveDocument.InlineShapes.AddChart(XlChartType.xlColumnStacked100)
    
    ' Ensure the Shape object contains a chart. If so,
    ' set the source data for the chart to the range A1:C3.
    If objShape.HasChart Then
        objShape.Chart.SetSourceData Source:="'Sheet1'!$A$1:$C$3"
    End If
End Sub


Sub ShowWorkbook_Word() 
    Dim objShape As InlineShape 
     
    ' Iterates each inline shape in the active document. 
    ' If the inline shape contains a chart, then display the 
    ' data associated with that chart and minimize the application 
    ' used to display the data. 
    For Each objShape In ActiveDocument.InlineShapes 
        If objShape.HasChart Then 
 
            ' Activate the topmost window of the application used to 
            ' display the data for the chart. 
            objShape.Chart.ChartData.Activate 
             
            ' Minimize the application used to display the data for 
            ' the chart. 
            objShape.Chart.ChartData.Workbook.Application.WindowState = -4140 
        End If 
    Next 
End Sub