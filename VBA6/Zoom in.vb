Sub ZoomIn()
    Dim ZP As Integer
    ZP = Int(ActiveWindow.ActivePane.View.Zoom.Percentage * 1.1)
    If ZP > 200 Then ZP = 200
    ActiveWindow.ActivePane.View.Zoom.Percentage = ZP
End Sub