Sub ZoomOut()
    Dim ZP As Integer
    ZP = Int(ActiveWindow.ActivePane.View.Zoom.Percentage * 0.9)
    If ZP < 10 Then ZP = 10
    ActiveWindow.ActivePane.View.Zoom.Percentage = ZP
End Sub