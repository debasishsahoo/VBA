Sub AddToolbarItem() 
    Dim btnNew As CommandBarButton 
    CustomizationContext = NormalTemplate 
    Set btnNew = CommandBars("Standard").Controls.Add _ 
        (Type:=msoControlButton, ID:=792, Before:=6) 
    With btnNew 
        .BeginGroup = True 
        .FaceId = 700 
        .TooltipText = "Word Count" 
    End With 
End Sub

Sub AddDoubleUnderlineButton() 
    CustomizationContext = NormalTemplate 
    CommandBars("Formatting").Controls.Add _ 
        Type:=msoControlButton, ID:=60, Before:=7 
End Sub