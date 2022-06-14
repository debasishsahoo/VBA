Private Sub GetUserName() 
 With UserForm1 
 .lstRegions.AddItem "North" 
 .lstRegions.AddItem "South" 
 .lstRegions.AddItem "East" 
 .lstRegions.AddItem "West" 
 .txtSalesPersonID.Text = "00000" 
 .Show 
 ' ... 
 End With 
End Sub

Private Sub UserForm_Initialize() 
 With UserForm1 
 With .lstRegions 
 .AddItem "North" 
 .AddItem "South" 
 .AddItem "East" 
 .AddItem "West" 
 End With 
 .txtSalesPersonID.Text = "00000" 
 End With 
End Sub