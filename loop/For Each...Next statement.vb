Dim Found, MyObject, MyCollection 
Found = False    ' Initialize variable. 
For Each MyObject In MyCollection    ' Iterate through each element.  
    If MyObject.Text = "Hello" Then    ' If Text equals "Hello". 
        Found = True    ' Set Found to True. 
        Exit For    ' Exit loop. 
    End If 
Next