If Not (Selection.PreviousField Is Nothing) Then 
 Selection.Fields.Update 
End If


Set myField = Selection.PreviousField 
If Not (myField Is Nothing) Then StatusBar = "Field found"