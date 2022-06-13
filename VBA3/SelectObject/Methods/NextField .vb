If Not (Selection.NextField Is Nothing) Then 
 Selection.Fields.Update 
End If

Set myField = Selection.NextField 
If Not (myField Is Nothing) Then StatusBar = "Field found"