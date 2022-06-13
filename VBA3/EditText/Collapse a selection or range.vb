Sub CollapseToBeginning() 
 Selection.Collapse Direction:=wdCollapseStart 
End Sub

Sub CollapseToEnd() 
 Dim rngWords As Range 
 
 Set rngWords = ActiveDocument.Words(1) 
 With rngWords 
 .Collapse Direction:=wdCollapseEnd 
 .Text = "(This is a test.) " 
 End With 
End Sub