MsgBox "There are " & Selection.Words.Count & " words."

Set myRange = ActiveDocument.Range(Start:=0, End:=Selection.End) 
For Each aWord In myRange.Words 
 If aWord.Text = "Franklin " Then aWord.Delete 
Next aWord