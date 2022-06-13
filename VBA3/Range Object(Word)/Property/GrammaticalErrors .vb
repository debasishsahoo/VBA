Set myErrors = ActiveDocument.Paragraphs(3).Range.GrammaticalErrors 
For Each myerr In myErrors 
 MsgBox myerr.Text 
Next myerr