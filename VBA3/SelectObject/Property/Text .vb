MsgBox Selection.Text

Documents.Add 
For i = 1 To 10 
 Selection.Text = "Line" & Str(i) & Chr(13) 
 Selection.MoveDown Unit:=wdParagraph, Count:=1 
Next i