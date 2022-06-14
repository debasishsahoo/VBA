Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
objSelection.TypeText "ABCDEFGHIJKLM"

objDoc.Saved = TRUE
objWord.Quit