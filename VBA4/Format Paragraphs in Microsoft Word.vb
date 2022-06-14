Const CENTERED = 1

Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.ParagraphFormat.Alignment = CENTERED
objSelection.ParagraphFormat.LineSpacing = 36

objSelection.Font.Name = "Arial"
objSelection.Font.Size = "18"
objSelection.TypeText "Here is some text typed in. "
objSelection.TypeText "Here is some more text typed in. "
objSelection.TypeText "Here is even more text typed in. "
objSelection.TypeText "This is the last of the text."