
Const END_OF_STORY = 6
Const MOVE_SELECTION = 0

Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Open("c:\scripts\word\testdoc.doc")
Set objSelection = objWord.Selection
objSelection.EndKey END_OF_STORY, MOVE_SELECTION

objSelection.TypeParagraph()
objSelection.TypeParagraph()

objSelection.Font.Size = "14"
objSelection.TypeText "" & Date()
objSelection.TypeParagraph()
objSelection.TypeParagraph()

objSelection.Font.Size = "10"