
Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objShape = objDoc.Shapes
objShape.AddPicture("C:\Scripts\Logo.jpg")