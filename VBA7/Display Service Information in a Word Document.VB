Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.TypeText "Services Report"
objSelection.TypeParagraph()
objSelection.TypeText "" & Now
objSelection.TypeParagraph()
objSelection.TypeParagraph()

strComputer = "."
Set objWMIService = _
    GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Service")

For Each objItem in colItems
    objSelection.TypeText objItem.DisplayName & " -- " & objItem.State
    objSelection.TypeParagraph()
Next