Set wrd = CreateObject("Word.Application")

wrd.Documents.Add


Set wrd = CreateObject("Word.Application") 
MsgBox wrd.Options.DefaultFilePath(wdStartupPath) 
wrd.Quit