ActiveDocument.Paragraphs(2).Range.LanguageID = wdFrench 
Set myDictionary = CustomDictionaries.Add(Filename:="French.dic") 
With myDictionary 
 .LanguageSpecific = True 
 .LanguageID = wdFrench 
End With