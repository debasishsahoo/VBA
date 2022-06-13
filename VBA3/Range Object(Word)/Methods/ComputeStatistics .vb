Set myRange = Documents("Report.doc").Paragraphs(1).Range 
wordCount = myRange.ComputeStatistics(Statistic:=wdStatisticWords) 
charCount = myRange.ComputeStatistics(Statistic:=wdStatisticCharacters) 
MsgBox "The first paragraph contains " & wordCount _ 
 & " words and a total of " & charCount & " characters."