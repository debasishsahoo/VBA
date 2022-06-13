Selection.Collapse Direction:=wdCollapseEnd 
Selection.InsertFile FileName:="C:\TEST.DOC", Link:=True


Documents.Add 
ChDir "C:\TMP" 
myName = Dir("*.TXT") 
While myName <> "" 
 With Selection 
 .InsertFile FileName:=myName, ConfirmConversions:=False 
 .InsertParagraphAfter 
 .InsertBreak Type:=wdSectionBreakNextPage 
 .Collapse Direction:=wdCollapseEnd 
 End With 
 myName = Dir() 
Wend