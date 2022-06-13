With Selection 
 ' Collapse current selection to insertion point. 
 .Collapse 
 ' Turn extend mode on. 
 .Extend 
 ' Extend selection to word. 
 .Extend 
 ' Extend selection to sentence. 
 .Extend 
End With

With Selection 
 ' Collapse current selection. 
 .Collapse 
 ' Expand selection to current sentence. 
 .Expand Unit:=wdSentence 
End With

With Selection 
 .StartIsActive = False 
 .Extend Character:="R" 
End Wit