Selection.Copy
'---------------------------------------
Documents(3).ActiveWindow.Selection.Cut
'---------------------------------------
ActiveDocument.ActiveWindow.Panes(1).Selection.Copy 
ActiveDocument.ActiveWindow.Panes(2).Selection.Paste
'----------------------------------------
Dim strTemp as String 
 
strTemp = Selection.Text 
If Right(strTemp, 1) = vbCr Then _ 
 strTemp = Left(strTemp, Len(strTemp) - 1)
'----------------------------------------
Selection.EndOf Unit:=wdStory, Extend:=wdMove 
Selection.HomeKey Unit:=wdLine, Extend:=wdExtend 
Selection.MoveUp Unit:=wdLine, Count:=2, Extend:=wdExtend

'-------------------------------------------
Options.ReplaceSelection = True 
ActiveDocument.Sentences(1).Select 
Selection.TypeText "Material below is confidential." 
Selection.TypeParagraph
'--------------------------------------------
With Documents(1) 
 .Paragraphs.Last.Range.Select 
 .ActiveWindow.Selection.Cut 
End With 
 
With Documents(2).ActiveWindow.Selection 
 .StartOf Unit:=wdStory, Extend:=wdMove 
 .Paste 
End With
'---------------------------------------------
If Selection.Font.Name = "Times New Roman" Then _ 
 Selection.Font.Name = "Tahoma"
'--------------------------------------------
If Selection.Type = wdSelectionIP Then 
 MsgBox Prompt:="You have not selected any text! Exiting procedure..." 
 Exit Sub 
End If
'--------------------------------------------
If Selection.Type <> wdSelectionNormal Then 
 MsgBox Prompt:="Not a valid selection! Exiting procedure..." 
 Exit Sub 
End If