Sub ShowRightIndent() 
 Dim dlgParagraph As Dialog 
 Set dlgParagraph = Dialogs(wdDialogFormatParagraph) 
 MsgBox "Right indent = " & dlgParagraph.RightIndent 
End Sub

Sub ShowRightIndexForSelectedParagraph() 
 MsgBox Selection.ParagraphFormat.RightIndent 
End Sub

Sub SetKeepWithNext() 
 With Dialogs(wdDialogFormatParagraph) 
 .KeepWithNext = 1 
 .Execute 
 End With 
End Sub

Sub SetKeepWithNextForSelectedParagraph() 
 Selection.ParagraphFormat.KeepWithNext = True 
End Sub