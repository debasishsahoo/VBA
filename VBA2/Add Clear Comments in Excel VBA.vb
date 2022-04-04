'how to add and clear comments using Excel VBA
Sub sbAddComment()
    'Deletes Existing Comments
    Range("A3").ClearComments
    
    'Creates Comment
    Range("A3").AddComment
    Range("A3").Comment.Text Text:="This is Example Comment Text"

End Sub

'we general write the comments in another set of range and add using VBA.

Sub sbAddComment_Example()
    For iCntr = 1 To 30
        'Clear if any existing comments
        Range("A3").ClearComments
    
        'Add a Comment from Column B
        Range("A" & iCntr).AddComment
        Range("A" & iCntr).Comment.Text Text:=Range("B" & iCntr).Value
    
    Next iCntr
End Sub