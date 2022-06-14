' Determine whether selection includes final paragraph mark

Dim SelectionIncludesFinalParagraphMark As Boolean

If Selection.Type = wdSelectionNormal _

And Selection.End = ActiveDocument.Content.End _

Then

SelectionIncludesFinalParagraphMark = True