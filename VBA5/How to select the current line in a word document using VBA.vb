'Select the current line in a word document
Sub Select_Line_Document()
    Selection.HomeKey Unit:=wdLine
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
End Sub