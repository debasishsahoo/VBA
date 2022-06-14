Dim begin As Range
Set startMark = ActiveDocument.Range(0, 0)
'range as Range, NumRows as Long, NumColumns as Long
Call ActiveDocument.Tables.Add(startMark, 3, 6) 'add to beginning of doc, 3 rows, 6 cols
