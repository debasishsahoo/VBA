Sub Word_Count_Worksheet()
Dim WordCnt As Long
Dim rng As Range
Dim S As String
Dim N As Long
For Each rng In ActiveSheet.UsedRange.Cells
S = Application.WorksheetFunction.Trim(rng.Text)
N = 0
If S <> vbNullString Then
N = Len(S) - Len(Replace(S, " ", "")) + 1
End If
WordCnt = WordCnt + N
Next rng
MsgBox "There are total " _
& Format(WordCnt, "#,##0") & _
" words in the active worksheet"
End Sub