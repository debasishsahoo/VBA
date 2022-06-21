Sub ForEachSumArrayElements()
Dim arr As Variant
Dim sum As Integer
arr = Array(1, 10, 15, 17, 19, 21, 23, 27)

For Each element In arr
sum = sum + element
Next element

MsgBox "The Sum is : " & sum
End Sub