Sub ForLoopWithArrays()
Dim arr() As Variant
arr = Array(10, 12, 8, 19, 21, 5, 16)

Dim min_number As Integer
Dim max_number As Integer

min_number = arr(0)
max_number = arr(0)

Dim loop_ctr As Integer
For loop_ctr = LBound(arr) To UBound(arr)
If arr(loop_ctr) > max_number Then
max_number = arr(loop_ctr)
End If

If arr(loop_ctr) < min_number Then
min_number = arr(loop_ctr)
End If

Next loop_ctr
MsgBox "Largest Number: " & max_number _
& " Smallest Number: " & min_number
End Sub