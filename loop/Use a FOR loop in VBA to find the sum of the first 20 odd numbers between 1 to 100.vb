Sub SumFirst20OddNumbers()
Dim loop_ctr As Integer
Dim odd_number_counter As Integer
Dim sum As Integer

For loop_ctr = 1 To 100
If (loop_ctr Mod 2 <> 0) Then
sum = sum + loop_ctr
odd_number_counter = odd_number_counter + 1
End If

If (odd_number_counter = 20) Then
Exit For
End If
Next loop_ctr

MsgBox "Sum of top 20 odd numbers is : " & sum
End Sub