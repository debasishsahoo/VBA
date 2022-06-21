Sub SumFirst15OddNumbers()
Dim loop_ctr As Integer
Dim odd_number_counter As Integer
Dim sum As Integer

loop_ctr = 1

Do While loop_ctr <= 100
If (loop_ctr Mod 2 <> 0) Then
sum = sum + loop_ctr
odd_number_counter = odd_number_counter + 1
End If

If (odd_number_counter = 15) Then
Exit Do
End If

loop_ctr = loop_ctr + 1
Loop

MsgBox "Sum of top 15 odd numbers is : " & sum
End Sub