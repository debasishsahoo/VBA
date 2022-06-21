Sub SumFirst20EvenNumbers()
Dim loop_ctr As Integer
Dim even_number_counter As Integer
Dim sum As Integer

loop_ctr = 1

Do Until loop_ctr < 100
If (loop_ctr Mod 2 = 0) Then
sum = sum + loop_ctr
even_number_counter = even_number_counter + 1
End If

If (even_number_counter = 20) Then
Exit Do
End If

loop_ctr = loop_ctr + 1
Loop

MsgBox "Sum of top 20 even numbers is : " & sum
End Sub