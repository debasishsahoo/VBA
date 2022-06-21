Dim Words, Chars, MyString 
For Words = 10 To 1 Step -1 ' Set up 10 repetitions. 
 For Chars = 0 To 9 ' Set up 10 repetitions. 
 MyString = MyString & Chars ' Append number to string. 
 Next Chars ' Increment counter 
 MyString = MyString & " " ' Append a space. 
Next Words