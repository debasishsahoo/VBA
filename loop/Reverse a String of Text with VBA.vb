Option Explicit
 
Private Sub CommandButton1_Click()
 
'Define Variables
 
Dim Original_String As String
Dim Reversed_String As String
Dim Next_Char As String
 
Dim Length As Integer
Dim Pos As Integer
 
'Get the Original String
 
Original_String = InputBox("Pls enter the original string: ")
 
'Find the revised length of the string
 
Length = Len(Original_String)
 
'Set up the reversed string
Reversed_String = ""
 
'Progress through the string on a character by character basis
'Starting at the last character and going towards the first character
 
For Pos = Length To 1 Step -1
 
    Next_Char = Mid(Original_String, Pos, 1)
    Reversed_String = Reversed_String & Next_Char
Next Pos
 
MsgBox "The reversed string is " & Reversed_String
 
End Sub