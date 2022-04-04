Sub removeChar()
Dim Rng As Range
Dim rc As String
rc = InputBox("Character(s) to Replace", "Enter Value")
For Each Rng In Selection
Selection.Replace What:=rc, Replacement:=""
Next
End Sub