Sub removeDecimals()
Dim lnumber As Double
Dim lResult As Long
Dim rng As Range
For Each rng In Selection
rng.Value = Int(rng)
rng.NumberFormat = "0"
Next rng
End Sub