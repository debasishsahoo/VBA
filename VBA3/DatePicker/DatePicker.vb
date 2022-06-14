Dim datEmpty As Date, datPicked As Date
datPicked = PickDate(BeginDate:=#4/15/2011#, EndDate:=#4/15/2015#)
If datPicked > datEmpty Then
    Debug.Print datPicked
End If