Public Function removeFirstC(rng As String, cnt As Long)
removeFirstC = Right(rng, Len(rng) - cnt)
End Function