If Selection.Information(wdWithInTable) = True Then 
 With Selection 
 .InsertColumns 
 .Shading.Texture = wdTexture10Percent 
 End With 
End If