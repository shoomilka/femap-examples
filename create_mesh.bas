Sub addMesh()
    Dim countElem As Double

	Dim I As Integer
	Dim iTem As Long

	countElem = wb.Sheets("data").Cells(3, 7).Value

	For I = 8 To 13
		iTem = wb.Sheets("index").Cells(I, 20).Value
		App.feMeshSizeSurface(-iTem, True, countElem, 10, 0, 0, 0, 0, False, 0, 0, False)
	Next

	' here you can code operations for entities, which were not meshed by previous commands
    ' I had big model, so there are few elements, cutted by surfaces with parametrical sizes
    ' so sometimes accidental surfaces were occurred...
End Sub