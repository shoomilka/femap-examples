Sub addProps()
	' Material creation
	Dim mt As femap.Matl
    Set mt = App.feMatl

	Dim Mat_Id As Long
	Mat_Id =mt.NextEmptyID

	mt.title = "Matl 1"
	mt.Ex = 2.1e11
	mt.Nuxy = 0.3
	mt.Density = 7800
	mt.Put(Mat_Id)
    ' Material creation end

    ' variables for properties
    Dim pr As femap.Prop
	Set pr = App.feProp

	Dim Prop_Id As Long
    Prop_Id =pr.NextEmptyID

    Dim countElem As Long
    ' variables for properties end

	pr.title = "Plate "+ wb.Sheets("data").Cells(3, 2).Value
    pr.type = FET_L_PLATE
    pr.matlID = Mat_Id
	pr.pval(0) = wb.Sheets("data").Cells(3, 3).Value
	pr.Put(Prop_Id)
	wb.Sheets("data").Cells(3, 5).Value = Prop_Id
	countElem = wb.Sheets("data").Cells(3, 7).Value

	Dim I As Integer
	Dim iTem As Long

	For I = 8 To 13
		iTem = wb.Sheets("index").Cells(I, 20).Value
		App.feMeshAttrSurface(-iTem, Prop_Id, 0)
	Next

	Prop_Id =pr.NextEmptyID

	pr.title = "Plate "+ wb.Sheets("data").Cells(8, 2).Value
    pr.type = FET_L_PLATE
    pr.matlID = Mat_Id
	pr.pval(0) = wb.Sheets("data").Cells(8, 3).Value
	pr.Put(Prop_Id)
	wb.Sheets("data").Cells(8, 5).Value = Prop_Id
	countElem = wb.Sheets("data").Cells(8, 7).Value

	For I = 3 To 5
		iTem = wb.Sheets("index").Cells(I, 26).Value
		App.feMeshAttrSurface(-iTem, Prop_Id, 0)
	Next

	Prop_Id =pr.NextEmptyID

	pr.title = "Plate "+ wb.Sheets("data").Cells(4, 2).Value
    pr.type = FET_L_PLATE
    pr.matlID = Mat_Id
	pr.pval(0) = wb.Sheets("data").Cells(4, 3).Value
	pr.Put(Prop_Id)
	wb.Sheets("data").Cells(4, 5).Value = Prop_Id
	countElem = wb.Sheets("data").Cells(4, 7).Value

	For I = 8 To 10
		iTem = wb.Sheets("index").Cells(I, 11).Value
		App.feMeshAttrSurface(-iTem, Prop_Id, 0)

		iTem = wb.Sheets("index").Cells(I, 12).Value
		App.feMeshAttrSurface(-iTem, Prop_Id, 0)
	Next

	Prop_Id =pr.NextEmptyID

	pr.title = "Plate "+ wb.Sheets("data").Cells(5, 2).Value
    pr.type = FET_L_PLATE
    pr.matlID = Mat_Id
	pr.pval(0) = wb.Sheets("data").Cells(5, 3).Value
	pr.Put(Prop_Id)
	wb.Sheets("data").Cells(5, 5).Value = Prop_Id
	countElem = wb.Sheets("data").Cells(5, 7).Value

	For I = 8 To 10
		iTem = wb.Sheets("index").Cells(I, 17).Value
		App.feMeshAttrSurface(-iTem, Prop_Id, 0)

		iTem = wb.Sheets("index").Cells(I, 18).Value
		App.feMeshAttrSurface(-iTem, Prop_Id, 0)
	Next

	Prop_Id =pr.NextEmptyID

	pr.title = "Plate "+ wb.Sheets("data").Cells(9, 2).Value
    pr.type = FET_L_PLATE
    pr.matlID = Mat_Id
	pr.pval(0) = wb.Sheets("data").Cells(9, 3).Value
	pr.Put(Prop_Id)
	wb.Sheets("data").Cells(9, 5).Value = Prop_Id
	countElem = wb.Sheets("data").Cells(9, 7).Value

	For I = 28 To 30
		iTem = wb.Sheets("index").Cells(I, 5).Value
		App.feMeshAttrSurface(-iTem, Prop_Id, 0)


		iTem = wb.Sheets("index").Cells(I, 6).Value
		App.feMeshAttrSurface(-iTem, Prop_Id, 0)

	Next

	Prop_Id =pr.NextEmptyID

	pr.title = "Plate "+ wb.Sheets("data").Cells(10, 2).Value
    pr.type = FET_L_PLATE
    pr.matlID = Mat_Id
	pr.pval(0) = wb.Sheets("data").Cells(10, 3).Value
	pr.Put(Prop_Id)
	wb.Sheets("data").Cells(10, 5).Value = Prop_Id
	countElem = wb.Sheets("data").Cells(10, 7).Value

	For I = 49 To 60
		iTem = wb.Sheets("index").Cells(I, 12).Value
		App.feMeshAttrSurface(-iTem, Prop_Id, 0)
	Next

	Dim sur As femap.Surface
	Set sur = App.feSurface
	Dim cg As Variant

	sur.Reset

	While sur.Next
		If sur.attrPID = False Then
			sur.cg(cg)
			If Abs(cg(2) - wb.Sheets("index").Cells(18, 6).Value) < 100 Then
				countElem = wb.Sheets("data").Cells(8, 7).Value
				App.feMeshAttrSurface(-sur.ID, wb.Sheets("index").Cells(8, 5).Value, 0)
			ElseIf Abs(cg(2) - wb.Sheets("index").Cells(3, 23).Value) < 100 Then
				countElem = wb.Sheets("data").Cells(3, 7).Value
				App.feMeshAttrSurface(-sur.ID, wb.Sheets("index").Cells(3, 5).Value, 0)
			ElseIf Abs(cg(2)) < 1000 Then
				countElem = wb.Sheets("data").Cells(10, 7).Value
				App.feMeshAttrSurface(-sur.ID, wb.Sheets("index").Cells(10, 5).Value, 0)
			Else
				countElem = wb.Sheets("data").Cells(8, 7).Value
				App.feMeshAttrSurface(-sur.ID, wb.Sheets("index").Cells(8, 5).Value, 0)
			End If
		End If
	Wend

	countElem = wb.Sheets("data").Cells(8, 7).Value
	App.feMeshAttrSurface(-124, wb.Sheets("index").Cells(8, 5).Value, 0)
	App.feMeshAttrSurface(-125, wb.Sheets("index").Cells(8, 5).Value, 0)
	App.feMeshAttrSurface(-126, wb.Sheets("index").Cells(8, 5).Value, 0)
End Sub