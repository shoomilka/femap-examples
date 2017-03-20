' This procedure creates nodal forces and nodal moments
' wb is public variable: Public wb As Excel.Workbook
' all coordinates, ids and values are imported from Excel file

Sub getLoads1
	Dim LS As femap.LoadSet
	Set LS = App.feLoadSet

	Dim sid As Long
	sid = LS.NextEmptyID

	LS.title = wb.Sheets("data").Cells(2, 16).Value
	LS.Put(sid)
	LS.Active = sid

	Dim LM As femap.LoadMesh
	Set LM = App.feLoadMesh

	Dim midd As Long
	midd = LM.NextEmptyID

	LM.type = FLT_NFORCE

	LM.XOn = True
	LM.YOn = True
	LM.ZOn = True

	LM.x = wb.Sheets("data").Cells(4, 14).Value
	LM.y = wb.Sheets("data").Cells(4, 15).Value
	LM.z = wb.Sheets("data").Cells(4, 16).Value

	LM.setID = LS.ID
	LM.meshID = wb.Sheets("index").Cells(10, 3).Value
	LM.Put(LM.NextEmptyID)

	Dim LMM As femap.LoadMesh
	Set LMM = App.feLoadMesh

	midd = LMM.NextEmptyID

	LMM.type = FLT_NMOMENT

	LMM.XOn = True
	LMM.YOn = True
	LMM.ZOn = True

	LMM.x = wb.Sheets("data").Cells(6, 14).Value
	LMM.y = wb.Sheets("data").Cells(6, 15).Value
	LMM.z = wb.Sheets("data").Cells(6, 16).Value

	LMM.setID = LS.ID
	LMM.meshID = wb.Sheets("index").Cells(10, 3).Value
	LMM.Put(LM.NextEmptyID)

	LS.Get(sid)
	LS.Expand()
	LS.Put(sid)
End Sub