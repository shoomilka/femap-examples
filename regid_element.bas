Sub RigidCent
	Dim rig As femap.Elem
	Set rig = App.feElem

	Dim nod As Variant

	Dim cent As femap.Node
	Set cent = App.feNode
	Dim centID As Long
	centID = cent.NextEmptyID

	cent.x = wb.Sheets("index").Cells(10, 4).Value
	cent.y = wb.Sheets("index").Cells(10, 5).Value
	cent.z = wb.Sheets("index").Cells(10, 6).Value
	cent.Put(centID)
	wb.Sheets("index").Cells(10, 3).Value = centID

	rig.type = 29
	rig.topology = 13
	rig.Node(0) = centID
	rig.release(0, 0) = 1
    rig.release(0, 1) = 1
    rig.release(0, 2) = 1
    rig.release(0, 3) = 1
    rig.release(0, 4) = 1
    rig.release(0, 5) = 1

    Dim I As Integer
    Dim J As Integer
    Dim brac As femap.Surface
    Set brac = App.feSurface
    Dim numN As Long
    Dim noID As Variant

    couRig = 0
	For I = 8 To 10
		brac.Get(wb.Sheets("index").Cells(I, 11).Value)
		brac.Nodes(True, False, numN, noID)
		getNodRig(numN, noID)

		brac.Get(wb.Sheets("index").Cells(I, 12).Value)
		brac.Nodes(True, False, numN, noID)
		getNodRig(numN, noID)

		brac.Get(wb.Sheets("index").Cells(I, 17).Value)
		brac.Nodes(True, False, numN, noID)
		getNodRig(numN, noID)

		brac.Get(wb.Sheets("index").Cells(I, 18).Value)
		brac.Nodes(True, False, numN, noID)
		getNodRig(numN, noID)
	Next

    I = 0
	ReDim nod(0 To couRig+1)
    For I = 0 To couRig-1
    	nod(I) = wb.Sheets("index").Cells(66+I, 2).Value
    Next

	rig.PutNodeList(0, couRig, nod, Null, Null, Null)
	rig.Put(rig.NextEmptyID)
End Sub