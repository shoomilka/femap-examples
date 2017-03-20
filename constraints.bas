Sub setConstr()
	Dim bcs As femap.BCSet
	Set bcs = App.feBCSet

	Dim bcs_id As Long
	bcs_id = bcs.NextEmptyID

	bcs.title = "Constr_1"
	bcs.Put(bcs_id)
	bcs.Active = bcs_id

	Dim bcn As femap.BCNode
	Set bcn = App.feBCNode

	Dim bcn_id As Long
	bcn_id = bcn.NextEmptyID

	Dim ns As femap.Node
	Set ns = App.feNode

	ns.Reset

	bcn.setID = bcs_id

	While ns.Next()
		If ns.z < 0.010 Then
			rc = bcn.Add(-ns.ID,True,True,True,True,True,True)
		End If
	Wend

	bcn.Put(bcn_id)
End Sub