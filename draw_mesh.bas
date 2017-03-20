Sub meshOnAll()
	Dim surs As femap.Surface
	Set surs = App.feSurface

	Dim suSet As femap.Set
	Set suSet = App.feSet

	surs.Reset()
	While surs.Next()
		suSet.Add(surs.ID)
	Wend

	App.feMeshSurfaceByAttributes(suSet.ID)
End Sub