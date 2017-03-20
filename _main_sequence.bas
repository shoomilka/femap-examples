Sub Main()

    ' here are your geometry modeling functions and procedures
    ' ...
    ' ...

    ' non-manifold function
    runNM()                                                   ' 1 step !!!! ONLY IN THIS SEQUENCE

    ' scale model. FEMAP works better with big numbers,
    ' so good practice is working in 10000+ meters and then
    ' decrease sizes...
	po(0) = 0
	po(1) = 0
	po(2) = 0
	App.feScale(FT_SOLID, -1, po, 0, 0.001, 0.001, 0.001)

    ' add mesh (as entity),
    ' mesh parameters to elements
	addMesh() ' see create_mesh.bas as example              	2 step !!!!
    ' add properties
	addProps() ' see properties.bas as example					3 step !!!!
    ' draw mesh on each line/surface/volume/...
	meshOnAll() ' see draw_mesh.bas as example					4 step !!!!

    ' add constraints, loads, rigits, etc						5 step !!!!
	setConstr() ' see .bas as example
	RigidCent() ' see .bas as example
	getLoads1() ' see .bas as example
	RigidPit()

	App.feMeshSurfaceByAttributes(tempset.ID)

	clearProfile()

    App.feViewRegenerate(inV)

															  ' 6 step is analysis
															  ' 7 step operations with results
End Sub