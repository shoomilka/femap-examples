' this function reflects entity

Function refl(id) As Long
	Dim c1(3) As Double
	c1(0) = 0
    c1(1) = 0
    c1(2) = 0
    Dim c2(3) As Double
    c2(0) = 10000000000000  ' THIS IS VERY VERY VERY NECESSARY
                            ' FEMAP will not work correctly with 1 or other small number
    c2(1) = 0
    c2(2) = 0
    ' 5 is entity type ... see to FEMAP help
	App.feGenerateReflect(5, -id, c1, c2, 0, False)
	Dim tempe As Variant
	tempe = getSurf()
	refl = tempe(1)
End Function