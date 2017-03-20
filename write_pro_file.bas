Sub writeSliceProg(idS As Long, idCs As Variant)
	Dim syr As femap.Surface
	Set syr = App.feSurface
	syr.Get(idS)
	Dim vari As Long
	syr.Solid(vari)

	Dim strPath As String
	strPath = CurDir() + "\slicing.pro"

	Dim fso As Object
	Set fso = CreateObject("Scripting.FileSystemObject")
	Dim oFile As Object
	Set oFile = fso.CreateTextFile(strPath)

	Dim tt As String
	tt = "{~1410}"
	oFile.WriteLine tt

	tt = "<@14004><PUSH><OK>"
	oFile.WriteLine tt

	tt = "<@11701>"+CStr(vari)+"<A-M><OK>"
	oFile.WriteLine tt

	Dim I As Integer

	tt = ""
	For I = 1 To idCs(0)
		tt = tt + "<@11701>" + CStr(idCs(I)) + "<A-M>"
	Next
	tt = tt + "<OK>"
	oFile.WriteLine tt

	tt = "<@10022>1000000.<OK>"
	oFile.WriteLine tt

	oFile.Close
	Set fso = Nothing
	Set oFile = Nothing
	MsgBox("hello")
End Sub