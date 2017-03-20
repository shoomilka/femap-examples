
Sub changeWP()
    ' I used programming tool
	App.feFileProgramRun(False,True, True, CurDir()+"\p197-wp.pro")
    ' This message is for good feelings and also this is manual sleep function
    ' So you can replace it with sleep(), but ...standart FEMAP dlls doesn't allow
    ' sleep func by default(((
	MsgBox("smile :-)")
End Sub