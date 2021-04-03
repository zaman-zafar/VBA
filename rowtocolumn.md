
<a href="red">**Sub Rows_to_Column_Formula()**</a>
<pre>


	Dim colum As Integer
	Dim rov As Integer
	Dim cellNo As Integer
	Dim cou As Integer
	Dim wsk As String
	Dim Msg As String
	Dim ans As String
	Startagain:
	On Error GoTo ErrHandler
	wsk = InputBox("Please enter worksheet Name like Sheet1")
	If wsk = "" Then Exit Sub
	rov = InputBox("please enter row number you want to get value from like row no 1 or 2")
	colum = InputBox("Please enter starting column from where first value comes")
	cou = InputBox("please enter total number of columns")
	For cellNo = 0 To cou
		Sheets("Journals").Select
		ActiveCell.Offset(cellNo, 0).Formula = "=" & "'" & wsk & "'" & "!" _
		& Cells(rov, cellNo + colum).Address
		Next cellNo
		Exit Sub
		ErrHandler:
		Msg = "please enter corrent value, Want to try again"
		ans = MsgBox(Msg, vbYesNo)
		If ans = vbYes Then Resume Startagain Else Exit Sub
	End Sub
	
