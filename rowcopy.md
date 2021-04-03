<a href="red">**Sub Row_Copy()**</a>
<pre>

	Application.ScreenUpdating = False
	ActiveSheet.Unprotect ("786")
	If Selection.Rows.Count <> 1 Then Exit Sub
	If Selection.Areas.Count <> 1 Then Exit Sub
	With Selection.EntireRow
		.Offset(-1).Copy
		.Insert
		On Error Resume Next
		.Offset(-1).SpecialCells(xlCellTypeConstants, 23).ClearContents
		On Error GoTo 0
	End With
	ActiveSheet.Protect Password:="786", AllowFormattingCells:=True, AllowFormattingColumns:=True, _
	AllowFormattingRows:=True
	
	Application.ScreenUpdating = True
End Sub
