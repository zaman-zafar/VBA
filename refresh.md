<a href="red">**Sub refreshETB()**</a>
<pre>

'
' refreshETB Macro
'
	
'
	Application.ScreenUpdating = False
	ActiveSheet.Unprotect ("1")
	
	ActiveSheet.Range("$B$13:$B$966").AutoFilter Field:=1, Criteria1:="Y"
	ActiveSheet.Protect Password:="1"
	Application.ScreenUpdating = True
End Sub
