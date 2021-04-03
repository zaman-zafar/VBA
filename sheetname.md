<pre>

Function SheetName(rCell As Range, Optional UseAsRef As Boolean) As String
	
	Application.Volatile
	
	If UseAsRef = True Then
		
		SheetName = "'" & rCell.Parent.Name & "'!"
		
	Else
		
		SheetName = rCell.Parent.Name
		
	End If
	
End Function
