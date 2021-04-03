
<a href="red">**Sub All_Reference()**</a>
<pre>
'   Local Variables
	Dim wks                 As Worksheet
	i = 1
	
	
'   Step 1 : Loop through all worksheets
' 1a : Clear all current hyperlinks
	Worksheets("Info page").Range("c22:c77").ClearContents
	
	For Each wks In Worksheets
		If wks.Visible = True Then
			
			wks.Range("j1").Copy
			Worksheets("info page").Cells(15 + i, 3).Select
			Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
			:=False, Transpose:=False
			Cells(15 + i, 3).Font.Color = -65536
			
			
			i = i + 1
			
		End If
		
		
		Next wks
		
		
		
		
	End Sub
