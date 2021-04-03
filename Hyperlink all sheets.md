<a href="red">**Sub Hyperlink_all_sheets()**</a>
<pre>
		
'   Local Variables
	Dim wks                 As Worksheet
	Dim rngLinkCell         As Range
	Dim strSubAddress       As String, strDisplayText       As String
	i = 1
'   Step 1 : Loop through all worksheets
' 1a : Clear all current hyperlinks
	Worksheets("Info page").Range("B16:B71").ClearContents
' 1b : Create Linked index list
	For Each wks In ActiveWorkbook.Worksheets
		
		If wks.Visible = True Then
			
			Set rngLinkCell = Worksheets("Info page").Range("B72").End(xlUp)
			If Worksheets("Info page").Range("B16") = "" Then
				Set rngLinkCell = Worksheets("Info page").Range("B16")
			End If
			If rngLinkCell <> "" Then Set rngLinkCell = rngLinkCell.Offset(1, 0)
			strSubAddress = "'" & wks.Name & "'!A1"
			strDisplayText = wks.Name
			Worksheets("Info page").Hyperlinks.Add Anchor:=rngLinkCell, Address:="", SubAddress:=strSubAddress, TextToDisplay:=strDisplayText
			Worksheets("info page").Cells(15 + i, 2).Font.Color = -16776961
			i = i + 1
		End If
		
		Next wks
		
		
		
		
	End Sub
	
