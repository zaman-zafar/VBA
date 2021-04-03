<pre>

Private Sub Cancel_Click()
	Unload UserForm1
	
End Sub
</pre>

<a href="red">**Private Sub Ok_Click()**</a>
<pre>

	Dim ran As Range
	Dim cel As Range
	
	Set ran = Selection
	Application.ScreenUpdating = False
	
	On Error GoTo ErrHandler
	
	If UpperCase Then
		For Each cel In ran
			cel.Value = UCase(cel.Value)
			Next cel
		End If
		
		If LowerCase Then
			For Each cel In ran
				cel.Value = LCase(cel.Value)
				Next cel
			End If
			
			
			If ProperCase Then
				For Each cel In ran
					cel.Value = WorksheetFunction.Proper(cel.Value)
					Next cel
				End If
				
				Application.ScreenUpdating = True
				
				Unload UserForm1
				
				Exit Sub
				
				
				ErrHandler:
				
				MsgBox "Please select 1 option"
				
			End Sub
		</pre>	
			
<a href="red">**Sub ChangeCase()**</a>
<pre>			

If TypeName(Selection) = "Range" Then

UserForm1.Show
End If

End Sub