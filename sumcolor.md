

<pre>
Function SumByColor(CellColor As Range, SumRange As Range)
	Application.Volatile
	Dim ICol As Integer
	Dim TCell As Range
	ICol = CellColor.Interior.ColorIndex
	For Each TCell In SumRange
		If ICol = TCell.Interior.ColorIndex Then
			SumByColor = SumByColor + TCell.Value
		End If
		Next TCell
	End Function
