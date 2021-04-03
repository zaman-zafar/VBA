<a href="red">**Sub Cheque_Printing()**</a>
<pre>

Dim StartValue As Variant
Dim CellCoun As Integer
Dim RowNumber As Variant
Dim rov As Integer
Dim colum As Integer
Dim ans As Integer
Dim msg As String


'Application.ScreenUpdating = False

Start:                         'this Lable help in starting macro again in case of error






On Error GoTo ErrHandler

rov = Selection.Row            'this gives the number of selected row
colum = Selection.Column       'this gives the number of selected column
CellCoun = 0
StartValue = InputBox("please enter starting Number")
If StartValue = "" Then Exit Sub   'for this to work variable type should be variant
RowNumber = InputBox("please enter Last Number")
If RowNumber = "" Then Exit Sub



StartAgain:


Range("a1").Select
Selection.Value = StartValue + CellCoun

DoEvents

Call RefreshForm

DoEvents
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
        
CellCoun = CellCoun + 1
If CellCoun <> RowNumber - (StartValue - 1) Then GoTo StartAgain 'this will continue untill cellcoun becomes equal to number to rows

Exit Sub


ErrHandler:         'this lable helps in starting the macro again and dealing with error
msg = "Please enter correct value, Want to Try Again"
ans = MsgBox(msg, vbYesNo)
If ans = vbYes Then Resume Start Else Exit Sub
End Sub






