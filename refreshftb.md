<a href="red">**Sub refreshFTB()**</a>
<pre>
'
' refreshFTB Macro
'

'
Application.ScreenUpdating = False
ActiveSheet.Unprotect ("1")

    ActiveSheet.Range("$A$10:$a$820").AutoFilter Field:=1, Criteria1:="Y"
ActiveSheet.Protect Password:="1", AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True

Application.ScreenUpdating = True
End Sub