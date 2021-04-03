<a href="red">**Sub newsheet()**</a>
<pre>
'
' Creates a new blank sheet at end
'

Application.ScreenUpdating = False

    Sheet25.Visible = True
    Sheet25.Select
    Cells.Select
    Selection.Copy
    ActiveWindow.ScrollWorkbookTabs Position:=xlLast
    Sheets.Add after:=Sheets(Sheets.Count)
    Cells.Select
    ActiveSheet.Paste
    Range("B2").Select
    Sheet25.Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
End Sub