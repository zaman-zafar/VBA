<a href="red">**Report Automation**</a>
<pre>

' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+a

Dim StartPoint As Range
Dim DataRange As Range
Dim NewRange As String
Dim LastCol As Long
Dim lastRow As Long
Dim Data_Sheet As Worksheet



    Columns("E:E").EntireColumn.Select
    Selection.Cut
    Columns("A:A").EntireColumn.Select
    Selection.Insert Shift:=xlToRight
    ActiveCell.Offset(0, 10).Range("A1").Select
    ActiveCell.FormulaR1C1 = "1"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "2"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.AutoFilter
    
    Worksheets("Product Receipt Vs Delivery").Range("k2").Value = "=VLOOKUP(A2,$A$2:$D$50139,2,FALSE)"
    Worksheets("Product Receipt Vs Delivery").Range("l2").Value = "=VLOOKUP(A2,$A$2:$D$50139,3,FALSE)"
    Worksheets("Product Receipt Vs Delivery").Range("m2").Value = "=VLOOKUP(A2,$A$2:$D$50139,4,FALSE)"
    
    
    
     Range("K2:M2").Select
    Selection.AutoFill Destination:=Range("K2:M" & Range("A" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).Select

    ActiveWorkbook.Worksheets("Product Receipt Vs Delivery").AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Product Receipt Vs Delivery").AutoFilter.Sort. _
        SortFields.Add Key:=Range("B1:B" & Range("A" & Rows.Count).End(xlUp).Row), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Product Receipt Vs Delivery").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Product Receipt Vs Delivery").AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Product Receipt Vs Delivery").AutoFilter.Sort. _
        SortFields.Add Key:=Range("C1:C" & Range("A" & Rows.Count).End(xlUp).Row), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Product Receipt Vs Delivery").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("K2:L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=6
    Range("K2:L" & Range("A" & Rows.Count).End(xlUp).Row).Copy
    Range("K2:L" & Range("A" & Rows.Count).End(xlUp).Row).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Product Receipt Vs Delivery").AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Product Receipt Vs Delivery").AutoFilter.Sort. _
        SortFields.Add Key:=Range("D1:D" & Range("A" & Rows.Count).End(xlUp).Row), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Product Receipt Vs Delivery").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("M2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=6
    Range("M2:M" & Range("A" & Rows.Count).End(xlUp).Row).Copy
    Range("M2:M" & Range("A" & Rows.Count).End(xlUp).Row).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False


    Set Data_Sheet = ActiveWorkbook.Worksheets("Product Receipt Vs Delivery")
    Data_Sheet.Activate
    Set StartPoint = Data_Sheet.Range("A1")
    LastCol = StartPoint.End(xlToRight).Column
    DownCell = StartPoint.End(xlDown).Row
    Set DataRange = Data_Sheet.Range(StartPoint, Cells(DownCell, LastCol))
    NewRange = Data_Sheet.Name & "!" & DataRange.Address(ReferenceStyle:=xlR1C1)



    Range("A1").Select
    ActiveWindow.SmallScroll Down:=-15
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        NewRange, Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="Sheet1!R3C1", TableName:="PivotTable2", DefaultVersion _
        :=xlPivotTableVersion14
    Sheets("Sheet1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("SALESID")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("1")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("2")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("3")
        .Orientation = xlRowField
        .Position = 4
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("company1"), "Sum of company1", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("company2"), "Sum of company2", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("difference"), "Sum of difference", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("company21"), "Sum of company21", xlSum
    Range("C7").Select
    ActiveSheet.PivotTables("PivotTable2").RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("PivotTable2").PivotFields("SALESID").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("ITEMBUYERGROUPID"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("STORENUMBER").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Vendor").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("ShippingDate"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("ITEMID").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("company1").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("company2").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("difference").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("company21").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("1").Subtotals = Array(False _
        , False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("2").Subtotals = Array(False _
        , False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("3").Subtotals = Array(False _
        , False, False, False, False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("ShippingDate")
        .Orientation = xlRowField
        .Position = 5
    End With
 
    Range("E4").Select
   
    
    
End Sub

