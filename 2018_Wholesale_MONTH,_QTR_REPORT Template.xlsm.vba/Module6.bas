Sub Step6()
Dim xTable As PivotTable, d1, d2 As Integer

For Each xTable In Worksheets("Wholesale Actual").PivotTables
xTable.RefreshTable

Next

Worksheets("Summary").Select

'Get the Month value
d1 = DateAdd("m", -1, Now)
d2 = Format(d1, "m")

End Sub
