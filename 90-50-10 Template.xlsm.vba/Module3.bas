Sub Step3()
Dim mySheet As Worksheet, eRange, dRange As Range
For Each mySheet In Worksheets

'Booking Tab for each team
If mySheet.Name = "NE Asia Team" Or mySheet.Name = "ROW Team" Or mySheet.Name = "Tradeshow Team" Then

Set eRange = mySheet.Range("A3").SpecialCells(xlCellTypeLastCell)
Set dRange = mySheet.Range("A3", eRange)
With mySheet.Sort
.SortFields.Clear
.SortFields.Add Key:=Range("A3"), SortOn:=xlSortOnValues, Order:=xlAscending
.SortFields.Add Key:=Range("C3"), SortOn:=xlSortOnValues, Order:=xlDescending
.SortFields.Add Key:=Range("B3"), SortOn:=xlSortOnValues, Order:=xlDescending
.SetRange dRange
.Header = xlYes
.Apply
End With

mySheet.Columns("Q:U").Hidden = True

'Room Tab for each team
ElseIf mySheet.Name = "NE Asia RN Block" Or mySheet.Name = "ROW RN Block" Or mySheet.Name = "Tradeshow RN Block" Then

Set eRange = mySheet.Range("A3").SpecialCells(xlCellTypeLastCell)
Set dRange = mySheet.Range("A3", eRange)
With mySheet.Sort
.SortFields.Clear
.SortFields.Add Key:=Range("D3"), SortOn:=xlSortOnValues, Order:=xlAscending
.SortFields.Add Key:=Range("A3"), SortOn:=xlSortOnValues, Order:=xlAscending
.SortFields.Add Key:=Range("F3"), SortOn:=xlSortOnValues, Order:=xlAscending
.SetRange dRange
.Header = xlYes
.Apply
End With

mySheet.Columns("J:L").Hidden = True

End If
Next
End Sub

