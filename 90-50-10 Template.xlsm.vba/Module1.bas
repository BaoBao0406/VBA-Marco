Sub Step1()
Dim x, y As Variant, aRange, bRange, xRange, yRange As Range, d1, d2 As String, mySheet, mySheet1 As Worksheet

'Date in CoverPage
d1 = Date
d2 = Format(d1, "dd mmmm yyyy")
Worksheets(1).Select
Range("A3").Value = "Date: " & d2

'Remove the MOP dollar sign
x = "MOP"
y = " "
Worksheets("Room Block").Select
Range("I1").EntireColumn.Replace what:=x, Replacement:=y

Worksheets("Raw Data").Select
Range("L1:O1").EntireColumn.Replace what:=x, Replacement:=y

'Revenue
Range("M4").Formula = "=SUMIF('Room Block'!$C:$C, 'Raw Data'!$I4, 'Room Block'!$I:$I)"
Range(Cells(4, 13), Range("M4").End(xlDown)).FormulaR1C1 = Range("M4").FormulaR1C1
Range(Cells(4, 13), Range("M4").End(xlDown)).FormulaR1C1 = Range(Cells(4, 13), Range("M4").End(xlDown)).Value

'Room Night
Range("L4").Formula = "=SUMIF('Room Block'!$C:$C, 'Raw Data'!$I4, 'Room Block'!$H:$H)"
Range(Cells(4, 12), Range("L4").End(xlDown)).FormulaR1C1 = Range("L4").FormulaR1C1
Range(Cells(4, 12), Range("L4").End(xlDown)).FormulaR1C1 = Range(Cells(4, 12), Range("L4").End(xlDown)).Value

'Unhide all columns in tab
For Each mySheet In Sheets
If mySheet.Name = "NE Asia Team" Or mySheet.Name = "ROW Team" Or mySheet.Name = "Tradeshow Team" Or mySheet.Name = "NE Asia RN Block" _
Or mySheet.Name = "Leisure RN Block" Or mySheet.Name = "ROW RN Block" Then
mySheet.Columns.EntireColumn.Hidden = False
End If
Next

'Clear History Data for Team tab
For Each mySheet1 In Sheets
If mySheet1.Name = "NE Asia Team" Or mySheet1.Name = "ROW Team" Or mySheet1.Name = "Tradeshow Team" Or mySheet1.Name = "NE Asia RN Block" _
Or mySheet1.Name = "Tradeshow RN Block" Or mySheet1.Name = "ROW RN Block" Then

Set bRange = mySheet1.Range("A4").SpecialCells(xlCellTypeLastCell)
Set aRange = mySheet1.Range("A4", bRange)
aRange.Clear
End If
Next


End Sub



