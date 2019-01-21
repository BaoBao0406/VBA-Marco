Sub Step2()
Dim aName, bName, cName As String, aRange, bRange As Range

aName = "NE ASIA Team"
bName = "Tradeshow Team"
cName = "ROW Team"

Worksheets("Raw Data").Select

'NE ASIA Team Tab
Range("A3").AutoFilter field:=19, Criteria1:=aName
Set bRange = Range("A4").SpecialCells(xlCellTypeLastCell)
Set aRange = Range("A4", bRange)
aRange.Copy Destination:=Worksheets("NE ASIA Team").Range("A4")
Range("A3").AutoFilter

'ROW Team Tab
Range("A3").AutoFilter field:=19, Criteria1:=cName
Set bRange = Range("A4").SpecialCells(xlCellTypeLastCell)
Set aRange = Range("A4", bRange)
aRange.Copy Destination:=Worksheets("ROW Team").Range("A4")
Range("A3").AutoFilter

'Leisure Team Tab
Range("A3").AutoFilter field:=19, Criteria1:=bName
Set bRange = Range("A4").SpecialCells(xlCellTypeLastCell)
Set aRange = Range("A4", bRange)
aRange.Copy Destination:=Worksheets("Tradeshow Team").Range("A4")
Range("A3").AutoFilter

Worksheets("Room Block").Select

'NE Asia RN Block
Range("A3").AutoFilter field:=10, Criteria1:=aName
Set bRange = Range("A4").SpecialCells(xlCellTypeLastCell)
Set aRange = Range("A4", bRange)
aRange.Copy Destination:=Worksheets("NE ASIA RN Block").Range("A4")
Range("A3").AutoFilter

'ROW RN Block
Range("A3").AutoFilter field:=10, Criteria1:=cName
Set bRange = Range("A4").SpecialCells(xlCellTypeLastCell)
Set aRange = Range("A4", bRange)
aRange.Copy Destination:=Worksheets("ROW RN Block").Range("A4")
Range("A3").AutoFilter

'Leisure RN Block
Range("A3").AutoFilter field:=10, Criteria1:=bName
Set bRange = Range("A4").SpecialCells(xlCellTypeLastCell)
Set aRange = Range("A4", bRange)
aRange.Copy Destination:=Worksheets("Tradeshow RN Block").Range("A4")
Range("A3").AutoFilter

End Sub


