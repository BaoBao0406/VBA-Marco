Sub ClearVM()
Dim aRange, bRange As Range, k As Integer

On Error GoTo errHandler

Worksheets("VM Room").Select

k = 7
Do Until Worksheets("VM Room").Cells(k, 1).Value = "***"
k = k + 1
Loop
k = k - 2

Set aRange = Worksheets("VM Room").Range(Cells(7, 1), Cells(k, 3))
Set bRange = Worksheets("VM Room").Range(Cells(7, 5), Cells(k, 12))

aRange.ClearContents
bRange.ClearContents

errHandler:
    Resume Next

End Sub

Sub ClearPA()
Dim aRange, bRange As Range, k As Integer

On Error GoTo errHandler

Worksheets("PA Room").Select

k = 7
Do Until Worksheets("PA Room").Cells(k, 1).Value = "***"
k = k + 1
Loop
k = k - 2

Set aRange = Worksheets("PA Room").Range(Cells(7, 1), Cells(k, 3))
Set bRange = Worksheets("PA Room").Range(Cells(7, 5), Cells(k, 12))

aRange.ClearContents
bRange.ClearContents

errHandler:
    Resume Next

End Sub

Sub ClearCM()
Dim aRange, bRange As Range, k As Integer

On Error GoTo errHandler

Worksheets("CM Room").Select

k = 7
Do Until Worksheets("CM Room").Cells(k, 1).Value = "***"
k = k + 1
Loop
k = k - 2

Set aRange = Worksheets("CM Room").Range(Cells(7, 1), Cells(k, 3))
Set bRange = Worksheets("CM Room").Range(Cells(7, 5), Cells(k, 12))

aRange.ClearContents
bRange.ClearContents

errHandler:
    Resume Next

End Sub

Sub ClearHI()
Dim aRange, bRange As Range, k As Integer

On Error GoTo errHandler

Worksheets("HI Room").Select

k = 7
Do Until Worksheets("HI Room").Cells(k, 1).Value = "***"
k = k + 1
Loop
k = k - 2

Set aRange = Worksheets("HI Room").Range(Cells(7, 1), Cells(k, 3))
Set bRange = Worksheets("HI Room").Range(Cells(7, 5), Cells(k, 12))

aRange.ClearContents
bRange.ClearContents

errHandler:
    Resume Next

End Sub

Sub ClearEvents()
Dim aRange, bRange, cRange, dRange, eRange, fRange, gRange, hRange, iRange As Range

On Error GoTo errHandler

Worksheets("Events").Select

'F&B min
Set aRange = Worksheets("Events").Range("A4").End(xlDown)
Set bRange = aRange.Offset(0, 4)
Set cRange = Range("A4", bRange)

cRange.ClearContents

'Rental
Set dRange = Worksheets("Events").Range("I4").End(xlDown)
Set eRange = dRange.Offset(0, 3)
Set fRange = Range("I4", eRange)

fRange.ClearContents

'Mtg Pkg
Set gRange = Worksheets("Events").Range("O4").End(xlDown)
Set hRange = gRange.Offset(0, 3)
Set iRange = Range("O4", hRange)

iRange.ClearContents

errHandler:
    Resume Next

End Sub

Sub ClearCommentPad()
Dim aRange, bRange As Range

Worksheets("CommentPad").Select

Set aRange = Worksheets("CommentPad").Range("A1").SpecialCells(xlCellTypeLastCell)
Set bRange = Range("A1", aRange)

bRange.Clear


End Sub

Sub ClearEventTable()
Dim aRange, bRange, cRange As Range

On Error GoTo errHandler

Worksheets("Event Table").Select
Set aRange = Worksheets("Event Table").Range("A2").End(xlDown)
Set bRange = aRange.Offset(0, 14)
Set cRange = Range("A2", bRange)

cRange.ClearContents

errHandler:
    Resume Next

End Sub

Sub ClearBKInfo()
Dim aRange As Range

On Error GoTo errHandler

Worksheets("BK Info").Select
Set aRange = Range("B2:B16")

aRange.ClearContents

cRange.ClearContents

errHandler:
    Resume Next

End Sub

Sub ClearRoomBL()

Dim aRange, bRange As Range

Worksheets("Rm Table").Select

Set aRange = Worksheets("Rm Table").Range("A1").SpecialCells(xlCellTypeLastCell)
Set bRange = Range("A2", aRange)

bRange.Clear

End Sub