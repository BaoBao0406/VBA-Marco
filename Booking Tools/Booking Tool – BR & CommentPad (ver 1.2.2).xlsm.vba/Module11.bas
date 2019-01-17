Sub Testing1()

Dim ThisWbk, BRWbk As Workbook, aRange, bRange, cRange, dRange, eRange, fRange As Range, i, j As Integer

Set aRange = Worksheets("Event Table").Range("A2").End(xlDown)
Set aRange = aRange.Offset(0, 8)
Set bRange = Worksheets("Event Table").Range("A2")
Set cRange = Worksheets("Event Table").Range(bRange, aRange)

'cRange.Copy

i = cRange.Rows.Count
j = cRange.Columns.Count

Cells(2, 11).Value = i + 22
Cells(2, 12).Value = j

Set dRange = Worksheets("Event Table").Range("B24")
Set eRange = dRange.Offset(i - 1, j - 1)
Set fRange = Worksheets("Event Table").Range(dRange, eRange)
fRange.Select

    'Not Working for Paste Special
    'Set dRange = BRWbk.Worksheets("Meeting Space").Range("B24")
    'Set eRange = dRange.Offset(i - 1, j - 1)
    'Set fRange = BRWbk.Worksheets("Meeting Space").Range(dRange, eRange)
    'dRange.PasteSpecial Paste:=xlPasteValues
    'Application.CutCopyMode = False


End Sub

Sub Testing2()
Dim ThisWbk, BRWbk As Workbook, x, z As Integer
Workbooks.Open Filename:="I:\10-Sales\Personal Folder\Admin & Assistant Team\Patrick Leong\Booking Tools - BR & CommentPad\BR Form_Macao_5.0.xlsm"
Set ThisWbk = ThisWorkbook
Set BRWbk = ActiveWorkbook

Workbooks("BR Form_Macao_5.0").Worksheets("Rooms").Activate

'Application.Run "'BR Form_Macao_5.0.xlsm'!UnhideRowsRequest1"
    x = 15
    If x > 9 Then
        z = 1
        For z = 1 To (x - 9)
            Application.Run "'BR Form_Macao_5.0.xlsm'!UnhideColRequest1"
        Next z
    End If

End Sub

Sub Testing3()
Dim dRange, eRange, fRange As Range
    Set dRange = Range("A34")
    Set eRange = Range("A33").End(xlDown)
    Set eRange = eRange.Offset(0, 1)
    Set fRange = Range(dRange, eRange)
fRange.Select


End Sub