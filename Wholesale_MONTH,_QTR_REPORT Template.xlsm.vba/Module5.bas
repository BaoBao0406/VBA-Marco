Sub Step5ChinaDataRNRev()
Dim d1, d2 As Integer, aRange, bRange, cRange, dRange, vRange, xRange, yRange, zRange, wRange As Range, a, b, c As String

'Get the Month value
d1 = DateAdd("m", -1, Now)
d2 = Format(d1, "m")


'Amend the Manager Name in here if needed
a = "Jacky"
b = "Janet"
c = "Sidney"

'Jacky China Account RN Rev
Worksheets("China figure (RN Rev)").Select

Set wRange = Range("N6:N12")

For Each xRange In wRange
If xRange.Value = a Then

Set yRange = xRange.Offset(0, 1)
Set zRange = xRange.Offset(0, 4)
Set vRange = Range(yRange, zRange)

vRange.Copy

End If
Next

Worksheets("RN Rev Raw data").Select

Set aRange = Range("A2:A30")

For Each bRange In aRange
If bRange.Value = a Then

Set cRange = bRange.Offset(0, (d2 * 4) + 1)

cRange.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Application.CutCopyMode = False

End If
Next

Set aRange = Nothing
Set bRange = Nothing
Set cRange = Nothing
Set vRange = Nothing
Set wRange = Nothing
Set xRange = Nothing
Set yRange = Nothing
Set zRange = Nothing

'Janet China Account RN Rev
Worksheets("China figure (RN Rev)").Select

Set wRange = Range("N6:N12")

For Each xRange In wRange
If xRange.Value = b Then

Set yRange = xRange.Offset(0, 1)
Set zRange = xRange.Offset(0, 4)
Set vRange = Range(yRange, zRange)

vRange.Copy

End If
Next

Worksheets("RN Rev Raw data").Select

Set aRange = Range("A2:A30")

For Each bRange In aRange
If bRange.Value = b Then

Set cRange = bRange.Offset(0, (d2 * 4) + 1)

cRange.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Application.CutCopyMode = False

End If
Next

Set aRange = Nothing
Set bRange = Nothing
Set cRange = Nothing
Set vRange = Nothing
Set wRange = Nothing
Set xRange = Nothing
Set yRange = Nothing
Set zRange = Nothing

'Sidney China Account RN Rev
Worksheets("China figure (RN Rev)").Select

Set wRange = Range("N6:N12")

For Each xRange In wRange
If xRange.Value = c Then

Set yRange = xRange.Offset(0, 1)
Set zRange = xRange.Offset(0, 4)
Set vRange = Range(yRange, zRange)

vRange.Copy

End If
Next

Worksheets("RN Rev Raw data").Select

Set aRange = Range("A2:A30")

For Each bRange In aRange
If bRange.Value = c Then

Set cRange = bRange.Offset(0, (d2 * 4) + 1)

cRange.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Application.CutCopyMode = False

End If
Next

Set aRange = Nothing
Set bRange = Nothing
Set cRange = Nothing
Set vRange = Nothing
Set wRange = Nothing
Set xRange = Nothing
Set yRange = Nothing
Set zRange = Nothing

End Sub