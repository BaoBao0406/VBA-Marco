Sub Step2()
Dim d1, i, j As Integer, aRange, bRange, cRange, dRange, eRange, fRange, gRange, hRange, xRange, yRange As Range, d2 As String

'Get the Month value
d1 = DateAdd("m", -1, Now)
d2 = Format(d1, "m")


'Unhide Worksheets
Worksheets("RN Goal (Hide)").Visible = True
Worksheets("RN Rev Goal (Hide)").Visible = True
Worksheets("China AC (hidden)").Visible = True

'Use China AC Hidden Table for Variants
Worksheets("China AC (hidden)").Select
Set aRange = Range("N2:N13")

For Each bRange In aRange
If d2 = bRange.Value Then

i = bRange.Offset(0, 1).Value
j = bRange.Offset(0, 2).Value

End If
Next

'Display RN Goal by Copying RN Goal
Worksheets("RN Goal").Select
Set aRange = Range("B4").End(xlDown)
Set bRange = aRange.Offset(0, j)
Set cRange = Range("B4").Offset(1, i)
Set dRange = Range(cRange, bRange)

dRange.Copy

Worksheets("RN Goal (Hide)").Select
Set eRange = Range("B4").End(xlDown)
Set fRange = eRange.Offset(0, j)
Set gRange = Range("B4").Offset(1, i)
Set hRange = Range(gRange, fRange)

hRange.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Application.CutCopyMode = False

Set aRange = Nothing
Set bRange = Nothing
Set cRange = Nothing
Set dRange = Nothing
Set eRange = Nothing
Set fRange = Nothing
Set gRange = Nothing
Set hRange = Nothing


'Display RN Rev Goal by Copying RN Rev Goal
Worksheets("RN Rev Goal").Select
Set aRange = Range("B4").End(xlDown)
Set bRange = aRange.Offset(0, j)
Set cRange = Range("B4").Offset(1, i)
Set dRange = Range(cRange, bRange)

dRange.Copy

Worksheets("RN Rev Goal (Hide)").Select
Set eRange = Range("B4").End(xlDown)
Set fRange = eRange.Offset(0, j)
Set gRange = Range("B4").Offset(1, i)
Set hRange = Range(gRange, fRange)

hRange.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Application.CutCopyMode = False

Set aRange = Nothing
Set bRange = Nothing
Set cRange = Nothing
Set dRange = Nothing
Set eRange = Nothing
Set fRange = Nothing
Set gRange = Nothing
Set hRange = Nothing

'Clear China RN figure
Worksheets("China figure (RN)").Select

Set xRange = Range("A4").End(xlDown)
Set yRange = Range("A4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

Set xRange = Range("C4").End(xlDown)
Set yRange = Range("C4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

Set xRange = Range("D4").End(xlDown)
Set yRange = Range("D4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

Set xRange = Range("F4").End(xlDown)
Set yRange = Range("F4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

Set xRange = Range("G4").End(xlDown)
Set yRange = Range("G4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

Set xRange = Range("I4").End(xlDown)
Set yRange = Range("I4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

Set xRange = Range("J4").End(xlDown)
Set yRange = Range("J4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

Set xRange = Range("L4").End(xlDown)
Set yRange = Range("L4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

'Clear China RN Rev figure
Worksheets("China figure (RN Rev)").Select

Set xRange = Range("A4").End(xlDown)
Set yRange = Range("A4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

Set xRange = Range("C4").End(xlDown)
Set yRange = Range("C4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

Set xRange = Range("D4").End(xlDown)
Set yRange = Range("D4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

Set xRange = Range("F4").End(xlDown)
Set yRange = Range("F4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

Set xRange = Range("G4").End(xlDown)
Set yRange = Range("G4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

Set xRange = Range("I4").End(xlDown)
Set yRange = Range("I4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

Set xRange = Range("J4").End(xlDown)
Set yRange = Range("J4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

Set xRange = Range("L4").End(xlDown)
Set yRange = Range("L4", xRange)
yRange.Clear
Set xRange = Nothing
Set yRange = Nothing

End Sub

