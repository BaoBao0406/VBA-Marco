Sub Step1()
Dim d1, d2 As Integer, VMRH, CMCC, HICC, PARIS, aRange, bRange, cRange, dRange, eRange As Range, a, b, x, y As String

'Get the Month value
d1 = DateAdd("m", -1, Now)
d2 = Format(d1, "m")


'Replace CMCC and HICC Country name mismatch
x = "Macau"
y = "Macao"
Worksheets("CMCC").Select
Range("C1").EntireColumn.Replace What:=x, Replacement:=y

Worksheets("HICC").Select
Range("C1").EntireColumn.Replace What:=x, Replacement:=y

'Replace Country Name Using Other Countries instead of Australia
a = "Australia"
b = "Other Countries"

Worksheets("VMRH").Select
Range("C1").EntireColumn.Replace What:=a, Replacement:=b

Worksheets("CMCC").Select
Range("C1").EntireColumn.Replace What:=a, Replacement:=b

Worksheets("HICC").Select
Range("C1").EntireColumn.Replace What:=a, Replacement:=b

Worksheets("PARIS").Select
Range("C1").EntireColumn.Replace What:=a, Replacement:=b

'Copy Equation to RN Tab
Worksheets("RN Raw data").Select

Set VMRH = Cells(2, (d2 * 4) + 2)
Set CMCC = Cells(2, (d2 * 4) + 3)
Set HICC = Cells(2, (d2 * 4) + 4)
Set PARIS = Cells(2, (d2 * 4) + 5)

Set aRange = Range("B1").End(xlDown)

Set bRange = aRange.Offset(0, d2 * 4)
Set cRange = aRange.Offset(0, (d2 * 4) + 1)
Set dRange = aRange.Offset(0, (d2 * 4) + 2)
Set eRange = aRange.Offset(0, (d2 * 4) + 3)

VMRH.Formula = "=SUMIF(VMRH!C:C, 'RN Raw data'!A2, VMRH!N:N)"
Range(VMRH, bRange).FormulaR1C1 = VMRH.FormulaR1C1
Range(VMRH, bRange).FormulaR1C1 = Range(VMRH, bRange).Value

CMCC.Formula = "=SUMIF(CMCC!C:C, 'RN Raw data'!A2, CMCC!N:N)"
Range(CMCC, cRange).FormulaR1C1 = CMCC.FormulaR1C1
Range(CMCC, cRange).FormulaR1C1 = Range(CMCC, cRange).Value

HICC.Formula = "=SUMIF(HICC!C:C, 'RN Raw data'!A2, HICC!N:N)"
Range(HICC, dRange).FormulaR1C1 = HICC.FormulaR1C1
Range(HICC, dRange).FormulaR1C1 = Range(HICC, dRange).Value

PARIS.Formula = "=SUMIF(PARIS!C:C, 'RN Raw data'!A2, PARIS!N:N)"
Range(PARIS, eRange).FormulaR1C1 = PARIS.FormulaR1C1
Range(PARIS, eRange).FormulaR1C1 = Range(PARIS, eRange).Value

Set VMRH = Nothing
Set CMCC = Nothing
Set HICC = Nothing
Set PARIS = Nothing
Set aRange = Nothing
Set bRange = Nothing
Set cRange = Nothing
Set dRange = Nothing
Set eRange = Nothing

'Copy Equation to RN Rev Tab
Worksheets("RN Rev Raw data").Select

Set VMRH = Cells(2, (d2 * 4) + 2)
Set CMCC = Cells(2, (d2 * 4) + 3)
Set HICC = Cells(2, (d2 * 4) + 4)
Set PARIS = Cells(2, (d2 * 4) + 5)

Set aRange = Range("B1").End(xlDown)

Set bRange = aRange.Offset(0, d2 * 4)
Set cRange = aRange.Offset(0, (d2 * 4) + 1)
Set dRange = aRange.Offset(0, (d2 * 4) + 2)
Set eRange = aRange.Offset(0, (d2 * 4) + 3)

VMRH.Formula = "=SUMIF(VMRH!C:C, 'RN Rev Raw data'!A2, VMRH!P:P)"
Range(VMRH, bRange).FormulaR1C1 = VMRH.FormulaR1C1
Range(VMRH, bRange).FormulaR1C1 = Range(VMRH, bRange).Value

CMCC.Formula = "=SUMIF(CMCC!C:C, 'RN Rev Raw data'!A2, CMCC!P:P)"
Range(CMCC, cRange).FormulaR1C1 = CMCC.FormulaR1C1
Range(CMCC, cRange).FormulaR1C1 = Range(CMCC, cRange).Value

HICC.Formula = "=SUMIF(HICC!C:C, 'RN Rev Raw data'!A2, HICC!P:P)"
Range(HICC, dRange).FormulaR1C1 = HICC.FormulaR1C1
Range(HICC, dRange).FormulaR1C1 = Range(HICC, dRange).Value

PARIS.Formula = "=SUMIF(PARIS!C:C, 'RN Rev Raw data'!A2, PARIS!P:P)"
Range(PARIS, eRange).FormulaR1C1 = PARIS.FormulaR1C1
Range(PARIS, eRange).FormulaR1C1 = Range(PARIS, eRange).Value

Set VMRH = Nothing
Set CMCC = Nothing
Set HICC = Nothing
Set PARIS = Nothing
Set aRange = Nothing
Set bRange = Nothing
Set cRange = Nothing
Set dRange = Nothing
Set eRange = Nothing

End Sub
