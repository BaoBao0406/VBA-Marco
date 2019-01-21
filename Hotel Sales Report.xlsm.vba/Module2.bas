Sub Step2()
Dim CurrentWbk, OldWbk As Workbook

With Application.FileDialog(msoFileDialogOpen)
.InitialFileName = "I:\10-Sales\01_Sales_Reports\Special Project Rpts\Hotel_Sales_Reports_Analysis\Macau Hotels combined"
If .Show = -1 Then .Execute
End With

Set OldWbk = ActiveWorkbook
Set CurrentWbk = ThisWorkbook

CurrentWbk.Worksheets("Email").Rows.EntireRow.Hidden = False

'First Table (Definite Business and Business Demand)
CurrentWbk.Worksheets("Email").Range("C6:C23").Value = OldWbk.Worksheets("Email").Range("D6:D23").Value

'First Table (Last year ranking)
CurrentWbk.Worksheets("Email").Range("F7").Value = OldWbk.Worksheets("Email").Range("G7").Value
CurrentWbk.Worksheets("Email").Range("F12").Value = OldWbk.Worksheets("Email").Range("G12").Value
CurrentWbk.Worksheets("Email").Range("F17").Value = OldWbk.Worksheets("Email").Range("G17").Value
CurrentWbk.Worksheets("Email").Range("F21").Value = OldWbk.Worksheets("Email").Range("G21").Value

'Second Table (No of Leads and RN Demand)
CurrentWbk.Worksheets("Email").Range("B31:B42").Value = OldWbk.Worksheets("Email").Range("C31:C43").Value
CurrentWbk.Worksheets("Email").Range("E31:E42").Value = OldWbk.Worksheets("Email").Range("F31:F43").Value

'Third Table (Conversion Comparison)
CurrentWbk.Worksheets("Email").Range("B49:B51").Value = OldWbk.Worksheets("Email").Range("C49:C51").Value

'Forth Table (Top Three highest leads)
CurrentWbk.Worksheets("Email").Range("A68:N74").Value = OldWbk.Worksheets("Email").Range("A58:N64").Value

End Sub

