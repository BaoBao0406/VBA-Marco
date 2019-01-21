Sub ConvertRental()
Dim i, j As Integer, RDate, RProp, RName, RRevenue As String

Worksheets("Event Table").Select

i = 2
j = 4
Do While Cells(i, 1).Value > 0
    If Cells(i, 17).Value = "Yes" Or Cells(i, 17).Value = "YES" Or Cells(i, 17).Value = "yes" Then
        RDate = Worksheets("Event Table").Cells(i, 1).Value
        RProp = Worksheets("Event Table").Cells(i, 18).Value
        RName = Worksheets("Event Table").Cells(i, 6).Value
        RRevenue = Format(Worksheets("Event Table").Cells(i, 7).Value, "#,##0.00")

        Worksheets("Events").Cells(j, 9).Value = RDate
        Worksheets("Events").Cells(j, 10).Value = RProp
        Worksheets("Events").Cells(j, 11).Value = RName
        Worksheets("Events").Cells(j, 12).Value = RRevenue
        
        j = j + 1
    
    End If
i = i + 1
Loop


End Sub

Sub ConvertFBmin()
Dim i, j As Integer, FBDate, FBProp, FBName, FBRevenue, FBpax, FBRate As String

Worksheets("Event Table").Select

i = 2
j = 4
Do Until Cells(i, 1).Value = 0
    If Cells(i, 16).Value = "Yes" Or Cells(i, 16).Value = "YES" Or Cells(i, 16).Value = "yes" Then
        FBDate = Worksheets("Event Table").Cells(i, 1).Value
        FBProp = Worksheets("Event Table").Cells(i, 18).Value
        FBName = Worksheets("Event Table").Cells(i, 6).Value
        FBpax = Worksheets("Event Table").Cells(i, 8).Value
        FBRate = Format(Worksheets("Event Table").Cells(i, 20).Value, "#,##0.00")
        
        'FBRevenue = Format(Worksheets("Event Table").Cells(i, 21).Value, "#,##0")

        Worksheets("Events").Cells(j, 1).Value = FBDate
        Worksheets("Events").Cells(j, 2).Value = FBProp
        Worksheets("Events").Cells(j, 3).Value = FBName
        Worksheets("Events").Cells(j, 4).Value = FBpax
        Worksheets("Events").Cells(j, 5).Value = FBRate
        
        j = j + 1
    End If
i = i + 1
Loop

End Sub

Sub ConvertPkg()
Dim i, j As Integer, MTDate, MTProp, MTpax As String

Worksheets("Event Table").Select

i = 2
j = 4
Do Until Cells(i, 1).Value = 0
    If Cells(i, 28).Value > 0 Then
            If Cells(i, 4).Value = "Package Meeting" Then
                MTDate = Worksheets("Event Table").Cells(i, 1).Value
                MTProp = Worksheets("Event Table").Cells(i, 18).Value
                MTpax = Worksheets("Event Table").Cells(i, 8).Value
            
                Worksheets("Events").Cells(j, 15).Value = MTDate
                Worksheets("Events").Cells(j, 16).Value = MTProp
                Worksheets("Events").Cells(j, 18).Value = MTpax
            j = j + 1
            End If
    End If
i = i + 1
Loop

End Sub
