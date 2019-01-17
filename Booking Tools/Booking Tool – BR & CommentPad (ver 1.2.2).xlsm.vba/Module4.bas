Sub MtgPkg(ByRef w As Integer)
Dim b, c, d, j, z As Integer, d1, d2, d3, MtgName, Mtgpax, MtgPrice, MtgRevenue, TRevenue, NetRevenue As String

Rental z
c = z

'Mtg Pkg
Worksheets("Events").Select

If Cells(2, 37).Value > 0 Then
    c = c + 3
    Worksheets("CommentPad").Cells(c, 1).Value = "Meeting Package:"
    
    'VMRH Mtg Pkg
    If Cells(4, 37).Value > 0 Then
        b = 4
        Worksheets("CommentPad").Cells(c + 1, 1).Value = "Venetian:"
                Do Until Cells(b, 16).Value = ""
                    If Cells(b, 16).Value = "VMRH" Then
                        d1 = Worksheets("Events").Cells(b, 15).Value
                        d2 = Format(d1, "mmm")
                        d3 = Format(d1, "dd")
        
                        Worksheets("CommentPad").Cells(c + 2, 1).Value = d2 & ", " & d3
        
                        MtgName = Worksheets("Events").Cells(b, 17).Value
                        Mtgpax = Format(Worksheets("Events").Cells(b, 18).Value, "#,##0")
                        MtgPrice = Format(Worksheets("Events").Cells(b, 19).Value, "#,##0")
                        MtgRevenue = Format(Worksheets("Events").Cells(b, 20).Value, "#,##0.00")
                        Worksheets("CommentPad").Cells(c + 2, 2).Value = MtgName & " " & Mtgpax & "pax " & " @ " & MtgPrice & " = " & MtgRevenue
                        c = c + 1
                    End If
                b = b + 1
                
                Loop
                c = c + 2
    End If
    
    'PARIS Mtg Pkg
    If Cells(7, 37).Value > 0 Then
        b = 4
        Worksheets("CommentPad").Cells(c + 1, 1).Value = "Parisian:"
                Do Until Cells(b, 16).Value = ""
                    If Cells(b, 16).Value = "PARIS" Then
                        d1 = Worksheets("Events").Cells(b, 15).Value
                        d2 = Format(d1, "mmm")
                        d3 = Format(d1, "dd")
        
                        Worksheets("CommentPad").Cells(c + 2, 1).Value = d2 & ", " & d3
        
                        MtgName = Worksheets("Events").Cells(b, 17).Value
                        Mtgpax = Format(Worksheets("Events").Cells(b, 18).Value, "#,##0")
                        MtgPrice = Format(Worksheets("Events").Cells(b, 19).Value, "#,##0")
                        MtgRevenue = Format(Worksheets("Events").Cells(b, 20).Value, "#,##0.00")
                        Worksheets("CommentPad").Cells(c + 2, 2).Value = MtgName & " " & Mtgpax & "pax " & " @ " & MtgPrice & " = " & MtgRevenue
                        c = c + 1
                    End If
                b = b + 1
                
                Loop
                c = c + 2
    End If
    
    'CMCC Mtg Pkg
    If Cells(5, 37).Value > 0 Then
        b = 4
        Worksheets("CommentPad").Cells(c + 1, 1).Value = "Conrad:"
                Do Until Cells(b, 16).Value = ""
                    If Cells(b, 16).Value = "CMCC" Then
                        d1 = Worksheets("Events").Cells(b, 15).Value
                        d2 = Format(d1, "mmm")
                        d3 = Format(d1, "dd")
        
                        Worksheets("CommentPad").Cells(c + 2, 1).Value = d2 & ", " & d3
        
                        MtgName = Worksheets("Events").Cells(b, 17).Value
                        Mtgpax = Format(Worksheets("Events").Cells(b, 18).Value, "#,##0")
                        MtgPrice = Format(Worksheets("Events").Cells(b, 19).Value, "#,##0")
                        MtgRevenue = Format(Worksheets("Events").Cells(b, 20).Value, "#,##0.00")
                        Worksheets("CommentPad").Cells(c + 2, 2).Value = MtgName & " " & Mtgpax & "pax " & " @ " & MtgPrice & " = " & MtgRevenue
                        c = c + 1
                    End If
                b = b + 1
                
                Loop
                c = c + 2
    End If

    TRevenue = Format(Worksheets("Events").Cells(2, 20).Value, "#,##0.00")
    NetRevenue = Format(Worksheets("Events").Cells(2, 21).Value, "#,##0.00")
    Worksheets("CommentPad").Cells(c + 1, 1).Value = "Total Revenue : " & TRevenue & "+ (" & NetRevenue & ")"
    w = c
Else
    w = c

End If


End Sub

