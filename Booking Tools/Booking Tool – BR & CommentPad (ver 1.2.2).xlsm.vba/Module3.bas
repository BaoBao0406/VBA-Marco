Sub Rental(ByRef z As Integer)
Dim b, c, d, j, y As Integer, d1, d2, d3, RName, RRevenue, TRevenue, NetRevenue As String

FBmin y
c = y

'Rental
Worksheets("Events").Select

If Cells(2, 36).Value > 0 Then
    c = c + 3
    Worksheets("CommentPad").Cells(c, 1).Value = "Rental:"
    
    'VMRH Rental
    If Cells(4, 36).Value > 0 Then
        b = 4
        Worksheets("CommentPad").Cells(c + 1, 1).Value = "Venetian:"
                Do Until Cells(b, 10).Value = ""
                    If Cells(b, 10).Value = "VMRH" Then
                        d1 = Worksheets("Events").Cells(b, 9).Value
                        d2 = Format(d1, "mmm")
                        d3 = Format(d1, "dd")
        
                        Worksheets("CommentPad").Cells(c + 2, 1).Value = d2 & ", " & d3
        
                        RName = Worksheets("Events").Cells(b, 11).Value
                        RRevenue = Format(Worksheets("Events").Cells(b, 12).Value, "#,##0.00")
                        If RRevenue > 0 Then
                            Worksheets("CommentPad").Cells(c + 2, 2).Value = RName & " = " & RRevenue
                        Else
                            Worksheets("CommentPad").Cells(c + 2, 2).Value = RName & " = waived"
                        End If
                        c = c + 1
                    End If
                b = b + 1
                
                Loop
                c = c + 2
    End If
    
    'PARIS Rental
    If Cells(7, 36).Value > 0 Then
        b = 4
        Worksheets("CommentPad").Cells(c + 1, 1).Value = "Parisian:"
                Do Until Cells(b, 10).Value = ""
                    If Cells(b, 10).Value = "PARIS" Then
                        d1 = Worksheets("Events").Cells(b, 9).Value
                        d2 = Format(d1, "mmm")
                        d3 = Format(d1, "dd")
        
                        Worksheets("CommentPad").Cells(c + 2, 1).Value = d2 & ", " & d3
        
                        RName = Worksheets("Events").Cells(b, 11).Value
                        RRevenue = Format(Worksheets("Events").Cells(b, 12).Value, "#,##0.00")
                        If RRevenue > 0 Then
                            Worksheets("CommentPad").Cells(c + 2, 2).Value = RName & " = " & RRevenue
                        Else
                            Worksheets("CommentPad").Cells(c + 2, 2).Value = RName & " = waived"
                        End If
                        c = c + 1
                    End If
                b = b + 1
                
                Loop
                c = c + 2
    End If
    
    'CMCC Rental
    If Cells(5, 36).Value > 0 Then
        b = 4
        Worksheets("CommentPad").Cells(c + 1, 1).Value = "Conrad:"
                Do Until Cells(b, 10).Value = ""
                    If Cells(b, 10).Value = "PARIS" Then
                        d1 = Worksheets("Events").Cells(b, 9).Value
                        d2 = Format(d1, "mmm")
                        d3 = Format(d1, "dd")
        
                        Worksheets("CommentPad").Cells(c + 2, 1).Value = d2 & ", " & d3
        
                        RName = Worksheets("Events").Cells(b, 11).Value
                        RRevenue = Format(Worksheets("Events").Cells(b, 12).Value, "#,##0.00")
                        If RRevenue > 0 Then
                            Worksheets("CommentPad").Cells(c + 2, 2).Value = RName & " = " & RRevenue
                        Else
                            Worksheets("CommentPad").Cells(c + 2, 2).Value = RName & " = waived"
                        End If
                        c = c + 1
                    End If
                b = b + 1
                
                Loop
                c = c + 2
    End If
    
    TRevenue = Format(Worksheets("Events").Cells(2, 12).Value, "#,##0.00")
    NetRevenue = Format(Worksheets("Events").Cells(2, 13).Value, "#,##0.00")
    Worksheets("CommentPad").Cells(c + 1, 1).Value = "Total Revenue : " & TRevenue & "+ (" & NetRevenue & ")"
    z = c
Else
    z = c

End If


End Sub

