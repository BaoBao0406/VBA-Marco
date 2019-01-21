Sub FBmin(ByRef y As Integer)
Dim b, c, d, j, x As Integer, d1, d2, d3, FBName, FBpax, FBPrice, FBRevenue, TRevenue, NetRevenue As String

Accommodation x
c = x

'F&B min
Worksheets("Events").Select

If Cells(2, 35).Value > 0 Then
    c = c + 3
    Worksheets("CommentPad").Cells(c, 1).Value = "F&B min:"
    
    'VMRH F&B min
    If Cells(4, 35).Value > 0 Then
        b = 4
        Worksheets("CommentPad").Cells(c + 1, 1).Value = "Venetian:"
                Do Until Cells(b, 2).Value = ""
                    If Cells(b, 2).Value = "VMRH" Then
                        d1 = Worksheets("Events").Cells(b, 1).Value
                        d2 = Format(d1, "mmm")
                        d3 = Format(d1, "dd")
        
                        Worksheets("CommentPad").Cells(c + 2, 1).Value = d2 & ", " & d3
        
                        FBName = Worksheets("Events").Cells(b, 3).Value
                        FBpax = Format(Worksheets("Events").Cells(b, 4).Value, "#,##0")
                        FBPrice = Format(Worksheets("Events").Cells(b, 5).Value, "#,##0.00")
                        FBRevenue = Format(Worksheets("Events").Cells(b, 6).Value, "#,##0.00")
                        If Worksheets("Events").Cells(b, 4).Value > 0 Then
                            Worksheets("CommentPad").Cells(c + 2, 2).Value = FBpax & "pax " & FBName & " @ " & FBPrice & " = " & FBRevenue
                        Else
                            Worksheets("CommentPad").Cells(c + 2, 2).Value = FBName & " = " & FBRevenue
                        End If
                        c = c + 1
                    End If
                b = b + 1
                
                Loop
                c = c + 2
    End If

    'PARIS F&B min
    If Cells(7, 35).Value > 0 Then
        b = 4
        Worksheets("CommentPad").Cells(c + 1, 1).Value = "Parisian:"
                Do Until Cells(b, 2).Value = ""
                    If Cells(b, 2).Value = "PARIS" Then
                        d1 = Worksheets("Events").Cells(b, 1).Value
                        d2 = Format(d1, "mmm")
                        d3 = Format(d1, "dd")
        
                        Worksheets("CommentPad").Cells(c + 2, 1).Value = d2 & ", " & d3
        
                        FBName = Worksheets("Events").Cells(b, 3).Value
                        FBpax = Format(Worksheets("Events").Cells(b, 4).Value, "#,##0")
                        FBPrice = Format(Worksheets("Events").Cells(b, 5).Value, "#,##0.00")
                        FBRevenue = Format(Worksheets("Events").Cells(b, 6).Value, "#,##0.00")
                        If Worksheets("Events").Cells(b, 4).Value > 0 Then
                            Worksheets("CommentPad").Cells(c + 2, 2).Value = FBpax & "pax " & FBName & " @ " & FBPrice & " = " & FBRevenue
                        Else
                            Worksheets("CommentPad").Cells(c + 2, 2).Value = FBName & " = " & FBRevenue
                        End If
                        c = c + 1
                    End If
                b = b + 1
                Loop
                c = c + 2
    End If
    
    'CMCC F&B min
    If Cells(5, 35).Value > 0 Then
        b = 4
        Worksheets("CommentPad").Cells(c + 1, 1).Value = "Conrad:"
                Do Until Cells(b, 2).Value = ""
                    If Cells(b, 2).Value = "CMCC" Then
                        d1 = Worksheets("Events").Cells(b, 1).Value
                        d2 = Format(d1, "mmm")
                        d3 = Format(d1, "dd")
        
                        Worksheets("CommentPad").Cells(c + 2, 1).Value = d2 & ", " & d3
        
                        FBName = Worksheets("Events").Cells(b, 3).Value
                        FBpax = Format(Worksheets("Events").Cells(b, 4).Value, "#,##0")
                        FBPrice = Format(Worksheets("Events").Cells(b, 5).Value, "#,##0.00")
                        FBRevenue = Format(Worksheets("Events").Cells(b, 6).Value, "#,##0.00")
                        If Worksheets("Events").Cells(b, 4).Value > 0 Then
                            Worksheets("CommentPad").Cells(c + 2, 2).Value = FBpax & "pax " & FBName & " @ " & FBPrice & " = " & FBRevenue
                        Else
                            Worksheets("CommentPad").Cells(c + 2, 2).Value = FBName & " = " & FBRevenue
                        End If
                        c = c + 1
                    End If
                b = b + 1
                Loop
                c = c + 2
    End If
    
    'HIMCC F&B min
    If Cells(6, 35).Value > 0 Then
        b = 4
        Worksheets("CommentPad").Cells(c + 1, 1).Value = "Holiday Inn:"
                Do Until Cells(b, 2).Value = ""
                    If Cells(b, 2).Value = "HIMCC" Then
                        d1 = Worksheets("Events").Cells(b, 1).Value
                        d2 = Format(d1, "mmm")
                        d3 = Format(d1, "dd")
        
                        Worksheets("CommentPad").Cells(c + 2, 1).Value = d2 & ", " & d3
        
                        FBName = Worksheets("Events").Cells(b, 3).Value
                        FBpax = Format(Worksheets("Events").Cells(b, 4).Value, "#,##0")
                        FBPrice = Format(Worksheets("Events").Cells(b, 5).Value, "#,##0.00")
                        FBRevenue = Format(Worksheets("Events").Cells(b, 6).Value, "#,##0.00")
                        If Worksheets("Events").Cells(b, 4).Value > 0 Then
                            Worksheets("CommentPad").Cells(c + 2, 2).Value = FBpax & "pax " & FBName & " @ " & FBPrice & " = " & FBRevenue
                        Else
                            Worksheets("CommentPad").Cells(c + 2, 2).Value = FBName & " = " & FBRevenue
                        End If
                        c = c + 1
                    End If
                b = b + 1
                Loop
                c = c + 2
    End If
    TRevenue = Format(Worksheets("Events").Cells(2, 6).Value, "#,##0.00")
    NetRevenue = Format(Worksheets("Events").Cells(2, 7).Value, "#,##0.00")
    Worksheets("CommentPad").Cells(c + 1, 1).Value = "Total Revenue : " & TRevenue & "+ (" & NetRevenue & ")"
    y = c
Else
    y = c

End If


End Sub
