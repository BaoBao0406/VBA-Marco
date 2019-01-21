Sub Accommodation(ByRef x As Integer)
Dim Room, VMroom, PAroom, CMroom, HIroom, aRange As Range, i, j, k As Integer, RoomType, RoomRNs, RoomRate, bbfpax, bbfRate, Ferrypax, Tkpax, TkRate, TkName, d1, d2, d3 As String, d4 As Date

'VMRH Room
Worksheets("VM Room").Select
Set VMroom = Worksheets("VM Room").Range("C7").SpecialCells(xlCellTypeLastCell)

d4 = Date
d5 = Format(d4, "mmmm dd, yyyy")

Worksheets("CommentPad").Cells(1, 1).Value = "Amended on " & d5

Worksheets("CommentPad").Cells(2, 1).Value = "Accommodation:"

If Worksheets("VM Room").Cells(3, 1).Value > 0 Then
    i = 7
    j = 5
    Worksheets("CommentPad").Cells(4, 1).Value = Worksheets("VM Room").Cells(4, 1).Value + ":"
    Do Until Cells(i, 3).Value = 0
    
        d1 = Worksheets("VM Room").Cells(i, 1).Value
        d2 = Format(d1, "mmm")
        d3 = Format(d1, "dd")
    
        Worksheets("CommentPad").Cells(j, 1).Value = d2 & ", " & d3
        RoomType = Worksheets("VM Room").Cells(i, 2).Value
        RoomRNs = Format(Worksheets("VM Room").Cells(i, 3).Value, "#,##0")
        RoomRate = Format(Worksheets("VM Room").Cells(i, 4).Value, "#,##0.00")
        bbfpax = Worksheets("VM Room").Cells(i, 5).Value
        bbfRate = Worksheets("VM Room").Cells(i, 6).Value
        Ferrypax = Worksheets("VM Room").Cells(i, 7).Value
        TkName = Worksheets("VM Room").Cells(5, 9).Value
        Tkpax = Worksheets("VM Room").Cells(i, 9).Value
        TkRate = Worksheets("VM Room").Cells(i, 10).Value
    
        If Cells(i, 5).Value = 0 And Cells(i, 7).Value = 0 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++"
        ElseIf Cells(i, 5).Value >= 1 And Cells(i, 7).Value = 0 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ")"
        ElseIf Cells(i, 5).Value = 0 And Cells(i, 7).Value >= 1 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & Ferrypax & " Cotai water jet ticket"
        ElseIf Cells(i, 5).Value = 0 And Cells(i, 7).Value = 0 And Cells(i, 9).Value >= 1 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        ElseIf Cells(i, 5).Value >= 1 And Cells(i, 7).Value >= 1 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ") "
        ElseIf Cells(i, 5).Value >= 1 And Cells(i, 7).Value = 0 And Cells(i, 9).Value >= 1 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ") and " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        ElseIf Cells(i, 5).Value = 0 And Cells(i, 7).Value >= 1 And Cells(i, 9).Value >= 1 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & Ferrypax & " Cotai water jet ticket and " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        Else
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "RN @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ") and " & Ferrypax & " Cotai water jet ticket and " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        End If
        
        If Cells(i, 11).Value = "C" Or Cells(i, 11).Value = "c" Then
            Worksheets("CommentPad").Cells(j, 3).Value = "(Comp)"
        ElseIf Cells(i, 11).Value = "U" Or Cells(i, 11).Value = "u" Then
            Worksheets("CommentPad").Cells(j, 3).Value = "(Upgrade)"
        End If
        
    i = i + 1
    j = j + 1
    Loop

    Worksheets("VM Room").Select
    k = 7
    Do Until Worksheets("VM Room").Cells(k, 1).Value = "***"
    k = k + 1
    Loop
    Worksheets("CommentPad").Cells(j + 1, 1).Value = "Total Revenue : $" & Format(Worksheets("VM Room").Cells(k, 14).Value, "Standard") & "+ 15% ($" & Format(Worksheets("VM Room").Cells(k + 1, 14).Value, "Standard") & ")"
    
    x = j
Else
    x = j = 1
End If


'PARIS Room

Worksheets("PA Room").Select
Set PAroom = Worksheets("PA Room").Range("C7").SpecialCells(xlCellTypeLastCell)

If Worksheets("PA Room").Cells(3, 1).Value > 0 Then
    i = 7
    j = 5 + x
    Worksheets("CommentPad").Cells(4 + x, 1).Value = Worksheets("PA Room").Cells(4, 1).Value + ":"
    Do Until Worksheets("PA Room").Cells(i, 3).Value = 0
    
        d1 = Worksheets("PA Room").Cells(i, 1).Value
        d2 = Format(d1, "mmm")
        d3 = Format(d1, "dd")
    
        Worksheets("CommentPad").Cells(j, 1).Value = d2 & ", " & d3
        RoomType = Worksheets("PA Room").Cells(i, 2).Value
        RoomRNs = Format(Worksheets("PA Room").Cells(i, 3).Value, "#,##0")
        RoomRate = Format(Worksheets("PA Room").Cells(i, 4).Value, "#,##0.00")
        bbfpax = Worksheets("PA Room").Cells(i, 5).Value
        bbfRate = Worksheets("PA Room").Cells(i, 6).Value
        Ferrypax = Worksheets("PA Room").Cells(i, 7).Value
        TkName = Worksheets("VM Room").Cells(5, 9).Value
        Tkpax = Worksheets("VM Room").Cells(i, 9).Value
        TkRate = Worksheets("VM Room").Cells(i, 10).Value
    
        If Cells(i, 5).Value = 0 And Cells(i, 7).Value = 0 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++"
        ElseIf Cells(i, 5).Value >= 1 And Cells(i, 7).Value = 0 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ")"
        ElseIf Cells(i, 5).Value = 0 And Cells(i, 7).Value >= 1 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & Ferrypax & " Cotai water jet ticket"
        ElseIf Cells(i, 5).Value = 0 And Cells(i, 7).Value = 0 And Cells(i, 9).Value >= 1 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        ElseIf Cells(i, 5).Value >= 1 And Cells(i, 7).Value >= 1 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ") "
        ElseIf Cells(i, 5).Value >= 1 And Cells(i, 7).Value = 0 And Cells(i, 9).Value >= 1 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ") and " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        ElseIf Cells(i, 5).Value = 0 And Cells(i, 7).Value >= 1 And Cells(i, 9).Value >= 1 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & Ferrypax & " Cotai water jet ticket and " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        Else
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ") and " & Ferrypax & " Cotai water jet ticket and " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        End If
        
        If Cells(i, 11).Value = "C" Or Cells(i, 11).Value = "c" Then
            Worksheets("CommentPad").Cells(j, 3).Value = "(Comp)"
        ElseIf Cells(i, 11).Value = "U" Or Cells(i, 11).Value = "u" Then
            Worksheets("CommentPad").Cells(j, 3).Value = "(Upgrade)"
        End If
    
    i = i + 1
    j = j + 1
    Loop

    Worksheets("PA Room").Select
    k = 7
    Do Until Worksheets("PA Room").Cells(k, 1).Value = "***"
    k = k + 1
    Loop
    Worksheets("CommentPad").Cells(j + 1, 1).Value = "Total Revenue : $" & Format(Worksheets("PA Room").Cells(k, 14).Value, "Standard") & "+ 15% ($" & Format(Worksheets("PA Room").Cells(k + 1, 14).Value, "Standard") & ")"
    
    x = j
Else
    x = j
End If

'CMCC Room
Worksheets("CM Room").Select
Set CMroom = Worksheets("CM Room").Range("C7").SpecialCells(xlCellTypeLastCell)

If Worksheets("CM Room").Cells(3, 1).Value > 0 Then
    i = 7
    j = 5 + x
    Worksheets("CommentPad").Cells(4 + x, 1).Value = Worksheets("CM Room").Cells(4, 1).Value + ":"
    Do Until Cells(i, 3).Value = 0
    
        d1 = Worksheets("CM Room").Cells(i, 1).Value
        d2 = Format(d1, "mmm")
        d3 = Format(d1, "dd")
    
        Worksheets("CommentPad").Cells(j, 1).Value = d2 & ", " & d3
        RoomType = Worksheets("CM Room").Cells(i, 2).Value
        RoomRNs = Format(Worksheets("CM Room").Cells(i, 3).Value, "#,##0")
        RoomRate = Format(Worksheets("CM Room").Cells(i, 4).Value, "#,##0.00")
        bbfpax = Worksheets("CM Room").Cells(i, 5).Value
        bbfRate = Worksheets("CM Room").Cells(i, 6).Value
        Ferrypax = Worksheets("CM Room").Cells(i, 7).Value
        TkName = Worksheets("VM Room").Cells(5, 9).Value
        Tkpax = Worksheets("VM Room").Cells(i, 9).Value
        TkRate = Worksheets("VM Room").Cells(i, 10).Value
    
        If Cells(i, 5).Value = 0 And Cells(i, 7).Value = 0 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++"
        ElseIf Cells(i, 5).Value >= 1 And Cells(i, 7).Value = 0 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ")"
        ElseIf Cells(i, 5).Value = 0 And Cells(i, 7).Value >= 1 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & Ferrypax & " Cotai water jet ticket"
        ElseIf Cells(i, 5).Value = 0 And Cells(i, 7).Value = 0 And Cells(i, 9).Value >= 1 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        ElseIf Cells(i, 5).Value >= 1 And Cells(i, 7).Value >= 1 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ") "
        ElseIf Cells(i, 5).Value >= 1 And Cells(i, 7).Value = 0 And Cells(i, 9).Value >= 1 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ") and " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        ElseIf Cells(i, 5).Value = 0 And Cells(i, 7).Value >= 1 And Cells(i, 9).Value >= 1 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & Ferrypax & " Cotai water jet ticket and " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        Else
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ") and " & Ferrypax & " Cotai water jet ticket and " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        End If
        
        If Cells(i, 11).Value = "C" Or Cells(i, 11).Value = "c" Then
            Worksheets("CommentPad").Cells(j, 3).Value = "(Comp)"
        ElseIf Cells(i, 11).Value = "U" Or Cells(i, 11).Value = "u" Then
            Worksheets("CommentPad").Cells(j, 3).Value = "(Upgrade)"
        End If
        
    i = i + 1
    j = j + 1
    Loop

    Worksheets("CM Room").Select
    k = 7
    Do Until Worksheets("CM Room").Cells(k, 1).Value = "***"
    k = k + 1
    Loop
    Worksheets("CommentPad").Cells(j + 1, 1).Value = "Total Revenue : $" & Format(Worksheets("CM Room").Cells(k, 14).Value, "Standard") & "+ 15% ($" & Format(Worksheets("CM Room").Cells(k + 1, 14).Value, "Standard") & ")"
    
    x = j
Else
    x = j
End If

'HICC Room
Worksheets("HI Room").Select
Set HIroom = Worksheets("HI Room").Range("C7").SpecialCells(xlCellTypeLastCell)

If Worksheets("HI Room").Cells(3, 1).Value > 0 Then
    i = 7
    j = 5 + x
    Worksheets("CommentPad").Cells(4 + x, 1).Value = Worksheets("HI Room").Cells(4, 1).Value + ":"
    Do Until Cells(i, 3).Value = 0
    
        d1 = Worksheets("HI Room").Cells(i, 1).Value
        d2 = Format(d1, "mmm")
        d3 = Format(d1, "dd")
    
        Worksheets("CommentPad").Cells(j, 1).Value = d2 & ", " & d3
        RoomType = Worksheets("HI Room").Cells(i, 2).Value
        RoomRNs = Format(Worksheets("HI Room").Cells(i, 3).Value, "#,##0")
        RoomRate = Format(Worksheets("HI Room").Cells(i, 4).Value, "#,##0.00")
        bbfpax = Worksheets("HI Room").Cells(i, 5).Value
        bbfRate = Worksheets("HI Room").Cells(i, 6).Value
        Ferrypax = Worksheets("HI Room").Cells(i, 7).Value
        TkName = Worksheets("VM Room").Cells(5, 9).Value
        Tkpax = Worksheets("VM Room").Cells(i, 9).Value
        TkRate = Worksheets("VM Room").Cells(i, 10).Value
    
        If Cells(i, 5).Value = 0 And Cells(i, 7).Value = 0 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++"
        ElseIf Cells(i, 5).Value >= 1 And Cells(i, 7).Value = 0 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ")"
        ElseIf Cells(i, 5).Value = 0 And Cells(i, 7).Value >= 1 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & Ferrypax & " Cotai water jet ticket"
        ElseIf Cells(i, 5).Value = 0 And Cells(i, 7).Value = 0 And Cells(i, 9).Value >= 1 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        ElseIf Cells(i, 5).Value >= 1 And Cells(i, 7).Value >= 1 And Cells(i, 9).Value = 0 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ") "
        ElseIf Cells(i, 5).Value >= 1 And Cells(i, 7).Value = 0 And Cells(i, 9).Value >= 1 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ") and " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        ElseIf Cells(i, 5).Value = 0 And Cells(i, 7).Value >= 1 And Cells(i, 9).Value >= 1 Then
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & Ferrypax & " Cotai water jet ticket and " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        Else
            Worksheets("CommentPad").Cells(j, 2).Value = RoomType & " " & RoomRNs & "rn @ " & RoomRate & "++ with " & bbfpax & "bbf ($" & bbfRate & ") and " & Ferrypax & " Cotai water jet ticket and " & TkName & " Ticket " & Tkpax & "pax @ $" & TkRate
        End If
        
        If Cells(i, 11).Value = "C" Or Cells(i, 11).Value = "c" Then
            Worksheets("CommentPad").Cells(j, 3).Value = "(Comp)"
        ElseIf Cells(i, 11).Value = "U" Or Cells(i, 11).Value = "u" Then
            Worksheets("CommentPad").Cells(j, 3).Value = "(Upgrade)"
        End If
        
    i = i + 1
    j = j + 1
    Loop

    Worksheets("HI Room").Select
    k = 7
    Do Until Worksheets("HI Room").Cells(k, 1).Value = "***"
    k = k + 1
    Loop
    Worksheets("CommentPad").Cells(j + 1, 1).Value = "Total Revenue : $" & Format(Worksheets("HI Room").Cells(k, 14).Value, "Standard") & "+ 15% ($" & Format(Worksheets("HI Room").Cells(k + 1, 14).Value, "Standard") & ")"
    
    x = j
Else
    x = j
End If



End Sub

