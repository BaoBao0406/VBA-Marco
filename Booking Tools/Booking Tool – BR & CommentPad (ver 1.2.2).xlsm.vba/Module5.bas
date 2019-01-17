Sub Others()
Dim a, f, c, k, m, n, w As Integer, CValue, DValue, DeRRevenue, DeERevenue, DeRPercent, DeEPercent, DeName, DeAmount, CxName, CxAmount, Att, Comm As String


MtgPkg w
c = w + 1

Worksheets("Events").Select

'Contract Value
k = 1
Do Until Cells(k, 25).Value = "###"
    If Cells(k, 25).Value = "***" Then
        c = c + 2
        CValue = Format(Cells(k, 26).Value, "#,##0.00")
        Worksheets("CommentPad").Cells(c, 1).Value = "Contract value : " & CValue
    End If
k = k + 1
Loop

'Deposit Total
f = 15
Do Until Cells(f, 25).Value = "###"
    If Cells(f, 25).Value = "**" Then
        DValue = Format(Cells(f, 27).Value, "#,##0.00")
        DeRRevenue = Format(Cells(f - 2, 27).Value, "#,##0.00")
        DeERevenue = Format(Cells(f - 1, 27).Value, "#,##0.00")
        DeRPercent = Cells(f - 2, 26).Value * 100
        DeEPercent = Cells(f - 1, 26).Value * 100
    End If
f = f + 1
Loop

'Deposit Table
m = 1
Worksheets("CommentPad").Cells(c + 2, 1).Value = "Deposit: "
Do Until Cells(m, 25).Value = "###"
    If Cells(m, 25).Value = "****" Then
        n = 0
            DeName = Cells(m + n, 23).Value
            DeAmount = Format(Cells(m + n, 27).Value, "#,##0.00")

            Worksheets("CommentPad").Cells(c + 3, 1).Value = DeName
            Worksheets("CommentPad").Cells(c + 3, 2).Value = " $" & DeAmount
        c = c + 1
        n = n + 1
    End If

m = m + 1
Loop

Worksheets("CommentPad").Cells(c + 3, 1).Value = "Total Deposit : "
If DeRRevenue > 0 And DeERevenue > 0 Then
    Worksheets("CommentPad").Cells(c + 3, 2).Value = DValue & " = " & DeRRevenue & " + " & DeERevenue & " (Rooms " & DeRPercent & "% + Events " & DeEPercent & "%) "
ElseIf DeRRevenue > 0 And DeERevenue = 0 Then
    Worksheets("CommentPad").Cells(c + 3, 2).Value = DValue & " = " & DeRRevenue & " (Rooms only " & DeRPercent & "%)"
ElseIf DeRRevenue = 0 And DeERevenue > 0 Then
    Worksheets("CommentPad").Cells(c + 3, 2).Value = DValue & " = " & DeERevenue & " (Event only " & DeEPercent & "%)"
Else
    Worksheets("CommentPad").Cells(c + 3, 2).Value = "Waived"
End If

c = c + n + 1

'Cxl
m = 1
Worksheets("CommentPad").Cells(c + 3, 1).Value = "Cancellation: "
Do Until Cells(m, 25).Value = "####"
    If Cells(m, 25).Value = "*****" Then
        n = 0
            CxName = Cells(m + n, 23).Value
            CxAmount = Format(Cells(m + n, 27).Value, "#,##0.00")

            Worksheets("CommentPad").Cells(c + 4, 1).Value = CxName
            Worksheets("CommentPad").Cells(c + 4, 2).Value = " $" & CxAmount
        c = c + 1
        n = n + 1
    End If
m = m + 1
Loop
c = c + n + 3

'Others
a = 1
If Worksheets("Events").Range("AA2").Value > 0 Then
    Att = Worksheets("Events").Range("AA2").Value
    Worksheets("CommentPad").Cells(c + a, 1).Value = "Attrition: "
    Worksheets("CommentPad").Cells(c + a, 2).Value = Att & "%"
    c = c + a
Else
    a = 1
End If

If Worksheets("Events").Range("AA3").Value > 0 Then
    Comm = Worksheets("Events").Range("AA3").Value
    Worksheets("CommentPad").Cells(c + a, 1).Value = "Commission: "
    Worksheets("CommentPad").Cells(c + a, 2).Value = Comm & "%"
    c = c + a
Else
    a = 1
End If

Worksheets("CommentPad").Cells(c + 2, 1).Value = "Concessions: "

Worksheets("CommentPad").Cells(c + 4, 1).Value = "Contract signer: "
Worksheets("CommentPad").Cells(c + 5, 1).Value = "Title: "
Worksheets("CommentPad").Cells(c + 7, 1).Value = "Remark: "


End Sub