Sub BRButton()
Dim ThisWbk, BRWbk As Workbook, aRange, bRange, cRange, dRange, eRange, fRange, gRange, hRange, iRange As Range, i, j, k, x, y, z, w As Integer, FBmin, PostAs, AcName, AgName, Region, Owner, BkType, Industry, BkID, Comm, Attr, LOS, Prop, HM As String, d1, d2 As Date

BRRoom

'Copy Event Table
    Set aRange = Worksheets("Event Table").Cells(346, 9)
    Set bRange = Worksheets("Event Table").Range("A2")
    Set cRange = Worksheets("Event Table").Range(bRange, aRange)

'Copy Booking Info
    PostAs = Worksheets("BK Info").Cells(2, 2).Value
    AcName = Worksheets("BK Info").Cells(3, 2).Value
    AgName = Worksheets("BK Info").Cells(4, 2).Value
    Region = Worksheets("BK Info").Cells(5, 2).Value
    Owner = Worksheets("BK Info").Cells(6, 2).Value
    BkType = Worksheets("BK Info").Cells(7, 2).Value
    Industry = Worksheets("BK Info").Cells(8, 2).Value
    BkID = Worksheets("BK Info").Cells(9, 2).Value
    Comm = Worksheets("BK Info").Cells(10, 2).Value / 100
    Attr = Worksheets("BK Info").Cells(11, 2).Value / 100
    d1 = Worksheets("BK Info").Cells(12, 2).Value
    d2 = Worksheets("BK Info").Cells(13, 2).Value
    Status = Worksheets("BK Info").Cells(14, 2).Value
    LOS = Worksheets("BK Info").Cells(19, 2).Value
    FBmin = Worksheets("Events").Cells(2, 6).Value
    
    'Rooming List
    If Worksheets("BK Info").Cells(15, 2).Value = "Individual Resv" Then
        HM = "Call-In"
    ElseIf Worksheets("BK Info").Cells(15, 2).Value = "Rooming List" Then
        HM = "Rooming List"
    End If
    
    'Booking ID
    If Worksheets("BK Info").Cells(16, 2).Value = "The Venetian Macao" Then
        k = 1
    ElseIf Worksheets("BK Info").Cells(16, 2).Value = "The Parisian Macao" Then
        k = 4
    ElseIf Worksheets("BK Info").Cells(16, 2).Value = "Conrad Macao Cotai Central" Then
        k = 2
    ElseIf Worksheets("BK Info").Cells(16, 2).Value = "Holiday Inn Macao Cotai Central" Then
        k = 3
    End If

'Copy BR Room
    Set dRange = Range("A34")
    Set eRange = Range("A33").End(xlDown)
    Set eRange = eRange.Offset(0, 1)
    Set fRange = Range(dRange, eRange)
    
    x = Worksheets("BK Info").Cells(19, 2).Value
    y = fRange.Rows.Count
    
    Set gRange = Range("G34")
    Set hRange = eRange.Offset(0, x + 5)
    Set iRange = Range(gRange, hRange)

'Open BR and copy
Workbooks.Open Filename:="X:\VML\Sales\Business_Review\BR Form_Macao_5.0.xlsm"
Set ThisWbk = ThisWorkbook
Set BRWbk = ActiveWorkbook

'Paste Event Table
    '**Must if you need to select different cells in different workbook**
    Workbooks("BR Form_Macao_5.0").Worksheets("Meeting Space").Activate
    BRWbk.Worksheets("Meeting Space").Unprotect ("mode")
    
    cRange.Font.Bold = False
    cRange.Font.Size = 10
    cRange.Borders(xlEdgeTop).Weight = xlHairline
    cRange.Borders(xlEdgeBottom).Weight = xlHairline
    cRange.Borders(xlEdgeLeft).Weight = xlHairline
    cRange.Borders(xlEdgeRight).Weight = xlHairline
    cRange.Borders(xlInsideVertical).Weight = xlHairline
    cRange.Borders(xlInsideHorizontal).Weight = xlHairline
    cRange.Copy Destination:=BRWbk.Worksheets("Meeting Space").Range("B24")
    
    BRWbk.Worksheets("Meeting Space").Range("G18:J20,B24:J373").Locked = False
    BRWbk.Worksheets("Meeting Space").Protect ("mode")

'Paste Booking Info
    Workbooks("BR Form_Macao_5.0").Worksheets("Rooms").Activate
    BRWbk.Worksheets("Rooms").Unprotect ("mode")
    
    BRWbk.Worksheets("Rooms").Cells(2, 2).Value = PostAs
    BRWbk.Worksheets("Rooms").Cells(4, 2).Value = AcName
    BRWbk.Worksheets("Rooms").Cells(5, 2).Value = AgName
    BRWbk.Worksheets("Rooms").Cells(6, 2).Value = Region
    BRWbk.Worksheets("Rooms").Cells(3, 10).Value = Owner
    BRWbk.Worksheets("Rooms").Cells(4, 10).Value = BkType
    BRWbk.Worksheets("Rooms").Cells(6, 10).Value = Industry
    BRWbk.Worksheets("Rooms").Cells(1 + k, 15).Value = BkID
    BRWbk.Worksheets("Rooms").Cells(6, 15).Value = Comm
    BRWbk.Worksheets("Rooms").Cells(7, 15).Value = Attr
    BRWbk.Worksheets("Rooms").Cells(8, 15).Value = HM
    BRWbk.Worksheets("Rooms").Cells(14, 2).Value = Status
    BRWbk.Worksheets("Rooms").Cells(15, 2).Value = d1
    BRWbk.Worksheets("Rooms").Cells(16, 2).Value = LOS
    If k = 1 Then
        BRWbk.Worksheets("Rooms").Cells(28, 2).Value = FBmin
    ElseIf k = 2 Then
        BRWbk.Worksheets("Rooms").Cells(37, 2).Value = FBmin
    ElseIf k = 4 Then
        BRWbk.Worksheets("Rooms").Cells(45, 2).Value = FBmin
    End If
    
'Paste BRRoom
    fRange.Copy Destination:=BRWbk.Worksheets("Rooms").Range("B70")
    iRange.Copy Destination:=BRWbk.Worksheets("Rooms").Range("G70")
    
'Add Date Column and Room Type Row
    'Date Column
    If x > 9 Then
        z = 1
        For z = 1 To (x - 9)
            Application.Run "'BR Form_Macao_5.0.xlsm'!UnhideColRequest1"
        Next z
    End If
    
    'Room Type Row
    If y > 4 Then
        w = 1
        For w = 1 To (y - 4)
            Application.Run "'BR Form_Macao_5.0.xlsm'!UnhideRowsRequest1"
        Next w
    End If
    
BRWbk.Worksheets("Rooms").Range("B14:B17,B28:B33,B37:B41,B45:B50,B60,A70:C85,E70:E85,G70:AL85,A169:N170").Locked = False

BRWbk.Worksheets("Rooms").Protect ("mode")



End Sub

Sub BRRoom()
Dim i, j, k, x As Integer, aRange, bRange, cRange, dRange, eRange, xRange, yRange, zRange As Range, Bbf, Ferry, Ticket, Date1 As String

CleanUpBRRoom

i = 34
'VMRH

Worksheets("VM Room").Select
If Cells(7, 21).Value > 0 Then

    k = 7
    Do Until Worksheets("VM Room").Cells(k, 1).Value = "***"
    k = k + 1
    Loop
    k = k - 2
    
    x = 4
    j = 7
    For j = 7 To k
        If Cells(j, 1).Value > 0 Then
            Worksheets("BK Info").Cells(i, 1).Value = "Venetian"
            Worksheets("BK Info").Cells(i, 3).Value = Cells(j, 2).Value
            
            Set aRange = Worksheets("BK Info").Range("G33")
            Set bRange = aRange.End(xlToRight)
            Set cRange = Range(aRange, bRange)
            
            Date1 = Cells(j, 1).Value
            Bbf = Cells(j, 5).Value
            Ferry = Cells(j, 7).Value
            TicketName = Cells(5, 9).Value
            Ticket = Cells(j, 9).Value
            If Cells(j, 5).Value > 0 And Cells(j, 7).Value = 0 And Cells(j, 9).Value = 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF"
            ElseIf Cells(j, 5).Value = 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value = 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Ferry & " Ferry ticket"
            ElseIf Cells(j, 5).Value = 0 And Cells(j, 7).Value = 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Ticket & " " & TicketName
            ElseIf Cells(j, 5).Value > 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value = 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF + " & Ferry & " Ferry ticket"
            ElseIf Cells(j, 5).Value = 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Ferry & " Ferry ticket + " & Ticket & " " & TicketName
            ElseIf Cells(j, 5).Value > 0 And Cells(j, 7).Value = 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF + " & Ticket & " " & TicketName
            ElseIf Cells(j, 5).Value > 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF + " & Ferry & " Ferry ticket + " & Ticket & " " & TicketName
            End If
            
            For Each dRange In cRange
                If dRange.Value = Date1 Then
                    Set eRange = dRange.Offset(i - 33, 0)
                    eRange.Value = Cells(j, 3)
            
                End If
            Next
            
        i = i + 1
        End If
    Next j
Else
    i = i
End If

'PARIS

Worksheets("PA Room").Select
If Cells(7, 21).Value > 0 Then

    k = 7
    Do Until Worksheets("PA Room").Cells(k, 1).Value = "***"
    k = k + 1
    Loop
    k = k - 2
    
    x = 4
    j = 7
    For j = 7 To k
        If Cells(j, 1).Value > 0 Then
            Worksheets("BK Info").Cells(i, 1).Value = "Parisian"
            Worksheets("BK Info").Cells(i, 3).Value = Cells(j, 2).Value
            
            Set aRange = Worksheets("BK Info").Range("G33")
            Set bRange = aRange.End(xlToRight)
            Set cRange = Range(aRange, bRange)
            
            Date1 = Cells(j, 1).Value
            Bbf = Cells(j, 5).Value
            Ferry = Cells(j, 7).Value
            TicketName = Cells(5, 9).Value
            Ticket = Cells(j, 9).Value
            If Cells(j, 5).Value > 0 And Cells(j, 7).Value = 0 And Cells(j, 9).Value = 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF"
            ElseIf Cells(j, 5).Value = 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value = 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Ferry & " Ferry ticket"
            ElseIf Cells(j, 5).Value = 0 And Cells(j, 7).Value = 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Ticket & " " & TicketName
            ElseIf Cells(j, 5).Value > 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value = 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF + " & Ferry & " Ferry ticket"
            ElseIf Cells(j, 5).Value = 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Ferry & " Ferry ticket + " & Ticket & " " & TicketName
            ElseIf Cells(j, 5).Value > 0 And Cells(j, 7).Value = 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF + " & Ticket & " " & TicketName
            ElseIf Cells(j, 5).Value > 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF + " & Ferry & " Ferry ticket + " & Ticket & " " & TicketName
            End If
            
            For Each dRange In cRange
                If dRange.Value = Date1 Then
                    Set eRange = dRange.Offset(i - 33, 0)
                    eRange.Value = Cells(j, 3)
            
                End If
            Next
            
        i = i + 1
        End If
    Next j
Else
    i = i
End If

'CMCC

Worksheets("CM Room").Select
If Cells(7, 21).Value > 0 Then

    k = 7
    Do Until Worksheets("CM Room").Cells(k, 1).Value = "***"
    k = k + 1
    Loop
    k = k - 2
    
    x = 4
    j = 7
    For j = 7 To k
        If Cells(j, 1).Value > 0 Then
            Worksheets("BK Info").Cells(i, 1).Value = "Conrad"
            Worksheets("BK Info").Cells(i, 3).Value = Cells(j, 2).Value
            
            Set aRange = Worksheets("BK Info").Range("G33")
            Set bRange = aRange.End(xlToRight)
            Set cRange = Range(aRange, bRange)
            
            Date1 = Cells(j, 1).Value
            Bbf = Cells(j, 5).Value
            Ferry = Cells(j, 7).Value
            TicketName = Cells(5, 9).Value
            Ticket = Cells(j, 9).Value
            If Cells(j, 5).Value > 0 And Cells(j, 7).Value = 0 And Cells(j, 9).Value = 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF"
            ElseIf Cells(j, 5).Value = 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value = 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Ferry & " Ferry ticket"
            ElseIf Cells(j, 5).Value = 0 And Cells(j, 7).Value = 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Ticket & " " & TicketName
            ElseIf Cells(j, 5).Value > 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value = 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF + " & Ferry & " Ferry ticket"
            ElseIf Cells(j, 5).Value = 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Ferry & " Ferry ticket + " & Ticket & " " & TicketName
            ElseIf Cells(j, 5).Value > 0 And Cells(j, 7).Value = 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF + " & Ticket & " " & TicketName
            ElseIf Cells(j, 5).Value > 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF + " & Ferry & " Ferry ticket + " & Ticket & " " & TicketName
            End If
            
            For Each dRange In cRange
                If dRange.Value = Date1 Then
                    Set eRange = dRange.Offset(i - 33, 0)
                    eRange.Value = Cells(j, 3)
            
                End If
            Next
            
        i = i + 1
        End If
    Next j
Else
    i = i
End If

'HIMCC

Worksheets("HI Room").Select
If Cells(7, 21).Value > 0 Then

    k = 7
    Do Until Worksheets("HI Room").Cells(k, 1).Value = "***"
    k = k + 1
    Loop
    k = k - 2
    
    x = 4
    j = 7
    For j = 7 To k
        If Cells(j, 1).Value > 0 Then
            Worksheets("BK Info").Cells(i, 1).Value = "Holiday Inn"
            Worksheets("BK Info").Cells(i, 3).Value = Cells(j, 2).Value
            
            Set aRange = Worksheets("BK Info").Range("G33")
            Set bRange = aRange.End(xlToRight)
            Set cRange = Range(aRange, bRange)
            
            Date1 = Cells(j, 1).Value
            Bbf = Cells(j, 5).Value
            Ferry = Cells(j, 7).Value
            TicketName = Cells(5, 9).Value
            Ticket = Cells(j, 9).Value
            If Cells(j, 5).Value > 0 And Cells(j, 7).Value = 0 And Cells(j, 9).Value = 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF"
            ElseIf Cells(j, 5).Value = 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value = 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Ferry & " Ferry ticket"
            ElseIf Cells(j, 5).Value = 0 And Cells(j, 7).Value = 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Ticket & " " & TicketName
            ElseIf Cells(j, 5).Value > 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value = 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF + " & Ferry & " Ferry ticket"
            ElseIf Cells(j, 5).Value = 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Ferry & " Ferry ticket + " & Ticket & " " & TicketName
            ElseIf Cells(j, 5).Value > 0 And Cells(j, 7).Value = 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF + " & Ticket & " " & TicketName
            ElseIf Cells(j, 5).Value > 0 And Cells(j, 7).Value > 0 And Cells(j, 9).Value > 0 Then
                Worksheets("BK Info").Cells(i, 4).Value = Bbf & " BBF + " & Ferry & " Ferry ticket + " & Ticket & " " & TicketName
            End If
            
            For Each dRange In cRange
                If dRange.Value = Date1 Then
                    Set eRange = dRange.Offset(i - 33, 0)
                    eRange.Value = Cells(j, 3)
            
                End If
            Next
            
        i = i + 1
        End If
    Next j
Else
    i = i
End If


'Copy Formula
Worksheets("BK Info").Range("B34:B200").Formula = Worksheets("BK Info").Range("B32").Formula
Worksheets("BK Info").Range("E34:E200").Formula = Worksheets("BK Info").Range("E32").Formula
Worksheets("BK Info").Range("F34:F200").Formula = Worksheets("BK Info").Range("F32").Formula
Worksheets("BK Info").Range("B34:B200").Formula = Worksheets("BK Info").Range("B34:B200").Value

'Summarize the table
SummarizeBRRoom

End Sub

Sub SummarizeBRRoom()
Dim Rng As Range, Dn As Range, n As Long, nRng As Range, i, k As Integer

Worksheets("BK Info").Select

Set Rng = Range(Range("B34"), Range("B" & Rows.Count).End(xlUp))
With CreateObject("scripting.dictionary")
.CompareMode = vbTextCompare
For Each Dn In Rng
    If Not .Exists(Dn.Value) Then
        .Add Dn.Value, Dn
    Else
        If nRng Is Nothing Then Set nRng = _
        Dn Else Set nRng = Union(nRng, Dn)
        i = 1
        k = Cells(19, 2).Value
        For i = 1 To k
            .Item(Dn.Value).Offset(, i + 5) = .Item(Dn.Value).Offset(, i + 5) + Dn.Offset(, i + 5)
        Next i
    End If
Next
If Not nRng Is Nothing Then nRng.EntireRow.Delete
End With

End Sub

Sub CleanUpBRRoom()
Dim aRange, bRange, cRange, dRange, eRange, fRange As Range

Worksheets("BK Info").Select
Set aRange = Range("A34:BC200")

aRange.Clear


End Sub