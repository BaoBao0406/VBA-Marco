Sub ConvertRoomBL()
Dim a, b, c, i, j, k, m, n, w, x, y, z As Integer, aRange, bRange, cRange, dRange As Range

Worksheets("Rm Table").Select

Cells(1, 1).EntireColumn.Merge across = False
a = Cells(Rows.Count, "A").End(xlUp).Row + 4

Cells(a, 1).Value = "***"


k = 1
Do Until Cells(k, 1) = "***"
    If Cells(k, 1).Value > 0 And Cells(k + 1, 1).Value = "" And Cells(k + 1, 2).Value <> "" Then
        Cells(k + 1, 1).Value = Cells(k, 1).Value
    End If
k = k + 1
Loop



x = 7
y = 7
z = 7
w = 7
For i = 1 To a
    'VMRH Room
    If Cells(i, 1) = "VMRH" Then

        n = 3
        j = i
        Set cRange = Cells(i, 1).End(xlToRight)
        Set dRange = cRange.Offset(0, 1)
        dRange.Value = "###"

        Do Until Cells(j, n).Value = "###"
            If Cells(j, n).Value = "Rooms" Then
                Set aRange = Cells(j, n).End(xlDown)
                Set bRange = aRange.Offset(1, 0)
                bRange.Value = "###"
            
                m = 1
                
                Do Until Cells(j + m, n).Value = "###"
                    If Cells(j + m, n).Value > 0 Then
                        Worksheets("VM Room").Cells(x, 1).Value = Cells(j + m, 1).Value
                        Worksheets("VM Room").Cells(x, 2).Value = Cells(j - 1, n).Value
                        Worksheets("VM Room").Cells(x, 3).Value = Cells(j + m, n).Value
                        Worksheets("VM Room").Cells(x, 12).Value = Cells(j + m, n + 1).Value
                        
                        x = x + 1
                    End If
                
                m = m + 1
                Loop
            End If
        n = n + 1
        Loop
    
    End If
    
    'PARIS Room
    If Cells(i, 1) = "PARIS" Then

        n = 3
        j = i
        Set cRange = Cells(i, 1).End(xlToRight)
        Set dRange = cRange.Offset(0, 1)
        dRange.Value = "###"

        Do Until Cells(j, n).Value = "###"
            If Cells(j, n).Value = "Rooms" Then
                Set aRange = Cells(j, n).End(xlDown)
                Set bRange = aRange.Offset(1, 0)
                bRange.Value = "###"
            
                m = 1
                
                Do Until Cells(j + m, n).Value = "###"
                    If Cells(j + m, n).Value > 0 Then
                        Worksheets("PA Room").Cells(y, 1).Value = Cells(j + m, 1).Value
                        Worksheets("PA Room").Cells(y, 2).Value = Cells(j - 1, n).Value
                        Worksheets("PA Room").Cells(y, 3).Value = Cells(j + m, n).Value
                        Worksheets("PA Room").Cells(y, 12).Value = Cells(j + m, n + 1).Value
                        
                        y = y + 1
                    End If
                
                m = m + 1
                Loop
            End If
        n = n + 1
        Loop
    End If
    
    
    'CMCC Room
    If Cells(i, 1) = "CMCC" Then

        n = 3
        j = i
        Set cRange = Cells(i, 1).End(xlToRight)
        Set dRange = cRange.Offset(0, 1)
        dRange.Value = "###"

        Do Until Cells(j, n).Value = "###"
            If Cells(j, n).Value = "Rooms" Then
                Set aRange = Cells(j, n).End(xlDown)
                Set bRange = aRange.Offset(1, 0)
                bRange.Value = "###"
            
                m = 1
                
                Do Until Cells(j + m, n).Value = "###"
                    If Cells(j + m, n).Value > 0 Then
                        Worksheets("CM Room").Cells(z, 1).Value = Cells(j + m, 1).Value
                        Worksheets("CM Room").Cells(z, 2).Value = Cells(j - 1, n).Value
                        Worksheets("CM Room").Cells(z, 3).Value = Cells(j + m, n).Value
                        Worksheets("CM Room").Cells(z, 12).Value = Cells(j + m, n + 1).Value
                        
                        z = z + 1
                    End If
                
                m = m + 1
                Loop
            End If
        n = n + 1
        Loop
    End If
    
    'HIMCC Room
    If Cells(i, 1) = "HIMCC" Then

        n = 3
        j = i
        Set cRange = Cells(i, 1).End(xlToRight)
        Set dRange = cRange.Offset(0, 1)
        dRange.Value = "###"

        Do Until Cells(j, n).Value = "###"
            If Cells(j, n).Value = "Rooms" Then
                Set aRange = Cells(j, n).End(xlDown)
                Set bRange = aRange.Offset(1, 0)
                bRange.Value = "###"
            
                m = 1
                
                Do Until Cells(j + m, n).Value = "###"
                    If Cells(j + m, n).Value > 0 Then
                        Worksheets("HI Room").Cells(w, 1).Value = Cells(j + m, 1).Value
                        Worksheets("HI Room").Cells(w, 2).Value = Cells(j - 1, n).Value
                        Worksheets("HI Room").Cells(w, 3).Value = Cells(j + m, n).Value
                        Worksheets("HI Room").Cells(w, 12).Value = Cells(j + m, n + 1).Value
                        
                        w = w + 1
                    End If
                
                m = m + 1
                Loop
            End If
        n = n + 1
        Loop
    
    End If
Next i

SortRoomBL

Worksheets("BK Info").Select

End Sub

Sub SortRoomBL()
Dim aRange, bRange As Range, k As Integer

'VM Room sorting
Worksheets("VM Room").Select
k = 7
Do Until Worksheets("VM Room").Cells(k, 1).Value = "***"
k = k + 1
Loop
k = k - 2

Set aRange = Cells(k, 12)
Set bRange = Range("A6", aRange)
With Worksheets("VM Room").Sort
.SortFields.Clear
.SortFields.Add Key:=Range("A6"), SortOn:=xlSortOnValues, Order:=xlAscending
.SetRange bRange
.Header = xlYes
.Apply
End With
Set aRange = Nothing
Set bRange = Nothing

'PA Room sorting
Worksheets("PA Room").Select
k = 7
Do Until Worksheets("PA Room").Cells(k, 1).Value = "***"
k = k + 1
Loop
k = k - 2

Set aRange = Cells(k, 12)
Set bRange = Range("A6", aRange)
With Worksheets("PA Room").Sort
.SortFields.Clear
.SortFields.Add Key:=Range("A6"), SortOn:=xlSortOnValues, Order:=xlAscending
.SetRange bRange
.Header = xlYes
.Apply
End With
Set aRange = Nothing
Set bRange = Nothing

'CM Room sorting
Worksheets("CM Room").Select
k = 7
Do Until Worksheets("CM Room").Cells(k, 1).Value = "***"
k = k + 1
Loop
k = k - 2

Set aRange = Cells(k, 12)
Set bRange = Range("A6", aRange)
With Worksheets("CM Room").Sort
.SortFields.Clear
.SortFields.Add Key:=Range("A6"), SortOn:=xlSortOnValues, Order:=xlAscending
.SetRange bRange
.Header = xlYes
.Apply
End With
Set aRange = Nothing
Set bRange = Nothing

'HI Room sorting
Worksheets("HI Room").Select
k = 7
Do Until Worksheets("HI Room").Cells(k, 1).Value = "***"
k = k + 1
Loop
k = k - 2

Set aRange = Cells(k, 12)
Set bRange = Range("A6", aRange)
With Worksheets("HI Room").Sort
.SortFields.Clear
.SortFields.Add Key:=Range("A6"), SortOn:=xlSortOnValues, Order:=xlAscending
.SetRange bRange
.Header = xlYes
.Apply
End With
Set aRange = Nothing
Set bRange = Nothing

End Sub