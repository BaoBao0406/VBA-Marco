Sub Step1()
Dim d1, d2, d3, y1, y2, y3, y4, y5, DateRange, Text1, Text2 As String

'Get the Month value
d1 = DateAdd("m", -1, Now)
d2 = "Jan"
d3 = Format(d1, "mmm")

'Get Year value
y1 = DateAdd("yyyy", -1, Now)
y2 = Format(y1, "yy")
y3 = Format(Now, "yy")
y4 = Format(y1, "yyyy")
y5 = Format(Now, "yyyy")

If d3 = "Jan" Then
DateRange = d2
Else
DateRange = d2 & " - " & d3
End If

'Convert Month & Year Title in Table
Range("C3, F3, B47").Value = DateRange & " " & y2
Range("D3, G3, C47").Value = DateRange & " " & y3

'Convert Month & Year in Header and Footer
Text1 = "On the MICE side find below some aspects of the YoY comparison between the same period of booking created from "
Text2 = "Note that all leads from Taiwan & Korea with less than 30 rooms on peak except leisure group are not included in MICE Hotel Sales Report "

Range("A1").Value = Text1 & DateRange & " " & y4 & " vs " & DateRange & " " & y5 & ":"
Range("A1").Characters(111).Font.Color = vbRed
Range("A1").Characters(111).Font.Bold = True
Range("A1").Characters(111).Font.Italic = True
Range("A76").Value = Text2 & DateRange & " " & y5 & "."

End Sub

