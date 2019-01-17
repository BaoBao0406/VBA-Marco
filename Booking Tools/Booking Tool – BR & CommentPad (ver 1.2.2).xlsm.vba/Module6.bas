Sub Button()

Others
DeleteDuplicate

End Sub

Sub DeleteDuplicate()
Dim a, b As Integer

Worksheets("CommentPad").Select

a = 1
Do Until Worksheets("CommentPad").Cells(a, 1).Value = "Title: "
    b = 1
    If Cells(a, 1).Value <> "" And Cells(a, 1).Value = Cells(a + b, 1).Value Then
        Do Until Cells(a, 1).Value <> Cells(a + b, 1).Value
            Cells(a + b, 1).Value = "-----------"
        b = b + 1
        Loop

    End If
a = a + 1
Loop

End Sub

Sub ClearButtonCommentpad()

ClearVM
ClearPA
ClearCM
ClearHI
ClearEvents
ClearCommentPad
ClearEventTable

End Sub

Sub ConvertEventButton()

ConvertRental
ConvertFBmin
ConvertPkg

Worksheets("Events").Select

End Sub

Sub ClearButtonBKInfo()

ClearRoomBL
ClearEventTable
ClearBKInfo

End Sub