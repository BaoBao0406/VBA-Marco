
Private Sub CommandButton1_Click()
Dim aa As Integer



For aa = 1 To Sheets.Count
Sheets(aa).PageSetup.RightFooter = "&""Arial,regular""&10" & " " + Format(FormatDateTime(Now, 2), "dd-mmm-yyyy")

Next aa
End Sub