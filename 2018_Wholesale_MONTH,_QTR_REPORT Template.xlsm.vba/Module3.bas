Sub Step3ChinaAccount()
Dim aName, bName, cName As String, aRange, bRange, cRange, dRange, eRange As Range


'VMRH China Account
Worksheets("VMRH").Select

Range("A5").AutoFilter field:=3, Criteria1:=Array( _
"Guangdong PRC", "Beijing PRC", "Other Cities of China", "Shanghai PRC", "Shenzhen PRC", "China", "Guangzhou PRC"), _
Operator:=xlFilterValues

'Copy VMRH China Account Name
Set aRange = Range("B6").End(xlDown)
Set bRange = Range("B5").Offset(1, 0)
Set cRange = Range(bRange, aRange)

cRange.Copy Destination:=Worksheets("China figure (RN)").Range("A4")
cRange.Copy Destination:=Worksheets("China figure (RN Rev)").Range("A4")
Application.CutCopyMode = False

'Copy VMRH China RN and RN Rev
Set dRange = cRange.Offset(0, 12)
dRange.Copy Destination:=Worksheets("China figure (RN)").Range("C4")
Application.CutCopyMode = False

Set eRange = cRange.Offset(0, 14)
eRange.Copy Destination:=Worksheets("China figure (RN Rev)").Range("C4")
Application.CutCopyMode = False

Range("A5").AutoFilter

Set aRange = Nothing
Set bRange = Nothing
Set cRange = Nothing
Set dRange = Nothing
Set eRange = Nothing



'CMCC China Account
Worksheets("CMCC").Select

Range("A5").AutoFilter field:=3, Criteria1:=Array( _
"Guangdong PRC", "Beijing PRC", "Other Cities of China", "Shanghai PRC", "Shenzhen PRC", "China", "Guangzhou PRC"), _
Operator:=xlFilterValues

'Copy CMCC China Account Name
Set aRange = Range("B6").End(xlDown)
Set bRange = Range("B5").Offset(1, 0)
Set cRange = Range(bRange, aRange)

cRange.Copy Destination:=Worksheets("China figure (RN)").Range("D4")
cRange.Copy Destination:=Worksheets("China figure (RN Rev)").Range("D4")
Application.CutCopyMode = False

'Copy CMCC China RN and RN Rev
Set dRange = cRange.Offset(0, 12)
dRange.Copy Destination:=Worksheets("China figure (RN)").Range("F4")
Application.CutCopyMode = False

Set eRange = cRange.Offset(0, 14)
eRange.Copy Destination:=Worksheets("China figure (RN Rev)").Range("F4")
Application.CutCopyMode = False

Range("A5").AutoFilter

Set aRange = Nothing
Set bRange = Nothing
Set cRange = Nothing
Set dRange = Nothing
Set eRange = Nothing


'HICC China Account
Worksheets("HICC").Select

Range("A5").AutoFilter field:=3, Criteria1:=Array( _
"Guangdong PRC", "Beijing PRC", "Other Cities of China", "Shanghai PRC", "Shenzhen PRC", "China", "Guangzhou PRC"), _
Operator:=xlFilterValues

'Copy HICC China Account Name
Set aRange = Range("B6").End(xlDown)
Set bRange = Range("B5").Offset(1, 0)
Set cRange = Range(bRange, aRange)

cRange.Copy Destination:=Worksheets("China figure (RN)").Range("G4")
cRange.Copy Destination:=Worksheets("China figure (RN Rev)").Range("G4")
Application.CutCopyMode = False

'Copy HICC China RN and RN Rev
Set dRange = cRange.Offset(0, 12)
dRange.Copy Destination:=Worksheets("China figure (RN)").Range("I4")
Application.CutCopyMode = False

Set eRange = cRange.Offset(0, 14)
eRange.Copy Destination:=Worksheets("China figure (RN Rev)").Range("I4")
Application.CutCopyMode = False

Range("A5").AutoFilter

Set aRange = Nothing
Set bRange = Nothing
Set cRange = Nothing
Set dRange = Nothing
Set eRange = Nothing


'PARIS China Account
Worksheets("PARIS").Select

Range("A5").AutoFilter field:=3, Criteria1:=Array( _
"Guangdong PRC", "Beijing PRC", "Other Cities of China", "Shanghai PRC", "Shenzhen PRC", "China", "Guangzhou PRC"), _
Operator:=xlFilterValues

'Copy PARIS China Account Name
Set aRange = Range("B6").End(xlDown)
Set bRange = Range("B5").Offset(1, 0)
Set cRange = Range(bRange, aRange)

cRange.Copy Destination:=Worksheets("China figure (RN)").Range("J4")
cRange.Copy Destination:=Worksheets("China figure (RN Rev)").Range("J4")
Application.CutCopyMode = False

'Copy PARIS China RN and RN Rev
Set dRange = cRange.Offset(0, 12)
dRange.Copy Destination:=Worksheets("China figure (RN)").Range("L4")
Application.CutCopyMode = False

Set eRange = cRange.Offset(0, 14)
eRange.Copy Destination:=Worksheets("China figure (RN Rev)").Range("L4")
Application.CutCopyMode = False

Range("A5").AutoFilter

Set aRange = Nothing
Set bRange = Nothing
Set cRange = Nothing
Set dRange = Nothing
Set eRange = Nothing

End Sub

