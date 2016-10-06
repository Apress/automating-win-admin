'xlupd.vbs
'updates values in specific range based on criteria
'create new instance of Excel application
Set objExcel = CreateObject("Excel.Application") 
  
With objExcel
 .Visible = True
 'load an existing spreadsheet
 .Workbooks.Open "C:\data.xls"
  Set objRange = .Range("Prices")
  'go through each cell in the range 
  For Each objCell In objRange
   'update cell value according to current value
   If objCell.Value<100 Then
    objCell.Value = objCell.Value * 1.04
   ElseIf objCell.Value< 200 Then
    objCell.Value = objCell.Value * 1.05  
   Else
    objCell.Value = objCell.Value * 1.07      
   End If
  Next
End With
