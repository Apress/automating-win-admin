'xlformat.vbs
Const xlSolid = 1
Const Red = 3
Const Yellow = 6
'create new instance of Excel application
Set objExcel = CreateObject("Excel.Application") 
With objExcel  
  .Visible = True
  'create a new Excel workbook
  .Workbooks.Add 
  'set the font color for range A4:F8 to red
  .Range("A4:F8").Font.ColorIndex = Red
  'set the fill of current selected cells to solid yellow
  With .Selection.Interior
    .ColorIndex = Yellow
    .Pattern = xlSolid   
  End With
End With
