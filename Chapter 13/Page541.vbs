Const adCmdText = 1
Dim objConn, objRst
Dim objConnDestination, objCopyData
Set objConnDestination = CreateObject("ADODB.Connection")
Set objConn = CreateObject("ADODB.Connection")
Set objRst = CreateObject("ADODB.Recordset")
'open the destination data file
objConnDestination.Open _
      "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=e:\Nwind.mdb;"

objConn.Open _
  "DRIVER={Microsoft Text Driver (*.txt; *.csv)};DBQ=e:\Code Download\Chapter 13;"

Set objRst = objConn.Execute("Select * From [Products.csv]")
Set objCopyData = CreateObject("WSHENT.CopyTable")

objCopyData.Source = objRst
'set the destination connection
objCopyData.Destination = objConnDestination 
objCopyData.Table = _
                             "Products" 'set the destination table
Call objCopyData.CopyTable

objRst.Close
objConn.Close
