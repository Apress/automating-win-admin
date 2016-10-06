Set objCmd = CreateObject("ADODB.Command")
'check if script is run using Wscript or Cscript
   If StrComp(Right(Wscript.Fullname,11),"cscript.exe", vbTextCompare) <>0 Then 
Wscript.Echo "This script is best run using Cscript.exe"
    Wscript.Quit
   End If

objCmd.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
            "Data Source=D:\data\Access\Samples\Northwind.mdb"
objCmd.CommandText = "[Employee Sales by Country]" 'set the query

'execute the command, passing an array of parameters to it.
Set objRst = objCmd.Execute(, Array(#1/1/1996#, #12/12/1996#))

While Not objRst.EOF
Wscript.Echo objRst("ShippedDate") & " " & objRst("SaleAmount")
objRst.MoveNext
Wend

objRst.Close
