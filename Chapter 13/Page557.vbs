Set objCmd = CreateObject("ADODB.Command")
objCmd.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
                 "Data Source=e:\nwind.mdb"

objCmd.CommandText = "qryListProducts" 'set the query
'execute the command, passing an array of parameters to it.
Set objRst = objCmd.Execute(, Array("Beverages", "Exotic Liquids"))

While Not objRst.EOF
Wscript.Echo objRst("ProductName") & " " & objRst("Unitprice")
objRst.MoveNext
Wend
objRst.Close

