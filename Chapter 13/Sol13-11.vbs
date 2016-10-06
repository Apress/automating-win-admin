Const adCmdTable = 2
Const adLockPessimistic = 2
Const adOpenForwardOnly = 1
Set objRst = CreateObject("ADODB.Recordset")
'open the data source
objRst.Open "Customers", _
 "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:\data\northwind.mdb", _
  adOpenForwardOnly, adLockPessimistic, adCmdTable
objRst.AddNew
objRst("CompanyName") = "Fred's Food Company"
objRst("CustomerID") = "MNOPQ"
objRst.Update
objRst.Close
