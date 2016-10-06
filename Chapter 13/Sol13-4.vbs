Dim objConn, objRst
Set objConn = CreateObject("ADODB.Connection")
'open a connection and provide a user id and password
objConn.Open "Provider=SQLOLEDB.1;Initial Catalog=pubs;Data Source=ODIN" _
    , "freds", "sderf"

'execute a query
Set objRst = objConn.Execute("Select Sum(ytd_sales) As TotalSales From titles")
'display the value
Wscript.Echo objRst("TotalSales").Value
objRst.Close
objConn.Close
