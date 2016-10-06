Set objConn = CreateObject("ADODB.Connection")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=D:\data\Access\Samples\Northwind.mdb"
Set objRst = CreateObject("ADODB.Recordset")
'open the Products table and delete a record
    objRst.Open "Products", objConn, adOpenForwardOnly, adLockOptimistic
objRst.Delete
objRst.Close
