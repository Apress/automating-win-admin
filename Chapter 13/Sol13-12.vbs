Const adCmdTable = 2
Const adLockOptimistic = 3
Const adOpenForwardOnly = 1
Dim objConn, objRst
Set objConn = CreateObject("ADODB.Connection")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;C:\data\northwind.mdb"
Set objRst = CreateObject("ADODB.Recordset")
'open the recordset
    objRst.Open "Products", objConn, _
    adOpenForwardOnly, adLockOptimistic, adCmdTable

'loop through all records
Do While Not objRst.EOF
    'only update if item is not discontinued
    If Not objRst("Discontinued")  Then
       objRst("UnitPrice") = objRst("UnitPrice") * 1.07
       objRst.Update
    End If
    objRst.MoveNext
Loop
