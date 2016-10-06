Dim objRst
Const adOpenDynamic = 2
Const adLockPessimistic = 2
Const adCmdText = 1
Set objRst = CreateObject("ADODB.Recordset")
'get all products
objRst.Open "Select UnitPrice From Products", _
"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=d:\data\nwind\northwind.mdb", _
  adOpenDynamic, adLockPessimistic, adCmdText
'loop through and update all product prices by 2%
Do While Not objRst.EOF
    objRst("UnitPrice") = objRst("UnitPrice") * 1.02
    objRst.MoveNext
Loop
objRst.Close
