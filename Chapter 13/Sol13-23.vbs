Const adCmdText = 1
Dim objConn 
Dim objDestConn, objRst, objCopy, dDate
Set objCopy = CreateObject("WSHENT.CopyTable")
Set objConn = CreateObject("ADODB.Connection")
Set objDestConn = CreateObject("ADODB.Connection")
Set objRst = CreateObject("ADODB.Recordset")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=d:\data\access\samples\Northwind.mdb;"

objDestConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=d:\data\access\samples\Northwind.mdb;"
dDate = #1/1/1996# 
'start transactions
objConn.BeginTrans
objDestConn.BeginTrans
objCopy.DESTINATION = objDestConn 'set the destination
'first get the order line items
Set objRst = objConn.Execute("SELECT [Order Details].* " & _
              "FROM Orders INNER JOIN [Order Details] ON " & _
              "Orders.OrderID =[Order Details].OrderID " & _
              "WHERE (ShippedDate<#" & dDate & "#)", , adCmdText)

objCopy.SOURCE = objRst
'the Order Details History table is a copy of the Order Details table structure
objCopy.Table = "[Order Details History]"

If Not objCopy.CopyTable() Then
    objConn.RollbackTrans
    objDestConn.RollbackTrans
End If

objConn.Execute "DELETE [Order Details].*, Orders.ShippedDate " & _
    "FROM Orders INNER JOIN [Order Details] ON Orders.OrderID = " & _
    "[Order Details].OrderID " & _
     "WHERE (((Orders.ShippedDate)<#" & dDate & "#));", , adCmdText

'get the details from order master
Set objRst = objConn.Execute("Select * From Orders Where ShippedDate<#" _
            & dDate & "#", , adCmdText)
objCopy.SOURCE = objRst
'the OrderHist table is a copy of the Order table structure
objCopy.Table = "OrderHist"

If Not objCopy.CopyTable() Then
    Wscript.Echo objCopy.Error
    objConn.RollbackTrans
    objDestConn.RollbackTrans
    Wscript.Quit
End If

Set objRst = objConn.Execute("Delete * From Orders Where ShippedDate<#" _
             & dDate & "#", , adCmdText)

objDestConn.CommitTrans
objConn.CommitTrans

objConn.Close
objDestConn.Close
