Dim objConn, objRst, nTotal
Set objConn = CreateObject("ADODB.Connection")
'open a connection to a text file using a DSN
'the ODBC DSN is defined in Solution 13-1.
objConn.Open "SalesCSV"
'select all data from the text file Orders data.txt
Set objRst = objConn.Execute("Select * From [Orders data.txt]")

'loop through and total the contents
While Not objRst.EOF
    nTotal = nTotal + objRst("SalesTotal").Value
    objRst.MoveNext
Wend

Wscript.Echo "Total sales is:" & nTotal
'close the connection and recordset object
objRst.Close
objConn.Close
