Const adOpenForwardOnly = 0
Const adLockOptimistic = 3
Dim objConn,objRst
Set objConn = CreateObject("ADODB.Connection")
Set objRst = CreateObject("ADODB.Recordset")
'open the datasource 
objConn.Open "FileDSN=ExcelData.dsn"
'open the named range LogData 
objRst.Open "LogData", objConn, adOpenForwardOnly, adLockOptimistic

'add a new row to the range
objRst.AddNew
objRst("LogTime") = Now
objRst("UserID") = "Fred Smith"
objRst("Description") = "An event has occurred"
'update and close data source
objRst.Update
objRst.Close
objConn.Close
