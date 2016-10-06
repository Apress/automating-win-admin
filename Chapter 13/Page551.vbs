Const DTSSQLStgFlag_UseTrustedConnection = 256
Dim objDTS
Set objDTS = CreateObject("DTS.Package")
'open the ProductImport package from the Odin server using NT 
'authentication
objDTS.LoadFromSQLServer "Odin", , , _ 
             DTSSQLStgFlag_UseTrustedConnection, , , , "ProductImport"

'set the data source for the text connection. 
objDTS.Connections("Connection 1").DataSource = "d:\importproducts.txt"

'enable writing of the completion status to event logs
objDTS.WriteCompletionStatusToNTEventLog = True

'set the supplier ID global variable to 4, which is the Tokyo Traders 'supplier
objDTS.GlobalVariables("SupplierID") = 4

objDTS.Execute
