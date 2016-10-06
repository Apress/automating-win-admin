Const adCmdStoredProc = 4
Set objCmd = CreateObject("ADODB.Command")
Set objConn = CreateObject("ADODB.Connection")
' open the pubs data source 
objConn.Open "Provider=SQLOLEDB.1;User ID=sa;Initial Catalog=pubs;Data Source=Odin"

'set the active connection
Set objCmd.ActiveConnection = objConn
'set the stored procedure to execute
objCmd.CommandText = "byroyalty"

'execute the stored procedure, passing parameters to it.
Set objRst = objCmd.Execute(,Array(50), adCmdStoredProc)

'loop through and display the results 
While Not objRst.Eof
  Wscript.Echo objRst(0)
  objRst.MoveNext
Wend

objConn.Close
