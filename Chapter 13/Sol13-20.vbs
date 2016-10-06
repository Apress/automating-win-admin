Const adCmdStoredProc = 4 
Const adStateClosed = 0
Dim objCmd, objConn, objRst 
Set objCmd = CreateObject("ADODB.Command")
Set objConn = CreateObject("ADODB.Connection")

' open the pubs data source 
objConn.Open "Provider=SQLOLEDB.1;User ID=sa;Initial Catalog=pubs;Data Source=Odin"
objCmd.CommandText = "reptq3" 

Set objCmd.ActiveConnection = objConn
Set objRst = objCmd.Execute(,Array(0, 1000, "business"), adCmdStoredProc)

'loop through each Recordset
Do While objRst.State <> adStateClosed 
 Do While Not objRst.Eof
  Wscript.Echo objRst(0)
  objRst.MoveNext
 Loop
Set objRst = objRst.NextRecordset
Loop
objConn.Close
