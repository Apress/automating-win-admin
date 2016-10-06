'list all connected users to an Access database
Const JET_SCHEMA_USERROSTER = _
                                      "{947bb102-5d43-11d1-bdbf-00c04fb92675}" 
Const adSchemaProviderSpecific = -1

Dim objConn,objRst,objField, nValue

Set objRst = CreateObject("ADODB.Recordset")
Set objConn = CreateObject("ADODB.Connection")

objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data.mdb "

'create a RecordSet using OpenSchema, returning the connected users. 
'OpenSchema queries the JET specfic property JET_SCHEMA_USERROSTER,
'which returns a list of connected users
Set objRst = objConn.OpenSchema(adSchemaProviderSpecific, , _
                                JET_SCHEMA_USERROSTER)

Do While Not objRst.EOF
 Wscript.Echo  objRst("COMPUTER_NAME") & " " & objRst("LOGIN_NAME")
 objRst.MoveNext
Loop
objRst.Close
objConn.Close
