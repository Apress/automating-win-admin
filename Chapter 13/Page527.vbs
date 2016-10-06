'method one
Set objRecord = CreateObject("ADODB.Record")
objRecord.Open "default.htm", "URL=http://www.acme.com"
'method two, use existing Connection object
Set objConn = CreateObject("ADODB.Connection")
Set objRecord = CreateObject("ADODB.Record")

'open a connection to a Web server using Internet Publishing
'OLE DB provider
 objConn.Open "Provider=MSDAIPP.DSO;Data " & _
        "Source=http://www.acme.com;Mode=Read|Write" 
objRecord.Open "default.htm", objConnection
