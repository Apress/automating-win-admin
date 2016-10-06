Dim objRst, objConn
Set objConn = CreateObject("ADODB.Connection")
Set objRst = CreateObject("ADODB.Recordset")
'open a connection to a Web server using Internet Publishing
'OLE DB provider
 objConn.Open "Provider=MSDAIPP.DSO;Data " & _
        "Source=http://odin/data;Mode=Read|Write;" & _
        "User ID=Administrator;Password=we56oi90"

'list all files from the folder
objRst.Open "*", objConn
'loop through all files
While Not objRst.EOF
'check if the size of file is a numeric value – indicates
'a file
If Not IsNull(objRst("RESOURCE_STREAMSIZE")) Then
'checks if file is older than 30 days and if it is a htm file    
If DateDiff("d", objRst("RESOURCE_LASTWRITETIME"), Date) < 30 _ 
   And Right(objRst("RESOURCE_DISPLAYNAME"), 3) = "htm" Then
         objRst.Delete
       End If
       
    End If
    objRst.MoveNext
Wend
objRst.Close
objConn.Close
