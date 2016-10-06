Dim objRst, strQuery, strContents, strKeyWord, strCriteria
Set objRst = CreateObject("ADODB.Recordset")
'check if there is either one or two parameters passed

If WScript.Arguments.Count = 0 Or WScript.Arguments.Count >2 Then 
  ShowUsage
  Wscript.Quit
End If 

'if only one argument has been passed, then set the search type to 'contents
If WScript.Arguments.Count = 1 Then
 strKeyWord = "Contents"
 strCriteria = WScript.Arguments(0)
Else
 strKeyWord = WScript.Arguments(0)
 strCriteria = WScript.Arguments(1)
End If

'build the query for the search engine. Replace all single quotes in the
'command line parameter with double quotes since the Index server wants
'double quotes.
 strQuery = "SELECT Path, DocLastSavedTm FROM Scope() WHERE CONTAINS(" _ 
               & strKeyWord & ",'" & Replace(strCriteria, "'", chr(34)) & "')>0"
'build the query for the search engine
  objRst.Open strQuery, "PROVIDER=MSIDXS"

'display each item
  While Not objRst.EOF
   Wscript.StdOut.WriteLine objRst(0) 
    objRst.MoveNext 
  Wend
   objRst.Close
   objRst.ActiveConnection.Close
   
Sub ShowUsage
 WScript.Echo "Idxsrv executes a query against Microsoft Index Server" & vbCrLf & _ 
     "Syntax:" &  vbCrLf & _
    "idxsrv [Type] IndexQuery" & vbCrLf & _
    "Type (optional) Server propery to search against. Default is Contents" & _ 
vbCrLf & "IndexQuery  query to execute"
End Sub
