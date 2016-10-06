'Exchange 2000 query
ExecuteQuery "SELECT cn FROM " & _
     "'LDAP://DC=acme,DC=com' WHERE" & _
    " department='Accounting' AND objectCategory='person'"

'Exchange 5.5 query
ExecuteQuery "SELECT cn,TelephoneNumber  FROM " & _
     "'LDAP://Odin' WHERE objectClass='organizationalPerson'" & _
    " AND department='Accounting'"

Sub ExecuteQuery(strQuery)
  Dim objConn, objRst
  Set objConn = CreateObject("ADODB.Connection")
  objConn.Provider = "ADsDSOObject"
  objConn.Open "Active Directory Provider"
  
  'execute a query against the Exchange server Odin, listing all mailboxes
  'where the department is Accounting
  Set objRst = _
    objConn.Execute(strQuery)
    'loop through all mailboxes and output the display name
   Do While Not objRst.EOF
    Wscript.Echo objRst("cn")
    objRst.MoveNext
   Loop

   objRst.Close
   objConn.Close
End Sub
