Const ADS_SCOPE_ONELEVEL = 1
'create a Connection object
Set objConn = CreateObject("ADODB.Connection")

objConn.Provider = "ADsDSOObject"
objConn.Open "Active Directory Provider"

'create a command object
Set objCmd = CreateObject("ADODB.Command")
Set objCmd.ActiveConnection = objConn

'search the one level only
objCmd.Properties("searchscope") = ADS_SCOPE_ONELEVEL objCmd.CommandText = _
   "SELECT cn FROM 'LDAP://OU=Sales,DC=Acme,DC=COM' WHERE objectClass='*'"
 
Set objRst = objCmd.Execute
While Not objRst.Eof 
  Wscript.Echo objRst("cn")
    objRst.MoveNext
Wend

objRst.Close
objConn.Close
