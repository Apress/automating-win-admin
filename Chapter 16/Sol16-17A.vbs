Dim objMember, objDL
'get a reference to a distribution list
Set objDL = _
      GetObject("LDAP://odin/cn=acctusers,cn=Recipients,ou=Office,o=Acme")
For Each objMember In objDL.Members
     Wscript.Echo objMember.Name 
Next
