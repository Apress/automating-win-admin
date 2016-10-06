Const FULL_ACCESS = 983551
Dim objSecurity 
Dim objContainer, objDACL , objACE 
Set objACE = CreateObject("AccessControlEntry")
'get a reference to the users container for the Acme domain
Set objContainer = GetObject("LDAP://cn=users,dc=acme,dc=com")

Set objSecurity = objContainer.Get("ntSecurityDescriptor")
Set objDACL = objSecurity.DiscretionaryAcl
objACE.Trustee = "Acme\Freds"
objACE.AccessMask = FULL_ACCESS

objDACL.AddAce objACE
objSecurity.DiscretionaryAcl = objDACL
'write security descriptor back to container
objContainer.Put "ntSecurityDescriptor", objSecurity
objContainer.SetInfo
