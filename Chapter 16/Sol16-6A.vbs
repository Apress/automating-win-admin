Dim objContainer, objDL
Dim strDisplayName, strAlias, strSMTPAddr
Set objContainer = GetObject("LDAP://odin/CN=Recipients,OU=Office,o=Acme")
'create a new distribution list
Set objDL = objContainer.create("groupOfNames", "cn=Acctusers")

'Set distribution  properties
objDL.cn = "Accounting Users" 'display name
objDL.uid = "Acctusers" 'alas
objDL.mail = "Acctusers@acme.com" ' default SMTP address

'X.400 address
objDL.textEncodedORaddress = _
            "c=US;a= ;p=HeadOffice;o=Acme;s=Acctusers"
objDL.SetInfo
