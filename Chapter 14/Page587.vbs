'bind to the container to add user to. In this example the acme domain.
Set objContainer = GetObject("LDAP://cn=Users,dc=acme,dc=com")

Set objUser = objContainer.Create("User", "cn=Fred Smith")
objUser.Put "samAccountName", "freds"
objUser.SetInfo

objUser.pwdLastSet = -1
objUser.SetPassword "we12oi90"
objUser.AccountDisabled = False

objUser.SetInfo
