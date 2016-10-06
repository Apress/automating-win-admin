'set the workstations a user can log on to using the Active Directory provider
Dim objUser

'get a reference to a user object
Set objUser=GetObject("LDAP://CN=Fred Smith,OU=Accounting,DC=acme,DC=com")

objUser.UserWorkstations ="odin,thor,loki"
objUser.SetInfo 'update user settings
