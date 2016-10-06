'get a reference to the Fred Smith Active Directory user object
Set objUser = _
       GetObject("LDAP://cn= Fred Smith,cn=Users,dc=acme,dc=com")
'set the container name to Fred Smith
objUser.cn = "Fred Smith"
objUser.SetInfo
