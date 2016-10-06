'get user Fred Smith
Set objUser = GetObject("LDAP://cn=Fred Smith,cn=Users,dc=acme,dc=com") 
'mail enable user with a Internet SMTP address
objUser.MailEnable "smtp:freds@hotmail.com"
'update settings
objUser.SetInfo
