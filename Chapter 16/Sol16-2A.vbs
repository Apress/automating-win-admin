'get a reference to the Freds mailbox 
Set objMailbox = _ 
       etObject("LDAP://odin/cn=Freds,cn=Recipients,ou=office,o=Acme")
'set the container name to Fred Smith
objMailBox.cn = "Fred Smith"
objMailBox.SetInfo
