Dim objContainer, objMailbox 
'get a reference to the Recipients container. This is where the new custom
'recipient will be store
Set objContainer = GetObject("LDAP://odin/CN=Recipients,OU=Office,o=Acme")
'create an instance of a Remote-Address object. This is the class custom
'recipients are created 
Set objMailbox = objContainer.create("Remote-Address", "cn=Fred Smith")

objMailbox.Put "cn", "Fred Smith at Hotmail"
objMailbox.Put "uid", "FredsHM"
objMailbox.Put "Target-Address", "SMTP:freds@hotmail.com"
objMailbox.SetInfo
