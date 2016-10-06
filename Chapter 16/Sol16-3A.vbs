Set objMailbox = _
      GetObject("LDAP://odin/cn=Freds,cn=Accountants,ou=Office,o=Acme")
objMailBox.Put "MDB-Use-Defaults", False
'set the mailbox to send warnings at 10 megabytes
objMailBox.Put "MDB-Storage-Quota", 10000
objMailBox.SetInfo
