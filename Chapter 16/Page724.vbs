Const ADS_PROPERTY_APPEND = 3
Const ADS_PROPERTY_DELETE = 4
'get a reference to Fred's mailbox
Set objMailbox = _
      GetObject("LDAP://odin/cn=Freds,cn=Accountants,ou=Office,o=Acme")

'add a new SMTP Internet address to a mailbox
objMailbox.PutEx ADS_PROPERTY_APPEND, "proxyaddresses", _
            Array("SMTP:freds@accounting.acme.com")
objMailbox.SetInfo

'delete the current primary SMTP address from the mailbox
objMailbox.PutEx ADS_PROPERTY_DELETE, "proxyaddresses", _
            Array("SMTP:fred@acme.com")
objMailbox.SetInfo

'add the previously delete primary address back, but this time 
'as a secondary SMTP address
objMailbox.PutEx ADS_PROPERTY_APPEND, "proxyaddresses", _
            Array("smtp:fred@acme.com")
objMailbox.SetInfo
