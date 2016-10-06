Const ADS_PROPERTY_APPEND = 3
Const ADS_PROPERTY_DELETE = 4
Dim objMailbox
'get a reference to Freds mailbox
Set objMailbox = _ 
       GetObject("LDAP://odin/CN=FredS,CN=Recipients,OU=Office,o=Acme")

'add a new SMTP Internet address to a mailbox
objMailbox.PutEx ADS_PROPERTY_APPEND, "otherMailbox", _
            Array("smtp$freds@accounting.acme.com")
objMailbox.SetInfo

'delete a address from the mailbox
objMailbox.PutEx ADS_PROPERTY_DELETE, "otherMailbox", _
            Array("smtp$freds@finance.com ")
objMailBox.SetInfo
