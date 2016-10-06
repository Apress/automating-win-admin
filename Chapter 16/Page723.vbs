Const ADS_PROPERTY_APPEND = 3
Dim objContainer, objMailbox
'get a reference to the Recipients container.
Set objContainer = GetObject("LDAP://odin/CN=Recipients,OU=Office,o=Acme")

'set a filter on organizationalPerson objects – this will only return
'mailboxes in the conter
 objContainer.Filter = Array("organizationalPerson")

'loop through each mailbox
 For Each objMailbox In objContainer
    'add a new Internet address to the mailbox, using the mailboxes 
    'alias as the Internet address name
    objMailbox.PutEx ADS_PROPERTY_APPEND, "otherMailbox", _
            Array("smtp$" & objMailbox.uId & "@acmeus.com")
       objMailbox.SetInfo
 Next
