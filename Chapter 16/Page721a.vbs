' get the container where the contact is to be stored
Set objContainer = GetObject("LDAP://ou=Contacts,dc=acme,dc=com")
'create the contact
Set objContact = objContainer.Create("contact", "CN=Joe SmithCX")

'set a few Active Directory contact object properties
objContact.givenName = "Joe"
objContact.sn = "Smith"
objContact.description = "Joe Smith, Salesperson Company X"
objContact.SetInfo
objContact.MailEnable "smtp:joes@companyx.com"
objContact.SetInfo
