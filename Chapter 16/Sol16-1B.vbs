'create a Exchange 2000 mailbox 
Dim strServer, strDomain, strOrganization, strAdminGroup
Dim strStorageGroup, strStoreName
Dim objPerson, objMailbox
strServer = "Odin " 
strDomain = "acme.com" 
strOrganization = "acme" 
strAdminGroup = "First Administrative Group" 
strStorageGroup = "First Storage Group" 
strStoreName = "Mailbox Store (Odin)" 
  
'get a user object from Active Directory to create mailbox for
 Set objMailbox = GetObject("LDAP://cn=Fred Smith,cn=Users,dc=acme,dc=com")
 ' create mailbox for specified server
 objMailbox.CreateMailbox "LDAP://" & _
                    strServer & _
                   "/CN=" & _
                   strStoreName & _
                   ",CN=" & _
                   strStorageGroup & ",CN=InformationStore,CN=" & _
                   strServer & _
                   ",CN=Servers,CN=" & _
                    strAdminGroup & "," & _
                   "CN=Administrative Groups,CN=" & _
                    strOrganization & "," & _
                   "CN=Microsoft Exchange,CN=Services," & _
                   "CN=Configuration,dc=acme,dc=com"
 objMailbox.SetInfo 
