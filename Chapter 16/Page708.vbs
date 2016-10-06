'these strings are built from the Server field
strOrganization = "acme" 
strAdminGroup = "First Administrative Group" 
strServer = "Odin " ' Exchange server name
'these strings are built from the Mailbox Store field
strStorageGroup = "First Storage Group" ' storage group
strStoreName = "Mailbox Store (Odin)" 'mail store name

strDomain = "acme.com" 'strPath =     "LDAP://" & _
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
