Set objMailbox = GetObject("LDAP://cn=Fred Smith,cn=Users,dc=acme,dc=com") 
'set mailbox storage quota to 2 megabytes
objMailbox.mDBStorageQuota  = 2000
objMailbox.SetInfo
