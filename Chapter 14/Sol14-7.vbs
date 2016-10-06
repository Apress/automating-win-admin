Const ADS_PROPERTY_CLEAR = 1
Const ADS_PROPERTY_UPDATE = 2
Const ADS_PROPERTY_APPEND = 3
Const ADS_PROPERTY_DELETE = 4
Set objUser = GetObject("LDAP://CN=Freds,CN=Users,DC=Acme,DC=com")

'add new home phone numbers to the user
objUser.PutEx ADS_PROPERTY_APPEND, "OtherhomePhone", _ 
              Array("555-1234", "222-2222")
objUser.SetInfo
 
'deletes a fax number
objUser.PutEx ADS_PROPERTY_DELETE, "otherFacsimileTelephoneNumber", _ 
              Array("555-3453")
objUser.SetInfo

'updates (sets) the addtional mailbox addresses for the user
objUser.PutEx ADS_PROPERTY_UPDATE, "otherMailbox", _
              Array("freddy@acme.com", "fred@acme.com")
objUser.SetInfo

'clear mobile phone
objUser.PutEx ADS_PROPERTY_CLEAR, "OtherMobile", Null
objUser.SetInfo
