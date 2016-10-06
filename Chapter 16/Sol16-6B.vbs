Dim objMember, objGroup
'get a reference to a group object list
Set objGroup = _
    GetObject("LDAP://cn=AG,cn=Users,dc=c3i,dc=com")
   objGroup.MailEnable
   objGroup.SetInfo
