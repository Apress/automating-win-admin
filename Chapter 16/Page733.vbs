'get a reference to the group object you want to disable
   Set objGroup = GetObject("LDAP://cn=AGROUP,cn=Users,dc=c3i,dc=com")
    objGroup.MailDisable
    objGroup.SetInfo
