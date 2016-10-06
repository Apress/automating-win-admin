Set objUser = GetObject("LDAP://CN=fred smith,CN=Users,DC=Acme,DC=com")
objUser.msRADIUSServiceType = 4 
objUser. msRADIUSCallbackNumber= "555-1245"
objUser.SetInfo
