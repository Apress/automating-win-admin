Const ADS_PROPERTY_CLEAR = 1
Set objUser = GetObject("LDAP://CN=fred smith,CN=Users,DC=Acme,DC=com")
'turn off static IP address
objUser.PutEx ADS_PROPERTY_CLEAR, "msRADIUSFramedIPAddress", 0
'turn off callback
objUser.PutEx ADS_PROPERTY_CLEAR, "msRADIUSServiceType", 0
objUser.SetInfo
