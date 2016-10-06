Const ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP = 4

Set objDomain = GetObject("WinNT://Acme")
Set objGroup = objDomain.Create("group", "Acctusers")
objGroup.groupType = ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP
objGroup.Description = "Accounting Users"
objGroup.SetInfo
