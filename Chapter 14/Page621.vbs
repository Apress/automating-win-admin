Const GROUP_TYPE_SECURITY_ENABLED = &h80000000
Dim objContainer, objGroup

Set objContainer = GetObject("LDAP://CN=Users,DC=Acme,DC=com")

'create the group 
Set objGroup = objContainer.Create("Group", "CN=Accounting Group")

'set the SAM account name for compatibility with existing NT and 
'Win9x clients
objGroup.samAccountName = "Acctusers"

objGroup.groupType = GROUP_TYPE_GLOBAL_GROUP Or GROUP_TYPE_SECURITY_ENABLED
objGroup.SetInfo
