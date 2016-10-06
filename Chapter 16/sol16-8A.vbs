Dim objContainer, objNewContainer, objNewContainer2
Set objContainer = GetObject(“LDAP://odin/OU=Office,O=Acme”)
'create a root recipients container
Set objNewContainer = objContainer.create(“Container”, “cn=External Users”)
objNewContainer.Put “Container-Info”, &H80000001
objNewContainer.SetInfo

'nest a recipients container inside the new External users container
Set objNewContainer2 = objNewContainer.create(“Container”, “cn=ABC Ltd”)
objNewContainer2.Put “Container-Info”, &H80000001
objNewContainer2.SetInfo
