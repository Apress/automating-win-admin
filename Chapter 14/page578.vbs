'get a reference to the rootDSE
Set objRootDSE = GetObject("LDAP://RootDSE")
'get the domain
strDomain = objRootDSE.Get("defaultNamingContext")
'get a reference to the domain object
Set objDomain = GetObject("LDAP://" & strDomain).
