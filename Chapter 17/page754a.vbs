'create security object and get file security descriptor
'for Windows NT/2000 and AD resource kit ADSSecurity object
Set objSecurity = CreateObject("ADsSecurity")
Set objSD = objSecurity.GetSecurityDescriptor("FILE://d:\data\report.doc")
Set objDACL = objSD.DiscretionaryAcl
