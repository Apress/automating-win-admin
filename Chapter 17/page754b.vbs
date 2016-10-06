'create security object and get file security descriptor
'for Windows XP and later 
Const ACETYPE_ACCESS_ALLOWED = 0
Const ADS_PATH_FILE = 1
Const ADS_SD_FORMAT_IID =1
Set objSecurity = CreateObject("ADsSecurityUtility")
Set objSD = objSecurity.GetSecurityDescriptor("c:\data\report.doc", ADS_PATH_FILE, ADS_SD_FORMAT_IID)
Set objDACL = objSD.DiscretionaryAcl
