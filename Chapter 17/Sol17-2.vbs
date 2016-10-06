'grant the user Freds Read/Execute permissions on the report.doc file
Const RIGHT_GENERIC_READ = 1179785
Const RIGHT_GENERIC_EXECUTE = 1179808
Const ACETYPE_ACCESS_ALLOWED = 0
Const ADS_PATH_FILE = 1
Const ADS_SD_FORMAT_IID = 1

Dim objDACL, objNewAce, objACE
Dim objSecurity, objSD, strFile

strFile = "c:\data\report.doc"

'if bADsSecurity is True, then use ADsSecurity object from
'ADSI 2.5 resource kit. If False, then use ADsSecurityUtility object
'available in Windows XP and 2003
bADsSecurity = True

If bADsSecurity Then
    Set objSecurity = CreateObject("ADsSecurity")
    Set objSD = objSecurity.GetSecurityDescriptor("FILE://" & strFile)
Else
    Set objSecurity = CreateObject("ADsSecurityUtility")
    Set objSD = objSecurity.GetSecurityDescriptor(strFile, _
                ADS_PATH_FILE, ADS_SD_FORMAT_IID)
End If

Set objDACL = objSD.DiscretionaryAcl 'get the Discretionary ACL DACL
Set objNewAce = CreateObject("AccessControlEntry")

'Set the properties for the ACE. Set the trustee to be the Freds account,
objNewAce.Trustee = "Acme\Freds"
'allow permissions to read and execute file
objNewAce.AccessMask = RIGHT_GENERIC_READ Or RIGHT_GENERIC_EXECUTE
'allow access to file
objNewAce.AceType = ACETYPE_ACCESS_ALLOWED

'add ACE to DACL
objDACL.AddAce objNewAce
'assign the DACL back to the security descriptor
objSD.DiscretionaryAcl = objDACL

'set the security descriptor for the file
If bADsSecurity Then
    objSecurity.SetSecurityDescriptor objSD
Else
    objSecurity.SetSecurityDescriptor strFile, _
                    ADS_PATH_FILE, objSD, ADS_SD_FORMAT_IID
End If
