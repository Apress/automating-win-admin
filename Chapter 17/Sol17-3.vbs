Const OBJECT_INHERIT_ACE = 1
Const CONTAINER_INHERIT_ACE = 2
Const INHERIT_ONLY_ACE = 8
Const ACETYPE_ACCESS_ALLOWED = 0
Const ADS_PATH_FILE = 1
Const ADS_SD_FORMAT_IID = 1

'file access types
Const FILE_GENERIC_READ = &H120089
Const FILE_GENERIC_WRITE = &H120116
Const FILE_GENERIC_EXECUTE = &H1200A0
Dim objDACL, objNewAce, objACE, objNewAce2
Dim objSecurity, objSD, strTrustee, bADsSecurity

'if bADsSecurity is True, then use ADsSecurity object from
'ADSI 2.5 resource kit. If False, then use ADsSecurityUtility object
'available in Windows XP and 2003
bADsSecurity = False

If bADsSecurity Then
    Set objSecurity = CreateObject("ADsSecurity")
    Set objSD = objSecurity.GetSecurityDescriptor("FILE://c:\data")
Else
    Set objSecurity = CreateObject("ADsSecurityUtility")
    Set objSD = objSecurity.GetSecurityDescriptor("c:\data", _
                ADS_PATH_FILE, ADS_SD_FORMAT_IID)
End If

Set objDACL = objSD.DiscretionaryAcl
Set objNewAce = CreateObject("AccessControlEntry")
strTrustee = "Acme\Freds"
'set file access to directory so any FredS can
'read existing file in the directory.
Set objNewAce = CreateObject("AccessControlEntry")
objNewAce.Trustee = strTrustee
objNewAce.AccessMask = FILE_GENERIC_READ Or FILE_GENERIC_EXECUTE
objNewAce.AceType = ACETYPE_ACCESS_ALLOWED
'permissions are to be inherited to any new files in the directory
objNewAce.AceFlags = INHERIT_ONLY_ACE Or OBJECT_INHERIT_ACE
objDACL.AddAce objNewAce

'set directory permissions so FredS can add files
Set objNewAce2 = CreateObject("AccessControlEntry")
objNewAce2.Trustee = strTrustee
objNewAce2.AccessMask = FILE_GENERIC_READ Or _
                        FILE_GENERIC_EXECUTE Or FILE_GENERIC_WRITE
objNewAce2.AceType = ACETYPE_ACCESS_ALLOWED
objNewAce2.AceFlags = CONTAINER_INHERIT_ACE
objDACL.AddAce objNewAce2
objSD.DiscretionaryAcl = objDACL

If bADsSecurity Then
    objSecurity.SetSecurityDescriptor objSD
Else
    objSecurity.SetSecurityDescriptor "c:\data", _
                    ADS_PATH_FILE, objSD, ADS_SD_FORMAT_IID
End If
