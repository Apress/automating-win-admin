'Exchangesec.vbs
'sets Admin access to the Recipients container
Option Explicit

Dim objSecurity, objSD, objDACL, objAce
Const ADS_ACETYPE_ACCESS_ALLOWED = 0

Const ADS_RIGHT_EXCH_ADD_CHILD = 1
Const ADS_RIGHT_EXCH_DELETE = 65536
Const ADS_RIGHT_EXCH_DS_REPLICATION = 64
Const ADS_RIGHT_EXCH_DS_SEARCH = 256
Const ADS_RIGHT_EXCH_MAIL_ADMIN_AS = 32
Const ADS_RIGHT_EXCH_MAIL_RECEIVE_AS = 16
Const ADS_RIGHT_EXCH_MAIL_SEND_AS = 8
Const ADS_RIGHT_EXCH_MODIFY_ADMIN_ATT = 4
Const ADS_RIGHT_EXCH_MODIFY_SEC_ATT = 128
Const ADS_RIGHT_EXCH_MODIFY_USER_ATT = 2

'create an instance of the ADsSecurity object
Set objSecurity = CreateObject("ADsSecurity")

'get the security descriptor for the object
Set objSD = objSecurity.GetSecurityDescriptor("LDAP://chaos/cn=Recipients,ou=c3i,o=c3i")

'get discretionary ACL for the object
Set objDACL = objSD.DiscretionaryAcl

'create an Access Control Entry (ACE)
Set objAce = CreateObject("AccessControlEntry")

'set the trustee
objAce.Trustee = "c3i\Ali"
objAce.AccessMask = ADS_RIGHT_EXCH_ADD_CHILD Or ADS_RIGHT_EXCH_DELETE Or _
    ADS_RIGHT_EXCH_MODIFY_ADMIN_ATT Or ADS_RIGHT_EXCH_MODIFY_USER_ATT Or _
    ADS_RIGHT_EXCH_MAIL_ADMIN_AS

'allow access
objAce.AceType = ADS_ACETYPE_ACCESS_ALLOWED

'add the ACE to the DACL
objDACL.AddAce objAce

objSD.DiscretionaryAcl = objDACL
objSecurity.SetSecurityDescriptor objSD

Set objSecurity = Nothing
Set objSD = Nothing
Set objDACL = Nothing
Set objAce = Nothing
