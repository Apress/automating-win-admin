Set objUser = GetObject("LDAP://CN=fred smith,CN=Users,DC=Acme,DC=com")
SetUserCannotChangePassword objUser, False

Sub  SetUserCannotChangePassword(objUser, bChange)
Const CHANGE_PASSWORD_GUID = "{ab721a53-1e2f-11d0-9819-00aa0040529b}"
Const ADS_ACETYPE_ACCESS_DENIED_OBJECT = &H6
'get security descriptor and set trustees
Set objSD = objUser.Get("ntSecurityDescriptor")
Set objDACL = objSD.DiscretionaryAcl

For Each objACE In objDACL
  If objACE.ObjectType = CHANGE_PASSWORD_GUID Then
     If bChange Then
       objACE.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT
     Else
       objACE.AceType = 5
     End If
  End If
Next

objSD.DiscretionaryAcl = objDACL
objUser.Put "nTSecurityDescriptor", objSD
objUser.SetInfo

End Sub
