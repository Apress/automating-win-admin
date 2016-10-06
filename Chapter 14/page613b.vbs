'get a user object
Set objUser = GetObject("LDAP://cn=Fred Smith,cn=Users,DC=Acme,DC=COM")
'create a Byte array conversion object
Set objBAC = CreateObject("BAC.Convert")

'get the login hours using the BAC object
obj = objBAC.ByteToVariant(objUser.LoginHours)

'set all hours on
For nF = LBound(obj) To UBound(obj)
 Debug.Print obj(nF)
  obj(nF) = 255
Next

objUser.LoginHours = objBAC.VariantToByte(obj)
objUser.SetInfo
