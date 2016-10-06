'get a user object
Set objUser = GetObject("WinNT://acme/freds,user")

'create a Byte array conversion object
Set objBAC = CreateObject("BAC.Convert")

'get the login hours using the BAC object 
obj = objBAC.ByteToVariant(objUser.LoginHours)

'enumerate array
For nF = Lbound (obj) to ubound (obj)
 Wscript.Echo obj(nF)
Next
