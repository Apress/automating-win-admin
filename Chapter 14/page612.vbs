'get user
Set objuser = GetObject("WinNT://acme/freds,user")

'get the login hours
obj = objuser.LoginHours
'display data type 
Wscript.Echo Vartype(obj)

'try to enumerate array, an error will occur:
For nF = Lbound (obj) to ubound (obj)
 Wscript.Echo obj(nF)
Next
