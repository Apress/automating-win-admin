Dim objDomain, objUser

'get a reference to a domain ojbect 
Set objDomain = GetObject("WinNT://Acme")
'filter on the user objects
objDomain.Filter = Array("user")
For Each objUser In objDomain
   Wscript.Echo objUser.Name
Next
