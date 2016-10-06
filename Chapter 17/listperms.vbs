'listperms.vbs
'lists all permissions for a specified Active Directory object
Dim objSecurity, objArgs
Dim objContainer, objDACL , objACE 
If Wscript.Arguments.Count <> 1 Then
  Wscript.Echo "Usage: listperms ADPath"
End If

 Set objACE = CreateObject("AccessControlEntry")
 'get a reference to users container
 Set objContainer = GetObject(Wscript.Arguments (0)) 
  Set objSecurity = objContainer.Get("ntSecurityDescriptor")
Set objDACL = objSecurity.DiscretionaryAcl

 For Each objACE In objDACL
   Wscript.Echo objACE.Trustee & ", " & objACE.AceFlags & _ 
     ", " & objACE.AceType & ", " & objACE.Flags & ", " & _ 
     objACE.AccessMask  & ", " & objACE.ObjectType & ", " & _
     objACE.InheritedObjectType 
 Next
