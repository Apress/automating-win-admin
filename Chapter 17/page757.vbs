'remove user Freds from file data.doc
Set objSecurity = CreateObject("ADsSecurity")
Set objSD = objSecurity.GetSecurityDescriptor("FILE://d:\data\data.doc")
Set objDACL = objSD.DiscretionaryAcl For Each objACE In objDACL
  If objACE.Trustee = "Acme\FredS" Then
     objDACL.RemoveAce objACE
  End If
Next
 objSD.DiscretionaryAcl = objDACL
 'set the security descriptor for the file
 objSecurity.SetSecurityDescriptor objSD
