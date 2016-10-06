'create an instance of the ENTWSH.FileSecurity object and set UseADSUtility property
Set objWS = CreateObject("ENTWSH.FileSecurity")
'list the permissions associated with a file
objWS.UseADSUtility = True
bSuccess = objWS.SetSecurity("c:\data\report.doc", "RXDO", "acme\Freds")
