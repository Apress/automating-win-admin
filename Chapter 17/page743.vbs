Dim objShare, objDescriptor, objACE, retval

Const WMIMoniker = "winmgmts:{impersonationLevel=impersonate}!”
'get a reference to the data share
Set objShare = _
    GetObject(WMIMoniker & "Win32_LogicalShareSecuritySetting.Name='Data'”)

retval = objShare.GetSecurityDescriptor(objDescriptor)
'loop through each ACE in the DACL and output the trustee's access
For Each objACE In objDescriptor.dacl
 'list the trustee (account) name and associated permission
 Wscript.Echo objACE.Trustee.Name, objACE.AccessMask
Next
