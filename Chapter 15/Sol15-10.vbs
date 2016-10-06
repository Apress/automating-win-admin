Set objDir = GetObject("IIS://odin/W3SVC/1/ROOT/images")
'get the IPSecurity
Set objIPSecurity = objDir.IPSecurity

'allow all access to directory  by default
objIPSecurity.GrantByDefault = True
'deny the address 10.5.5.1 and range of address 192.5.5.x
objIPSecurity.IPDeny = _ 
             Array("10.5.5.1, 255.255.255.255", "192.5.5.0, 255.255.255.0")

'deny all computer from domain acme.com
objIPSecurity. DomainDeny= _ 
         Array("acme.com")

objDir.IPSecurity = objIPSecurity 
objDir.SetInfo
