Set objDir = GetObject("IIS://odin/W3SVC/1/ROOT/images")

'get the IPSecurity
Set objIPSecurity = objDir.IPSecurity

'deny all access to directory  by default
objIPSecurity.GrantByDefault = False

'grant the address 10.5.5.1 and range of address 192.5.5.1 to 192.5.5.254
objIPSecurity.IPGrant = _
             Array("10.5.5.1, 255.255.255.255", "192.5.5.0, 255.255.255.0")

'grant all computers from domain acme.com
objIPSecurity.DomainGrant = _
         Array("acme.com")

objDir.IPSecurity = objIPSecurity
objDir.SetInfo
