'get a reference to a container to remove object from
Set objContainer = _
    GetObject("LDAP://odin/cn=Recipients,ou=Office,o=Acme"

'delete an object
objContainer.Delete "organizationalPerson", "cn=freds"
