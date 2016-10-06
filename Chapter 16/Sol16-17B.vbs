Dim objMember, objGroup
'get a reference to a group object list
Set objGroup = _
                  GetObject("LDAP://cn=A Group,cn=Users,dc=acme,dc=com")
'enumerate the group objects
For Each objMember In objGroup.Members
    If Not objMember.mailNickName = "" Then
     Wscript.Echo objMember.Name
    Else
     Wscript.Echo objMember.Name & " is not an Exchange enabled object"
    End If
Next
