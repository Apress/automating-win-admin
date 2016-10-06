Set objUser = GetObject("LDAP://cn=Fred Smith,cn=Users,DC=Acme,DC=COM")

Wscript.Echo "Last login " & objUser.LastLogin
Wscript.Echo "Account expires " & objUser.AccountExpirationDate
Wscript.Echo "Password last set " & _
               ConvToDate(objUser.pwdlastset)
Wscript.Echo "Last bad password attempt " & _
               ConvToDate(objUser.badPasswordTime)


Function ConvToDate(objTime)
 ConvToDate = #1/1/1601# + ((objTime.HighPart * 2 ^ 32) + _
                   objTime.LowPart) / 864000000000
End Function
