'list domain properties for LDAP and WinNT ADSI providers
Set objDom = GetObject("WinNT://Acme")
Set objDom2 = GetObject("LDAP://DC=Acme,DC=COM")

Wscript.Echo "Min password length " & objDom.MinPasswordLength & _
            " " & objDom2.minpwdlength
Wscript.Echo "Min password age " & objDom.MinPasswordAge & _
            " " & ConvToSeconds(objDom2.minPwdAge)
Wscript.Echo "Max password length " & objDom.MaxPasswordAge & _
            " " & ConvToSeconds(objDom2.maxPwdAge)
Wscript.Echo objDom.MaxBadPasswordsAllowed & _
            " " & objDom2.lockoutthreshold
Wscript.Echo objDom.PasswordHistoryLength & _
            " " & objDom2.pwdhistorylength
Wscript.Echo objDom.AutoUnlockInterval & _
            " " & ConvToSeconds(objDom2.lockoutDuration)
Wscript.Echo objDom.LockoutObservationInterval & _
            " " & ConvToSeconds(objDom2.LockoutObservationWindow)

Function ConvToSeconds(objTime)
Dim nLow

'get the lower 32 bits
nLow = objTime.LowPart
'if negative value add 2^32 to low value
If nLow < 0 Then
    nLow = nLow + 2 ^ 32
End If

ConvToSeconds = CDbl((objTime.HighPart * 2 ^ 32) + _
                    nLow) / CDbl(-10000000)
End Function
