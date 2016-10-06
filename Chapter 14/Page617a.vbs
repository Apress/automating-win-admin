Const UF_PASSWD_CANT_CHANGE = 64
Const UF_DONTEXPIRE_PASSWD = 65536

Dim objUser
Set objUser = GetObject("WinNT://ACME/fsmith,user")
  'turn of the UF_PASSWD_CANT_CHANGE flag
   objUser.userFlags = objUser.userFlags And Not UF_PASSWD_CANT_CHANGE
  'toggle the UF_DONTEXPIRE_PASSWD flag
   objUser.userFlags = objUser.userFlags Xor UF_DONTEXPIRE_PASSWD
objUser.SetInfo
