Const UF_DONTEXPIRE_PASSWD = 65536  
Set objUser = GetObject("WinNT://ACME/fsmith,user")
objUser.Put "userFlags", usr.Get("UserFlags") Or ADS_UF_DONTEXPIRE_PASSWD
objUser.SetInfo
