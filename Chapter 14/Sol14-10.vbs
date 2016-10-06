'get a reference to the user 
Set objUser = GetObject("WinNT://Acme/freds")

 'list the machines the users have access to
For Each station In objUser.LoginWorkstations
  Wscript.Echo station
Next

'set the machines a user is permitted to logon to 
objUser.LoginWorkstations = Array("thor","odin","loki")
objUser.SetInfo
