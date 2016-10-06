'list all services on the computer Odin
'get a reference to a computer
Set objComputer = GetObject("WinNT://Odin")
'filter on the Service object class
objComputer.Filter = Array("Service")
'enumerate the services
   For Each objService In objComputer
    Wscript.Echo objService.Name, objService.DisplayName
   Next
