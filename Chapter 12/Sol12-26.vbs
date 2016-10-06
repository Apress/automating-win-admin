' create a session then log on
Set objSession = CreateObject("MAPI.Session")
' change the parameters to valid values for your configuration
objSession.Logon "Valid Profile"

'get a reference to the Recipient
Set objList = objSession.AddressLists("Personal Address Book")
Set objAddressEntries = objList.AddressEntries

'add an Internet mail address
   Set objAddressEntry = objAddressEntries.Add("SMTP", "Fred Smith", "freds@x.com")

objAddressEntry.Update