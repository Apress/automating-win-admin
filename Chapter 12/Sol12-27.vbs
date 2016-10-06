Set objAddresslist = objSession.AddressLists("Personal Address Book")
Set objAddressEntries = objAddresslist.AddressEntries
'add a new address entry called New List. Make it a distribution list
Set objAddressDLEntry = objAddressEntries.Add("MAPIPDL", "New List", _
"dist@x.com")
objAddressDLEntry.Update 'update AddressEntry so changes are saved

'add user Fred Smith to distribution list
Set objAddressEntry = objAddressDLEntry.Members.Add("SMTP", "Fred Smith", _
"freds@x.com")
objAddressEntry.Update 'update AddressEntry so changes are saved

'add user Joe Blow to distribution list
Set objAddressEntry = objAddressDLEntry.Members.Add("SMTP", "Joe Blow", _
"joeb@x.com")
objAddressEntry.Update 'update AddressEntry so changes are saved

'add user Sally Jones to distribution list
Set objAddressEntry = objAddressDLEntry.Members.Add("SMTP", "Sally Jones", _
"sallyj@x.com")
objAddressEntry.Update 'update AddressEntry so changes are saved
