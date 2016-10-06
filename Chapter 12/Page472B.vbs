'this will work, returns a reference to the first AddressEntry in 
'the AddressEntries collection
Set objAddressEntry = objAddressList.AddressEntries(1)
'this WILL NOT work, you cannot reference an AddressEntry object
'using the display name
Set objAddressEntry = objAddressList.AddressEntries("Fred Smith")
