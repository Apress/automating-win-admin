'get a reference to the Session's AddressLists collection
Set objAddressLists = objSession.AddressLists
'go through each AddressList and display the name for each of the 
'individual 
' AddressList objects also identify if the AddressList can be modified.
For Each objAddressList In objAddressLists
     WScript.Echo objAddressList.Name & " " & objAddressList.IsReadOnly
Next

'get a reference to the Personal Address Book Address List object
Set objAddressList = objAddressLists.Item("Personal Address Book")
