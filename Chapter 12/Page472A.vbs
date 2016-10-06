'add a new recipient to the message
Set objRecipient = objMessage.Recipients.Add

'set the display name to resolve from the Address book
objRecipient.Name = "Fred Smith"
objRecipient.Resolve ' resolve the name
'show the AddressEntry type of the resolved address.
WScript.Echo objRecipient.AddressEntry.Type & vbCrLf & objRecipient.Address
