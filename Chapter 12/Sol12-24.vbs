'get the Fields collection from a message
Set objFields = objMessage.Fields
'continue if error occurs - certain field types cannot be outputted and will 
'generate an error if attempted  to display
On Error Resume Next
'loop through all of the fields in the objFields collection
For Each objField In objFields
'display the Field value and ID.
     WScript.Echo objField.Values & " " & objField.ID
Next
