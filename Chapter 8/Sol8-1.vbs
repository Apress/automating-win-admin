Dim objRegExp, strAddress

 'create a new instance of the regexp object
 Set objRegExp = New RegExp
 'set case matching off
 objRegExp.IgnoreCase = True
 'set pattern
 objRegExp.Pattern = "\w+(\.\w+)?@\w+(\.\w+)+" 

 strAddress= InputBox("Enter an E-mail address")
 'check if address is valid
 If Not objRegExp.Test(strAddress) Then 
   MsgBox "Not a valid E-mail address"
 End If
