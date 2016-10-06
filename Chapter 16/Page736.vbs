  Dim objConn, objRst, aMailBoxes, nPos, objMailbox, nF
  Set objConn = CreateObject("ADODB.Connection")
  objConn.Provider = "ADsDSOObject"
  objConn.Open "Active Directory Provider"
  
  'execute a query against the Exchange server Odin, listing all mailboxes
  'where the department is Accounting
  Set objRst = _
    objConn.Execute("SELECT ADsPath FROM " & _
        "'LDAP://Odin" & _
        "' WHERE objectClass='organizationalPerson' AND department='Accounting'")
   
   'loop through all mailboxes and output the display name
   Do While Not objRst.EOF
    'get the mailbox object from directory using the object path
    Set objMailbox = GetObject(objRst("ADsPath"))
    aMailBoxes = objMailbox.otherMailbox
    
    'check if aMailBoxes returns an array of values
    If VarType(aMailBoxes) = 8204 Then     
 'loop through each E-mail address in array
        For nF = 0 To UBound(aMailBoxes)
      'check if E-mail address contains Acme.com, if so
      ' replace with accounting.acme.com
         
    If StrComp(Right(aMailBoxes(nF), 9),"@acme.com",vbTextCompare) = 0 Then
            nPos = InStr(aMailBoxes(nF), "@acme.com")
            aMailBoxes(nF) = Left(aMailBoxes(nF), nPos-1) & "accounting.acme.com"
            objMailbox.Put "otherMailbox", aMailBoxes
            objMailbox.SetInfo
          End If
      Next
    End If
    
    objRst.MoveNext
   Loop
   objRst.Close
   objConn.Close 
