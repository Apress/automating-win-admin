'ieadmin.vbs
Dim objIE, objFSO, objDoc

'create an instance of the IE browser
Set objIE = Wscript.CreateObject("InternetExplorer.Application","ie_")

'objIE.FullScreen = True
'turn off all on screen 'clutter'
objIE.AddressBar= False
objIE.MenuBar= False
objIE.ToolBar= 0
objIE.Navigate "c:\Code Download\chapter 11\adminform.htm" '

 'wait to load page
 While objIE.Busy
  WScript.Sleep 100
 Wend

 Set objDoc = objIE.Document
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 Set objTS = objFSO.OpenTextFile("settings.ini",1) 

 Do While Not objTS.AtEndOfStream 
   strLine = objTS.ReadLine
   nPos = Instr(strLine,"=")
   If nPos > 0 Then
     Set objItem = objDoc.all(Trim(Left(strLine,nPos-1)))
      If objItem.type = "checkbox" Then
        objItem.checked = Cbool(Mid(strLine,nPos+1))
      Else
      'objDoc.all(Trim(Left(strLine,nPos-1))).value = Mid(strLine,nPos+1)
       objItem.value = Mid(strLine,nPos+1)
     End If 
     
   End If
 Loop

 Set objDOC.All("cmdCreateUser").onclick = GetRef("cmdCreateUser_OnClick")
 Set objDOC.All("cmdQuit").onclick = GetRef("Quit_OnQuit")
 Set objDOC.All("cmdSaveSettings").onclick = _
                                      GetRef("cmdSaveSettings_OnChange")

 Set objDOC.All("txtUserName").OnChange = GetRef("txtUserName_OnChange")
 Set objDOC.All("txtLastName").OnChange = GetRef("txtName_OnChange")
 Set objDOC.All("txtFirstName").OnChange = GetRef("txtName_OnChange")

 Set objDOC.All("txtDescription").OnChange = GetRef("txtDescription_OnChange")

 objIE.Visible = True

 bDone = False
  While Not bDone 
   wscript.sleep 100
  Wend

 Sub Quit_OnQuit
  objIE.Quit
  bDone = True
 End Sub

 'event fires when value in Full Name field is changed
 Sub txtUserName_OnChange
 
 objDoc.all("txtUserShare").value = objDoc.all("txtShareComputer").value _ 
                                & "\" & objDoc.all("txtUserName").value & "$"  

 End Sub

 'event fires when value in last or first Name fields is changed
 Sub txtName_OnChange
  objDoc.all("txtAccountName").value = objDoc.all("txtFirstName").value & " " _ 
                                        & objDoc.all("txtLastName").value
  objDoc.all("txtUserName").value = objDoc.all("txtFirstName").value  

  If Len(objDoc.all("txtLastName").value)> 1 Then
  		objDoc.all("txtUserName").value = objDoc.all("txtUserName").value _
		                    & Left(objDoc.all("txtLastName").value,1)
  End If
                                        
  txtUserName_OnChange
 End Sub



'event fires when value of Description field changes
 Sub txtDescription_OnChange
  'set Exchange Title field to value of account description
  objDoc.all("txtTitle").value = objDoc.all("txtDescription").value
 End Sub

'event fires when Create User button is clicked
 Sub cmdSaveSettings_OnChange

  'open settings file
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objTS = objFSO.OpenTextFile("settings.ini",2,True)

  'loop through all elements in the HTML document
  For Each objItem In objDOC.All

  'check if the element is an input element (input box, check box etc.)
  'and contains a value
  If TypeName(objItem) = "HTMLInputElement" Then 
    If (objItem.type = "text") _ 
         And objItem.value<>"" Then
     'write the value to the settings file
     objTS.WriteLine objItem.name & "=" & objItem.value
    ElseIf  objItem.type = "checkbox" Then
     objTS.WriteLine objItem.name & "=" & objItem.checked
     End If  
  End If
 Next
  objTS.Close
End Sub

 Sub cmdCreateUser_OnClick()
  Dim strLine, objD, objTS
  Dim strServer, strDomain, strOrganization, strAdminGroup
  Dim strStorageGroup, strStoreName
  Dim objPerson, objMailbox

  Set objD = objDOC.All
  Set objContainer = GetObject("LDAP://" & objD("txtContainer").value)
  Set objUser = objContainer.Create("User", "cn=" &  objD("txtAccountName").value)
  objUser.Put "samAccountName", objD("txtUserName").value

  objUser.SetInfo

  objUser.pwdLastSet = -1

  objUser.GivenName =  objD("txtFirstName").value 
  objUser.sn =  objD("txtLastName").value 
 
  If objD("txtDisplayName").value<>"" Then
    objUser.DisplayName =  objD("txtDisplayName").value 
  End If 
 
  objUser.AccountDisabled = False

  If objD("txtDescription").value<>"" Then  
   objUser.Description = objD("txtDescription").value
  End If

'set home directory?
If objD("chkHomeDirectory").checked Then
 objUser.HomeDrive = objD("txtHomeDirectory").value
 objUser.HomeDirectory = objD("txtUserShare").value
End If

 
 objUser.SetInfo

 objUser.SetPassword objD("txtPassword").value
 
  
 If objDOC.All("chkCreateExchangeAccount").checked Then

   strServer = objD("txtExchangeServer").value
   strDomain = objD("txtExchangeDomain").value
   strOrganization = objD("txtExchangeOrganization").value
   strAdminGroup = objD("txtAdminGroup").value

   strStorageGroup =  objD("txtStorageGroup").value
   strStoreName =  objD("txtStoreName").value

   strServer = objD("txtExchangeServer").value
   strDomain = objD("txtExchangeDomain").value
   strOrganization = objD("txtExchangeOrganization").value
   strAdminGroup = objD("txtAdminGroup").value


   Wscript.Echo objD("txtExchangeServer").value
   Wscript.Echo objD("txtExchangeDomain").value
   Wscript.Echo objD("txtExchangeOrganization").value
   Wscript.Echo objD("txtAdminGroup").value

   Wscript.Echo strStorageGroup 
   Wscript.Echo strStoreName

   'create mailbox for specified server
    objUser.CreateMailbox "LDAP://" & _
                   strServer & _
                   "/CN=" & _
                   strStoreName & _
                   ",CN=" & _
                   strStorageGroup & ",CN=InformationStore,CN=" & _
                   strServer & _
                   ",CN=Servers,CN=" & _
                    strAdminGroup & "," & _
                   "CN=Administrative Groups,CN=" & _
                    strOrganization & "," & _
                   "CN=Microsoft Exchange,CN=Services," & _
                   "CN=Configuration," & objD("txtServerDN").value

   objUser.SetInfo 
 End If
  
End Sub
