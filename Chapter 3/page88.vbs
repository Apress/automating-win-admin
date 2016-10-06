Dim objConn, objRS, objRoot, objDomain
Dim strFilter, objNetwork

Set objNetwork = CreateObject("WScript.Network")

'get logged on user name
 strUser = objNetwork.UserName

'get current domain
Set objRoot = GetObject("LDAP://rootDSE")
Set objDomain = GetObject("LDAP://" & _
        objRoot.Get("defaultNamingContext"))

'build a query to find user object's samAccountName property for logon name
strFilter = "(&(objectCategory=person)(objectClass=user)(samAccountName=" _
  & strUser & "))"

strQuery = "<" & objDomain.ADsPath & ">" _
        & ";" & strFilter & ";adsPath;subTree"

'connect to OLEDB Active directory provider
Set objConn = CreateObject("ADODB.Connection")

objConn.Open _
  "Data Source=Active Directory Provider;Provider=ADsDSOObject"
  
Set objRS = objConn.Execute(strQuery)

'if user object found display
If Not objRS.EOF Then
    Set objUser = GetObject(objRS("adsPath"))
    
Set objIE = CreateObject("InternetExplorer.Application")
    'go to the page
  
    objIE.Navigate objUser.wWWHomePage

    'wait to load page
    While objIE.Busy
    Wend
    objIE.Visible = True
End If
