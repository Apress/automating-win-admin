'createmailbox.vbs
Const ADS_RIGHT_EXCH_MODIFY_USER_ATT = 2
Const ADS_RIGHT_EXCH_MAIL_SEND_AS = 8
Const ADS_RIGHT_EXCH_MAIL_RECEIVE_AS = 16
Const ADS_SID_WINNT_PATH = 5
Const ADS_SID_HEXSTRING = 1
Dim objMailbox, objContainer, strServer
Dim strAlias, strMTA, strMDB, strSMTPAddr, strDisplayName
Dim objSid, strSidHex, strComputer, strSite, strOrg
Dim objComputer, strUserID, strDomain
Dim objSec, objSD, objDACL, objAce
strServer = "Odin"
strSite = "Office"
strDomain = "Acme"
strDisplayName = "Fred Smith"
strSMTPAddr = "FredSmith@acme.com"
strAlias = "Freds"
strUserID = "Freds"

Set objComputer = GetObject("LDAP://" & strServer)
'get the organization
strOrg = objComputer.o
'get the recipients container for the site
Set objContainer = _
           GetObject("LDAP://" & strServer & "/CN=Recipients,OU=" _
                   & strSite & ",o=" & strOrg)

'get the SID for the account to be associated with the new mailbox
Set objSid = CreateObject("ADsSID")
objSid.SetAs ADS_SID_WINNT_PATH, "WinNT://" & strDomain & "/" & strUserID
strSidHex = objSid.GetAs(ADS_SID_HEXSTRING)

'create a new MailBox
Set objMailbox = objContainer.create("organizationalPerson", "cn=" & strAlias) 
'set display name and alias
objMailbox.Put "mailPreferenceOption", 0
objMailbox.Put "cn", strDisplayName
objMailbox.Put "uid", strAlias
objMailbox.Put "Home-MTA", _
     "cn=Microsoft MTA,cn=" & strServer & ",cn=Servers,cn=Configuration,ou=" _
  & strSite & ",o=" & strOrg

objMailbox.Put "Home-MDB", _
    "cn=Microsoft Private MDB,cn=" & strServer & _
    ",cn=Servers,cn=Configuration,ou=" _
    & strSite & ",o=" & strOrg
objMailbox.Put "MAPI-Recipient", True
'objMailbox.Put "rfc822Mailbox", strSMTPAddr
objMailbox.rfc822Mailbox = strSMTPAddr

objMailbox.Put "textEncodedORaddress", _
        "c=US;a= ;p=" & strSite & ";o=" & strOrg & ";s=" & strAlias

objMailbox.textEncodedORaddress = _
            "c=US;a= ;p=" & strSite & ";o=" & strOrg & ";s=" & strAlias
objMailbox.Put "Assoc-NT-Account", strSidHex
objMailbox.SetInfo
'create security objects
Set objSec = CreateObject("ADsSecurity")
Set objAce = CreateObject("AccessControlEntry")

Set objSD = objSec.GetSecurityDescriptor("LDAP://" & strServer & _
                              "/CN=Recipients,OU=" & strSite & ",o=" & strOrg)
Set objDACL = objSD.DiscretionaryAcl
objAce.Trustee = strDomain & "\" & strUserID
objAce.AccessMask = ADS_RIGHT_EXCH_MODIFY_USER_ATT Or _
                                   ADS_RIGHT_EXCH_MAIL_SEND_AS Or _
                                   ADS_RIGHT_EXCH_MAIL_RECEIVE_AS
objAce.AceType = ADS_ACETYPE_ACCESS_ALLOWED
objDACL.AddAce objAce
objSD.DiscretionaryAcl = objDACL
objSec.SetSecurityDescriptor objSD

