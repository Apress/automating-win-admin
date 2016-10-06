'logonoff.vbs
Dim objConn, objNetwork, strCmp

   'create a ADO connection and open a Access database
   Set objConn = CreateObject("ADODB.Connection")
   objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=\\odin\data\logonoff.mdb"

   Set objNetwork = CreateObject("WScript.Network")
   strCmp = ""
   
Do While strCmp = ""
    strCmp = Trim(objNetwork.ComputerName) & ""
    Wscript.Sleep 10
Loop

'if user is logging off then update the last computer logon record
If Wscript.Arguments(0) = "Logoff" Then
   objConn.Execute "UPDATE tblUserLog SET tblUserLog.LogoffTime = Now() WHERE " & _
                    "tblUserLog.LogEntryID=DMax(""LogEntryID"",""tblUserLog"",""UserName = '" & _
                    "Administrator" & "' And LogOffTime Is Null"");"
                    
Else 'add new logon record
   objConn.Execute "INSERT INTO tblUserLog (UserName,LogonTime,LogonComputer) VALUES (""" _
                 & objNetwork.UserName & """,#" & Now & "#,""" & strCmp & """)", nF
End If

objConn.Close
