Dim objConn, objRst, objNetwork, strDay
'get the current name of day e.g. Monday
strDay = WeekDayName(Weekday(Date))
Set objNetwork = CreateObject("WScript.Network")
Set objConn = CreateObject("ADODB.Connection")

'open a connection using the ExcelUserData file DSN
objConn.Open "FileDSN=E:\Code Download\Chapter 13\ExcelUserData.dsn"

'get the record for todays date for the current user 
    Set objRst = objConn.Execute("Select " & strDay & _ 
       " From UserList Where UserID='" & objNetwork.UserName & "'")
    'check if the time column for the current day. If it is empty, then 
    'update the time.
    If IsNull(objRst(strDay).Value) Then
        objConn.Execute "UPDATE UserList Set " & strDay & "='" & Time & _ 
           "' Where UserID='" & objNetwork.UserName & "'"
    End If
objRst.Close
objConn.Close
