'qadna.vbs
'Quick And Dirty Network Administration interface
Const Domain = "Acme" 
Const ForAppending = 8
Const DataFile= "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\odin\Data\qadna.mdb"
Const ConnectShare = 5 
Const ConnectPrinter = 4
Const UpdateRegistry = 100

Dim objConn, objRst, objNetwork
Dim strGroups, strUser, objGroup
Dim strQuery, strComputer, objShell, strLogFile

Set objShell = CreateObject("Wscript.Shell")

Set objNetwork = CreateObject("Wscript.Network")

'set the log file name - ensure uniquness by combining
'user id and date and time
strLogFile = "\\thor\e$\" & strUser & " " &  Month(date) & "-" & _
              Day(date) & "-" & Year(date) & " " & _
              Hour(time) & "_" & Minute(time) & "_" _
              & Second(time) & ".txt"

'get name of logged on user - loop until user id is retrieved
'required for Win9x
Do While strUser = ""
 strUser = objNetwork.UserName
 Wscript.Sleep 100
Loop

strComputer = objNetwork.ComputerName

'On Error Resume Next

'get ADSI User object from domain
Set objUser = GetObject("WinNT://" & Domain & "/" & strUser & ",User")

If Err Then 
  LogIt strLogFile,"Error getting ADSI user object for " & strUser & _
        vbCr & Err.Description & " " & Err
  Wscript.Quit -1
End If


'build query to execute against QADNA database
strQuery = "Select * From qryLinkAccountsWithActions WHERE " & _
           "(AccountName In (" 

'enumerate all user groups
For Each objGroup In objUser.Groups
  strQuery = strQuery & "'" & objGroup.Name & "',"

Next
Wscript.Echo strQuery
'add user name to query
strQuery = Left(strQuery, Len(strQuery) - 1) & _
           ")) Or AccountName ='" & strUser & "'"

'create ADO object and open QADNA database
Set objConn = CreateObject("ADODB.Connection")
objConn.Open DataFile

If Err Then 
  LogIt strLogFile,"Error opening database " & DataFile & _
        " on computer " & strComputer & vbCrLf & _
        Err.Description & " " & Err
  Wscript.Quit -1
End If

Set objRst = objConn.Execute(strQuery)

 'loop through each record and perform specified operation
 Do While Not objRst.Eof
   Select Case objRst("ActionType")
     Case ConnectShare
        'remove any existing connected drive
        objNetwork.RemoveNetworkDrive objRst("DriveLetter") & ":", True
        'clear any errors ocurred, such as removing non-connected drive
        Err.Clear 
       'connect drive   
        strPath = "\\" & objRst("ObjectSource") & "\" _
                  & objRst("ObjectName")
        objNetwork.MapNetworkDrive objRst("DriveLetter") & ":", _
                                   strPath, True

       If Err Then 
        LogIt strLogFile,"Error connecting to  " & strPath & _
              " on computer " & strComputer & vbCrLf  & _
              Err.Description & " " & Err

       End If

     Case ConnectPrinter 'value 4 - add printer
        Err.Clear 
        strPath = "\\" & objRst("ObjectSource") & "\" & _
                   objRst("ObjectName")
        objNetwork.AddWindowsPrinterConnection strPath

       If Err Then 
        LogIt strLogFile,"Error adding printer " & strPath & _
              " on computer " & strComputer & vbCrLf & _
              Err.Description & " " & Err
       End If

     Case UpdateRegistry 

         'write registy value
        strPath = objRst("Path1")
        'check if DWORD type value
     
     If objRst("DriveLetter") = "D" Then
         objShell.RegWrite strPath, objRst("Path2"),"REG_DWORD"
        Else
         objShell.RegWrite strPath, objRst("Path2")
        End If

       If Err Then
 
         LogIt strLogFile,"Error setting registry key  " & strPath & _
              " on computer " & strComputer & vbCrLf & _
              Err.Description & " " & Err
       End If

   End Select
   objRst.MoveNext
 Loop 

objRst.Close
objConn.Close
