Dim objFSO, objTxtStrm, strLine
 Dim objRegExp, objIP, objDict 
 Dim strResolve, strKey
  
 'create dictionary, FSO and IPnetwork objects..
 Set objDict = CreateObject("Scripting.Dictionary")
 Set objIP = CreateObject("SScripting.IPNetwork")
 Set objFSO = CreateObject("Scripting.FileSystemObject")

 'create Regular expression
 Set objRegExp = CreateObject("Vbscript.RegExp")
'set pattern to validate ip address.. x.x.x.x
 objRegExp.Pattern = "(\d+(\.|\b)){4}"
 
 'open log file
 Set objTxtStrm = _
    objFSO.OpenTextFile("d:\winnt\system32\logfiles\w3svc1\ex990611.log")
  
 'loop through and process each line
 Do While Not objTxtStrm.AtEndOfStream
   strLine = objTxtStrm.ReadLine
   'test line against regular expression
   If objRegExp.test(strLine) Then
     'reverse lookup IP address in line
     strResolve = _
           objIP.DNSLookup(Mid(strLine, 10, InStr(10, strLine, " ") - 10))
    'if resolved to valid domain address add to dictionary
    If Not strResolve = "" Then
      'if already exists, increase count
      If objDict.Exists(strResolve) Then
        objDict(strResolve) = objDict(strResolve) + 1
      Else
        objDict.add strResolve, 1
      End If
    End If
   
   End If
 Loop
 
 'loop through and list domain name hit counts
 For Each strKey In objDict.Keys
    WScript.Echo strKey & "  " & objDict.item(strKey)
 Next
