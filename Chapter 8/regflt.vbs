'regflt.vbs
'command line regular expression filter
Dim nF, strDelim, strLine
 If WScript.Arguments.Count < 1 Or WScript.Arguments.Count > 2 Then
   ShowUsage
 End If 

 strDelim = "," 'set delimiter
 strPattern = WScript.Arguments(0) 
  
 'check if alternate output delimiter specified
 If WScript.Arguments.Count = 2 Then strDelim = WScript.Arguments(1)

 'create regular expression object and set properties
 Set objRegExp = New RegExp 
  
 objRegExp.Pattern = strPattern 
 objRegExp.IgnoreCase = True
 objRegExp.Global = True

  'loop until the end of the text stream has been encountered
  Do While Not WScript.StdIn.AtEndOfStream

  'read line from standard input 
   strLine = WScript.StdIn.ReadLine

  'execute regular expression match
  Set objMatches = objRegExp.Execute(strLine)
  strOut = ""
  'if matches are made, loop through each match and append to output
  If objMatches.Count > 0 Then
     For nF = 0 To objMatches(0).SubMatches.Count - 2
       Wscript.Stdout.Write objMatches(0).Submatches(nF) & strDelim  
     Next
    Wscript.Stdout.WriteLine objMatches(0).Submatches(nF) 
  End If
 Loop 

Sub ShowUsage() 
WScript.Echo "regflt filters standard input against " & _
     "a regular expression." & vbLf & _ 
     "Syntax:" &  vbLf & _
    "regflt.vbs regexp [delimiter]" &  vbLf & _
    "regexp     regular expression" &  vbLf & _
     "delimiter  optional. Character to delimiter output columns"
    WScript.Quit -1
End Sub
