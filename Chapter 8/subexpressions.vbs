Dim objRegExp, objMatches, objSubMatches
'Set objRegExp = New RegExp
Set objRegExp = CreateObject("VBScript.RegExp")
' match "digits slash digits slash digits"
objRegExp.Pattern = "(\d+)\/(\d+)\/(\d+)"
Set objMatches = objRegExp.Execute("5/13/2000")
'the first element of the matches collection contains the submatches
Set objSubMatches = objMatches(0).SubMatches
'list the number of sub expression matches and matches
Wscript.Echo "Subexpression count " & objSubMatches.Count
Wscript.Echo "Subexpression 1 " & objSubMatches (0)
Wscript.Echo "Subexpression 2 " & objSubMatches (1)
Wscript.Echo "Subexpression 3 " & objSubMatches (2)
