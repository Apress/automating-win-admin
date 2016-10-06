'calc.vbs
Dim objArgs

Set objArgs = Wscript.Arguments

If objArgs.Count=1 Then
 Wscript.StdOut.WriteLine Eval(objArgs(0))
Else
 Wscript.StdErr.WriteLine "calc.vbs. Performs mathematical operations" & _
                          vbCrLf & "Syntax: " & vbCrLf & _
                         "calc.vbs expression " & vbCrLf & _
                         "expression   mathematical expression"
End If
