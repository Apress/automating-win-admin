'extree.vbs
Set objShell = CreateObject("WScript.shell")

'execute the DOS tree command
Set objRun = objShell.Exec ("%Comspec% /c tree e:\")

'loop while application executes
  'build output text with results of tree command
  strText =  strText & objRun.StdOut.ReadAll


'output results
WScript.Echo strText
Set objShell = Nothing
