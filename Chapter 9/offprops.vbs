  'offprops.vbs
  'lists office documents that contain document properties that 
  'meet certain criteria
  Dim strCriteria,  bSubDirs
  Dim objRegExp, objEvent
  bSubDirs = False
  'check that two arguments are being passed
  If Wscript.Arguments.Count <> 2 Then
    ShowDetails
      Wscript.Quit
  End If
    'create a FSO and recursedir object (see Chatper 5,File Operations)
    Set objEvent = Wscript.CreateObject ("ENTWSH.RecurseDir","ev_")
    'set the path to search and get criteria
    objEvent.Path = Wscript.Arguments(0)
    strCriteria = Wscript.Arguments(1)
    'filter only on DOC, XLS and PPT documents
    objEvent.Filter = "^\w+\.(doc|xls|ppt)$" 
    'replace ` (ASCII 96) characters with double quotes
    strCriteria = Replace(strCriteria, "`", chr(34),1,-1,1)
    'replace all instances of document criteria doc.property with
    'objDoc.BuiltinDocumentProperties(property)
    Set objRegExp = New RegExp    
    objRegExp.Pattern = "\[\w+\]"
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    strCriteria = objRegExp.Replace(strCriteria, GetRef("Repl"))
    strCriteria = Replace(strCriteria, "_", " ",1,-1,1)
    Call objEvent.Process()

Sub ShowDetails
    WScript.Echo "offprops Queries office document properties." & vbCrLf & _ 
     "Syntax:" &  vbCrLf & _
    "offprops.vbs path criteria" &  vbCrLf & _
    "path      path to search" & vbCrLf & _ 
    "criteria  office property criteria " & vbCrLf & _ 
    "Example: List all documents authored by Fred Smith " & vbCrLf & _
    " offprops.vbs d:\data\word ""[Author]= `Fred Smith`"""
End Sub
'
Function Repl(strMatch, nPos, strSource)
  Repl = "objDoc.BuiltinDocumentProperties" & _ 
          "(""" & Mid(strMatch,2, len(strMatch)-2) & """)"
End Function

Sub ev_FoundFile(strPath)
   Dim objDoc, bResult
   On Error Resume Next 
   'get reference to document found
   Set objDoc = GetObject(strPath)
   bResult = Eval(strCriteria)
   If  bResult And Not Err Then
      If Not Err Then
       Wscript.StdOut.WriteLine strPath 
      Else
       Wscript.StdErr.WriteLine "Error opening file " & strPath _
                & vbCrLf & "Error:" & Err.Description  
      End If
   End If
End Sub
