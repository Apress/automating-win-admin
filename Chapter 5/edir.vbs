'edir.vbs
  'enhanced directory utility. Lists all files that meet criteria from
  'specified directories
  Dim strCriteria, objFile, objFSO, objArgs 

  Set objArgs = Wscript.Arguments

  'check if less than 2 arguments are being passed
  If objArgs.Count < 2 Then
     ShowUsage
     Wscript.Quit -1
  End If

   Set objFSO = CreateObject("Scripting.FileSystemObject")

  If Not objFSO.FolderExists(objArgs(0)) Then
     Wscript.Echo "Path '" & objArgs(0) & "' not found "
	 Wscript.Quit -1
  End If

    Set objEvent = Wscript.CreateObject ("ENTWSH.RecurseDir","ev_")

    objEvent.Path = objArgs(0) 
    strCriteria = objArgs(1) 

     strCriteria = Replace(strCriteria, "Size","objFile.Size",1,-1,1)

     strCriteria = Replace(strCriteria, "Modified", _ 
                "objFile.DateLastModified",1,-1,1)

     strCriteria = Replace(strCriteria, "Accessed", _
                   "objFile.DateLastAccessed",1,-1,1)

strCriteria = Replace(strCriteria, "Created", _ 
                  "objFile.DateCreated",1,-1,1)

strCriteria = Replace(strCriteria, "Attributes", _ 
              "objFile.Attributes",1,-1,1)
  
   'check if third argument passed - this is regular expression file filter
   If objArgs.Count = 3 Then
    objEvent.Filter = objArgs(2)
   End If
   
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   objEvent.Process
Sub ShowUsage
    WScript.Echo "edir. Enhanced directory." & vbCrLf & _ 
        "Syntax:" &  vbCrLf & _
       "edir.vbs path criteria [filter] [/s]" &  vbCrLf & _
       "path     path to folder to search " & vbCrLf & _ 
       "criteria criteria to filter files on " & vbCrLf & _ 
       "filter   option regular expression filter "
 End Sub

 Sub ev_FoundFile(strPath)
    Set objFile = objFSO.GetFile(strPath)

   If Eval(strCriteria) Then
     Wscript.StdOut.WriteLine strPath
   End If
End Sub
