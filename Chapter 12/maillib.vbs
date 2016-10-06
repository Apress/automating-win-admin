'Procedure GetFolderObj
'Description
'Returns a reference to a Folder object for the specified folder path
'the folder path is specified with the full Exchange folder path
'delimited with backslashes.
'Parameters  objSession reference to MAPI session object
'            sFolderSearch Folder path delimited with backslashes
'Returns     referece to Folder object if folder found. If folder not
'            found, returns Nothing
Function GetFolderObj(objSession, sFolderSearch)
Dim objFolder, objFolder2, objInfoStore, strFolder

On Error Resume Next

'get a reference to the Infostore object for the path
Set objInfoStore = objSession.Infostores.Item(StripPath(sFolderSearch))

'check if problem getting reference Infostore.
If Err Then
    Set GetFolderObj = Nothing
    Exit Function
End If

'get a reference to the root folder for the Infostore
Set objFolder = objInfoStore.RootFolder
'loop through path searching for the specified folder
Do While Len(sFolderSearch) > 0
    'get next folder in hierarchy
    strFolder = StripPath(sFolderSearch)
    
    Set objFolder2 = objFolder.Folders.Item(strFolder)
    'check if error - folder not found
    If Err Then
        Set GetFolderObj = Nothing
        Exit Function
    End If
    'this additional step must be taken due to bug in CDO 1.x
    For Each objFolder2 In objFolder.Folders
        If objFolder2.Name = strFolder Then
            Set objFolder = objFolder2
            Exit For
        End If
    Next
Loop

'return reference to folder
Set GetFolderObj = objFolder
End Function

'Procedure: StripPath
'Description
'Returns the next level from a folder path
'Parameters  sPath Folder path delimited with backslashes
'Returns     next level in path.
Function StripPath(sPath)
 Dim nF
 'look for the next level
 nF = InStr(sPath, "\")
 'if more levels in path, return name of level
 If nF > 0 Then
    StripPath = Left(sPath, nF - 1)
    sPath = Trim(Mid(sPath & " ", nF + 1))
 Else
    StripPath = Trim(sPath)
    sPath = ""
 End If
End Function


'Procedure: CreateFolder
'Description
'Creates a new message folder.
'Parameters  objSession reference to MAPI session object
'            sFolderSearch Folder path for new folder, delimited with 
'            backslashes 
'Returns     Reference to folder object if successful, otherwise Nothing.
Function CreateFolder(objSession, ByVal sFolderSearch)

Dim objfolder, objInfoStore, objfldr, sFindFolder 
On Error Resume Next
'get a reference to the Infostore object for the path
Set objInfoStore = objSession.InfoStores.Item(StripPath(sFolderSearch))
'check if problem getting reference Infostore.
If Err Then
    Set CreateFolder = Nothing
    Exit Function
End If

'get a reference to the root folder for the Infostore
Set objfolder = objInfoStore.RootFolder

'loop through path searching for the specified folder
Do While Len(sFolderSearch) > 0
    sFindFolder = StripPath(sFolderSearch)
    For Each objfldr In objfolder.Folders
        If UCase(objfldr.Name) = UCase(sFindFolder) Then
           Exit For
        End If
    Next
   
    If objfldr Is Nothing Then
        Set objfolder = objfolder.Folders.Add(sFindFolder)
    Else
        Set objfolder = objfldr
    End If
Loop
Set CreateFolder = objfolder
End Function

'Procedure GetAddressObj
'Description
'Returns a reference to a AddressEntry object for the specified
'recipient display name
'Parameters  objSession reference to MAPI session object
'            sAddressList Address list to search. E.g. Recipients
'            sAddress display name of recipient you are searching for
'Returns     reference to AddressEntry object if recipient is found. 
 'If recipient not found, returns Nothing
Function GetAddressObj(objSession, sAddressList, sAddress)
Dim objAddressList, objAddressEntry
On Error Resume Next
'get a reference to the specified address list
Set objAddressList = objSession.AddressLists(sAddressList)
'if error, then unable to get specified address list
If Err Then
    Set GetAddressObj = Nothing
    Exit Function
End If
Set GetAddressObj = Nothing
'loop through all addresses and search for name
For Each objAddressEntry In objAddressList.AddressEntries
  If UCase(sAddress) = Ucase(objAddressEntry.Name) Then
    Set GetAddressObj = objAddressEntry
    Exit For
  End If
Next
End Function
