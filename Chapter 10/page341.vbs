Dim objService, objVolumes, objVolume, nError, bDefragRecommended, objDA

'get the defrag status on the C: for the local computer
Set objService = GetObject("winmgmts:root\cimv2")

'get a reference to the C: drive
Set objVolumes = objService.ExecQuery("Select * from Win32_Volume Where Name = 'C:\\'")

'loop through the collection and process the single returned drive
For Each objVolume In objVolumes
    nError = objVolume.DefragAnalysis(bDefragRecommended, objDA)
    If Not nError Then
         Wscript.Echo objDA.FilePercentFragmentation & "% of files are fragmented."
       
        If bDefragRecommended Then
           Wscript.Echo "This volume should be defragged."
        Else
           Wscript.Echo "This volume does not need to be defragged."
        End If
    End If
Next
