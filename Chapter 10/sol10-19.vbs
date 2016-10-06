Dim objWMIService, objVolumes, objVolume, bDefragRecommended
Dim strComputer, nError, objDefragAnalysis

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

'get all local fixed drives
Set objVolumes = objWMIService.ExecQuery _
    ("Select * from Win32_Volume Where DriveType=3")

'loop through each drive volume and only
For Each objVolume In objVolumes
    'check if drive needs to be defragged
    nError = objVolume.DefragAnalysis(bDefragRecommended _
            , objDefragAnalysis)
    

    'if defrag recommended then defrag drive
    If bDefragRecommended Then
        Wscript.Echo "Attempting to defrag " & objVolume.DriveLetter
        nError = objVolume.Defrag(True, objDefragAnalysis)
        
        'if no error then report success
        If Not nError Then
            Wscript.Echo "Drive " & objVolume.DriveLetter & " successfully defragged "
		Else
            Wscript.Echo "Error " & nError & " defragging " & objVolume.DriveLetter
        End If
    
    Else
        Wscript.Echo "Defrag not required on drive " & objVolume.DriveLetter 
    End If
Next
