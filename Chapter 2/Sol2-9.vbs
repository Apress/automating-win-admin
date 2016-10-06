Dim objShell
Dim strText, strFontSize, strFont, strFileName

strFont = "Arial"
strFontSize = 12

If Not Wscript.Arguments.Count = 2 Then
    WScript.Echo "mkbutton creates text image buttons" & vbCrLf & _
    "Syntax:"  &  vbCrLf & _
    "mkimage.vbs buttontext filename" &  vbCrLf & _
    "buttontext  Text forcreate button" &  vbCrLf & _
    "filename    File name for text button image" &  vbCrLf & _
    "Example:" & vbCrLf & _
    "mkbutton Home d:\data\images\homebutton"
  WScript.Quit
End If

strText = Wscript.Arguments(0)
strFileName = Wscript.Arguments(1)

Set objShell = CreateObject("WScript.Shell")

objShell.Run "Photodrw.exe"
Wscript.Sleep 100
objShell.AppActivate "Microsoft Photodraw"
Wscript.Sleep 1000 'wait for PhotoDraw to start

objShell.SendKeys "{ESC}^n" 'new document
Wscript.Sleep 100
objShell.SendKeys "^t" 'text mode
objShell.SendKeys strText
objShell.SendKeys "{tab}"
objShell.SendKeys strFont
objShell.SendKeys "~{tab}"
objShell.SendKeys strFontSize
objShell.SendKeys "{tab}"
objShell.SendKeys "%oes" 'Format menu - Effects - Shadow
objShell.SendKeys "{tab}{down 3}{right}~"

objShell.SendKeys "%oei" 'Format menu - Effects - Designer Text
objShell.SendKeys "{tab}" '{down 3}{right}"
objShell.SendKeys "~"

objShell.SendKeys "%vf" 'fit picture to selection
objShell.SendKeys "^s"
objShell.SendKeys strFileName
objShell.SendKeys "{tab}"
objShell.SendKeys "j~~~"
objShell.SendKeys "%fcn" 'close
objShell.SendKeys "%fx" 'quit
