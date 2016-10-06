Const YesButton = 6
Const QuestionMark = 32
Const YesNo = 4
'display a pop-up with yes/no buttons and question mark icon
Set objShell = CreateObject("WScript.Shell")
intValue = objShell.Popup("Do you wish to continue?", _
    , , QuestionMark + YesNo)
'test if the Yes button was selected
If intValue = YesButton Then
    'do something
End If
