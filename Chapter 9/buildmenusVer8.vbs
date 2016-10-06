'buildmenusVer8.vbs
'builds roll over menus using CorelDraw Version 8
Const Height = 18
Dim aMenus, nF, strPath, objCorel

 If Wscript.Arguments.Count <> 2 Then
   ShowUsage
   Wscript.Quit
 End If

  'get destination path and menu names
  strPath = Wscript.Arguments(0)
  aMenus = Split(Wscript.Arguments(1),";")

  Set objCorel = CreateObject("CorelPhotoPaint.Automation.8")

  'loop through and build menu elements 
  For nF = 0 To UBound(aMenus)
  'build 
   BuildElements CStr(aMenus(nF)), strPath  _
                & aMenus(nF) & "ON.jpg", 0, 0, 0
   BuildElements CStr(aMenus(nF)), strPath  _ 
                & aMenus(nF) & ".jpg", 255, 255, 255
  Next

Sub ShowUsage()
WScript.Echo _
    "buildmenus.vbs builds on/off images for Web rollovers ." _ 
     & vbCrLf & "Syntax:" &  vbCrLf & _
    "buildmenus.vbs Path Menus" &  vbCrLf & _
    "Path      path where images are stored" &  vbCrLf & _
    "Destination Titles for each button, separated by semi-colon" &  vbCrLf & _
    "Example:" & vbCrLf & " buildmenus.vbs d:\images Home;Shop;Help"
End Sub

Sub BuildElements(strText, strFileName, nRed, nGreen, nBlue)
 Dim nWidth, nHeight
 'calculate width of box
 nWidth = Int(7 * Len(strText))
With objCorel
   'create a new file with white background
  .FileNew nWidth + 2, HEIGHT, 1, 72, 72, False, _ 
           False, 1, 0, 0, 0, 0, 255, 255, 255, 0, False
  'draw a blue rectangle
  .RectangleTool 0, 0, 0, 0, True, False, True
  .FillSolid 5, 32, 102, 176, 0
  .Rectangle 0, 0, nWidth, HEIGHT
  'add centred text box 
  .TextTool nWidth / 2, 2, strText
  .SetPaintColor 5, nRed, nGreen, nBlue, 0
  .TextSettings 400, False, False, 1, "Arial", 14, True, 0, 100, 0, False
  'save file and close
  .FileSave strFileName, 774, 0
  .FileClose
 End With
End Sub
