'buildmenus.vbs
Const Height = 18
Dim aMenus, nF, strPath, objCorel

 If Wscript.Arguments.Count <> 2 Then
   ShowUsage
   Wscript.Quit
 End If

  'get destination path and menu names
  strPath = Wscript.Arguments(0)
  aMenus = Split(Wscript.Arguments(1),";")

  Set objCorel = CreateObject("CorelPhotoPaint.Automation.10")

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

Sub BuildElements(strText, strFileName, nRed, nGreen, nBlue) 'nRed, nGreen, nBlue
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
  'add text box 
  .TextTool 1, 12, FALSE, TRUE, 0
  .SetPaintColor 5, nRed, nGreen, nBlue, 0
  .TextSetting "Fill", nRed & "," & nGreen & "," & nBlue
  
  .TextSetting "Font", "Arial"
  .TextSetting "TypeSize", "12.000"
  .TextAppend strText
  .TextRender 

  'save file and close

 .FileSave strFileName, 774, 0
  .FileClose
 End With
End Sub

