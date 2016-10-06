'thumbnail.vbs
Const Height= 18 
Const JPEG = 774
Dim aMenus, nF, strPath, objCorel
Dim objFSO , objFolder, objFile, strNew
Dim strDestination, strSource


If WScript.Arguments.Count <> 2 Then
        ShowUsage
        WScript.Quit
End If

  'get destination path and menu names
  strSource =  Wscript.Arguments(0) 
  strDestination =  Trim(Wscript.Arguments(1))
'make sure destination path ends in a backslash
 If Not Right(strDestination,1) = "\" Then _
                strDestination = strDestination & "\"
Set objCorel = CreateObject("CorelPhotoPaint.Automation.10")
'get a reference to the source folder to read
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 Set objFolder = objFSO.GetFolder(strSource)
  For Each objFile In objFolder.files
   strNew = objFile.Name
   strNew = Left(strNew, InStr(strNew, ".") - 1) & "tm.jpg"
  CreateThumbnail objFile.Path, strDestination & strNew
 Next

Sub ShowUsage()
WScript.Echo _
    "thumbnail.vbs creates jpg. image thumbnails ." _ 
     & vbCrLf & "Syntax:" &  vbCrLf & _
    "thumbnail.vbs Source Destination" &  vbCrLf & _
    "Source      path to source directory with images" &  vbCrLf & _
    "Destination destination directory to store thumbnails" &  vbCrLf & _
     "Example:" & vbCrLf & " thumbnail.vbs d:\pictures d:\pictures\thumbs\"
End Sub
 
Sub CreateThumbnail(strSource, strDestination)

 With objCorel
    'arguments 2 to 5 represent left, top, right, bottom co-ordinates '
    'of image. Argument 6 represents load type, 7 and 8 are used
     'if movie file is being loaded and represents start and end frame.
    .FileOpen strSource, 0, 0, 0, 0, 0, 1, 1 
    'check if width is greater than height and resize accordingly 
    If objCorel.GetDocumentWidth < objCorel.GetDocumentHeight Then
       'arguments 1 and 2 repesent width and height. 3 and 4 are 
       'horizontal and vertical resolution in dots per inch and 
       'argument 5 is anti-aliasing flag, which if True sets 
       'anti-aliasing on
        .ImageResample 107, 143, 144, 144, True
    Else
        .ImageResample 144, 108, 144, 144, True
    End If
    'save resized. Second argument represents image format and third
     'is compression format used
    .FileSave strDestination, JPEG, 0
    .FileClose
 End With
End Sub
