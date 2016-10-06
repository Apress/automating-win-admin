Dim objFSO, objTextFile 

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile("D:\data.txt")
objTextFile.WriteLine "Write a line to a file with end of line character"
objTextFile.Write "Write string without new line character"
objTextFile.Close
