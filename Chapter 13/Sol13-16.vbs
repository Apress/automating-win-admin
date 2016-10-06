Const adVarWChar = 202
Const adWchar = 130
Dim objConn, strDestinationFile
Dim objRst
Set objConn = CreateObject("ADODB.Connection")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=e:\nwind.mdb;"
    Set objRst = objConn.Execute("Select * From Products")
     CreateCSVFile "d:\output.txt", objRst, ","
    objRst.Close
    objConn.Close
Sub CreateCSVFile(strDestinationFile, objRst, strDelimiter)
Dim objField, strLine , objFileSystem, objTextFile
'create a file scripting object
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
'create the output file..
Set objTextFile = objFileSystem.CreateTextFile(strDestinationFile, True)

    'loop through each record in the recordset
    Do While Not objRst.EOF
    strLine = ""
    'loop through each field in the record, building the output string.
    For Each objField In objRst.Fields
            Select Case objField.Type
            Case adVarWChar, adWchar
             strLine = strLine & """" & objField.Value & """" & ","
             Case Else
                    strLine = strLine & objField.Value & ","
            End Select
    Next
            'write the line to the file
            objTextFile.WriteLine Left(strLine, Len(strLine) - 1)
    objRst.MoveNext
    Loop
    objTextFile.Close
End Sub
