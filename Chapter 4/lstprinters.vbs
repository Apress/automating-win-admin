'lstprinters.vbs
Set objNetwork = CreateObject("Wscript.Network")
Set objPrinters = objNetwork.EnumPrinterConnections

'loop through and display all connected printers..
For nF = 0 To objPrinters.Count - 1 Step 2
    Wscript.Echo  objPrinters(nF) & _
            " is connected to " & objPrinters(nF + 1)
    objNetwork.RemovePrinterConnection objPrinters(nF + 1)
Next
Set objNetwork = Nothing
