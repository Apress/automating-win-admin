'create an instance of the ENTWSH.SysInfo object. This object 
'is created In Solution 10.1
'create an instance of the SysInfo object    
Set objSysInfo = CreateObject("ENTWSH.SysInfo")

On Error Resume Next
'attempt to get reference to running copy of Excel
Set objExcel = GetObject(,"Excel.Application")

'if no running copies of Excel, start a new one  
If Err Then Set objExcel = CreateObject("Excel.Application") 
  
With objExcel  
 .Application.Visible = True
 'create a Excel workbook based on the inventory template
 .Workbooks.Add "e:\Code Download\Chapter 9\inventory.xlt"
    
 .Application.Goto "BIOS"
 .ActiveCell.FormulaR1C1 = objSysInfo.BIOSVersion

 .Application.Goto "ComputerName"
 .ActiveCell.FormulaR1C1 = "computername"

 .Application.Goto "OSSerial"
 .ActiveCell.FormulaR1C1 = objSysInfo.SerialNumber

 .Application.Goto "VirtualMemory"
 .ActiveCell.FormulaR1C1 = objSysInfo.VirtualMemory

 .Application.Goto "CPU"
 .ActiveCell.FormulaR1C1 = objSysInfo.CPU

 .Application.Goto "Memory"
 .ActiveCell.FormulaR1C1 = objSysInfo.Memory

 .Application.Goto "OSVersion"
 .ActiveCell.FormulaR1C1 = objSysInfo.OS

 .Application.Goto "OSUser"
 .ActiveCell.FormulaR1C1 = objSysInfo.RegisteredUser

End With

Set objSysInfo = Nothing
Set objExcel = Nothing


