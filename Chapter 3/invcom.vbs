'invcom.vbs
'inventories computer upon logon
Option Explicit

Const DataFile = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\odin\Data\qadna.mdb"

Const adOpenDynamic= 2 
Const adLockOptimistic= 3 

Dim objSysInfo, objService, objNetwork, objRst, objConn
Dim strComputerName, nCompID, bUpdate, bNew
  
Set objNetwork = CreateObject("Wscript.Network")
strComputerName = objNetwork.ComputerName

'create ADO object and open QADNA database
Set objConn = CreateObject("ADODB.Connection")
Set objRst = CreateObject("ADODB.Recordset")
objConn.Open DataFile

bUpdate = False
bNew = False

'open tblcomputers table and determine if computer exists
objRst.Open "tblComputers", objConn, adOpenDynamic, adLockOptimistic
objRst.Find "ComputerName='" & strComputerName & "'"

'if computer does not exist, add to inventory
If objRst.EOF Then
    objRst.AddNew
    objRst("ComputerName") = strComputerName
    bNew = True
Else
   'check if computer needs updating..
   If Date>= CDate(objRst("UpdateDate")) Then bUpdate = True
End If

If bUpdate Or bNew Then  

 Set objService = _
    GetObject("winmgmts:{impersonationLevel=impersonate}!root\cimv2")

 'create an instance of the SysInfo object    
 Set objSysInfo = GetObject("script:\\odin\wsc$\SysInfo.wsc")

 objRst("BIOSVersion") = objSysInfo.BIOSVersion
 objRst("CPU") = objSysInfo.CPU
 objRst("Memory") = objSysInfo.Memory
 objRst("OSVersion") = objSysInfo.OS
 objRst("RegisteredUser") = objSysInfo.RegisteredUser
 objRst("OSSerialNum") = objSysInfo.SerialNumber
 objRst("VirtualMemory") = objSysInfo.VirtualMemory
 objRst("UpdateDate") = Date  + 180
 nCompID = objRst("ComputerID")
 
 objRst.Update
 objRst.Close
 'if new computer in inventory need to get the computer ID 
 If bNew Then
  Set objRst = objConn.Execute("SELECT Max(ComputerID) FROM tblComputers")
  nCompID = objRst(0)
 End If

 'delete any existing computer inventory for computer
 objConn.Execute "Delete From tblComputerItems " & _ 
                "Where ComputerID = "  & nCompID
 
 'inventory the hardrive, memory, network adapter and modem info
 WMIInv "Select Description, Size From Win32_DiskDrive" _
       , Array("Description" , "Size"),1, nCompID

 WMIInv "Select AdapterType, AdapterRAM From " & _ 
       "Win32_VideoConfiguration"  _
       ,Array("AdapterType", "AdapterRAM"), 2, nCompID

 WMIInv "Select Description, MacAddress From " & _
       "Win32_NetworkAdapterConfiguration Where IPEnabled = True" _
       , Array("Description", "MacAddress"), 3, nCompID

 WMIInv "Select Model, ProviderName From " & _ 
       "Win32_POTSModem", Array("Model", "ProviderName"), 4, nCompID

 WMIInv "Select Description, Manufacturer From " & _ 
       "Win32_SCSIController" , Array("Description", _ 
       "Manufacturer"), 5, nCompID

End If

objConn.Close

'WMIInv
'returns specified information from WMI query
'Parameters:
'strQuery    SQL query to execute against WMI service
'aFields     Array of fields to store
'nType       Type of item
'nCompID     Unique computer identifier
Sub WMIInv(strQuery, aFields, nType, nCompID)
 Dim objInstance, objEnumerator, objProp
 Dim nF, strSQL, strText

 Set objEnumerator = objService.ExecQuery(strQuery)

  'loop through each instance and build the output lines
  For Each objInstance In objEnumerator
  strSQL="INSERT INTO tblComputerItems (ComputerID,ItemType, Item1,Item2)" _
         & " Values (" & nCompID & "," & nType & ","
 
  'loop through each property
   For nF = LBound(aFields) To UBound(aFields)
    Set objProp = objInstance.Properties_(aFields(nF))
    strText = objProp.Value
    strSQL = strSQL & "'" & strText & "',"
   Next

  strSQL = Left(strSQL, Len(strSQL) - 1) & ");"
  objConn.Execute strSQL

 Next

End Sub
