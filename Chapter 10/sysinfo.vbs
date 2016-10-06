Option Explicit

Class SysInfo

Dim objService, objComputer, objProcessor
Dim objLocator, objOS, objMemory, objBIOS

Private Sub Class_Terminate()
 Set objLocator = Nothing
 Set objService = Nothing
 Set objOS = Nothing
 Set objMemory = Nothing
 Set objBIOS = Nothing
 Set objComputer = Nothing
 Set objProcessor = Nothing
End Sub

Private Sub Class_Initialize()

  Set objLocator = CreateObject("WbemScripting.SWbemLocator")

  'connect to specified machine
  Set objService = objLocator.ConnectServer()
  objService.Security_.ImpersonationLevel = 3

  Set objOS = GetReference("Win32_OperatingSystem")
  Set objMemory = GetReference("Win32_LogicalMemoryConfiguration")
  Set objBIOS = GetReference("Win32_BIOS")
  Set objComputer = GetReference("Win32_ComputerSystem")
  Set objProcessor = GetReference("Win32_Processor")
 End Sub

Private Function GetReference(strObjectName)
 Dim objInstance, objObjectSet

 'get reference to object
 Set objObjectSet = objService.InstancesOf(strObjectName)
 'loop through and get reference to specified
 For Each objInstance In objObjectSet
    Set GetReference = objInstance
 Next

End Function

Public Property Get BIOSObject()
  Set BIOSObject = objBIOS
End Property

Public Property Get ProcessorObject()
  Set ProcessorObject = objProcessor
End Property

Public Property Get MemoryObject()
  Set MemoryObject = objMemory
End Property

Public Property Get OSObject()
  Set OSObject = objOS
End Property


End Class