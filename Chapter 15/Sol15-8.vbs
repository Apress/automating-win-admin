Const MediumProtection = 2
'get a reference to the directory NewApp under the root directory of the first site
Set objNewApp = GetObject("IIS://thor/W3SVC/1/Root/NewApp")

'create a new application for the directory
objNewApp.AppCreate True
objNewApp.AppFriendlyName = "Default Application"
objNewApp.AppIsolated = MediumProtection

objNewApp.SetInfo
