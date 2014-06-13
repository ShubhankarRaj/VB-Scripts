strComputer =  "."

'Set objWMIService = GetObject("winmgmts:" & _
'"{impersonationLevel=Impersonate}!\\" & strComputer & "\root\cimv2")

'Set objWMIService = GetObject("winmgmts:")
'Set colItems = objWMIService.ExecQuery("Select * from Win32_Service")
'For Each objItem in colItems
'    Msgbox objItem.Name
'Next

'Set objSWbemServices = GetObject("winmgmts:\\" & strComputer & "\root")
Set objSWbemServices = GetObject("winmgmts:")
Set colNameSpaces = objSWbemServices.InstancesOf("__NAMESPACE")
For Each objNS in colNameSpaces
	Msgbox objNS.Name
Next







