strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems
var_language =  objOperatingSystem.OSLanguage
Next
Msgbox var_language
Msgbox Environment.Value("OS")
Msgbox Environment.LoadFromFile ("D:\STUDY\VB Scpt\TestData.xml")
Select Case var_language
   Case "1033" ' English - United States
		   DataTable("OS_Language", dtGlobalSheet) = "English"
   Case "1031" ' German - Germany
	   	   DataTable("OS_Language", dtGlobalSheet) = "German"
   Case "1036" ' French - France
	   	   DataTable("OS_Language", dtGlobalSheet) = "French"
   Case "1034" ' Spanish - Spain
	   	   DataTable("OS_Language", dtGlobalSheet) = "Spanish"
   Case "2052" ' Chinese- China
	   	   DataTable("OS_Language", dtGlobalSheet) = "Chinese"
   Case "1041" ' Japanese - Japan
	   	   DataTable("OS_Language", dtGlobalSheet) = "Japanese"
   Case "1040" ' Italian - Italy
	   	   DataTable("OS_Language", dtGlobalSheet) = "Italian"
   Case "1046" ' Portugese - Brazil
	   	   DataTable("OS_Language", dtGlobalSheet) = "Portugese"
   Case Else
			DataTable("OS_Language", dtGlobalSheet) = "English"
End Select 

	
'Once we have the OS Language, we can select the OR file to be picked

Select Case (DataTable("OS_Language", dtGlobalSheet))
	Case "English"
		StrPath = "....\English.tsr"
		RepositoriesCollection.Add(StrPath)
End Select








