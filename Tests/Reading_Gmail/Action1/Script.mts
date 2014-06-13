
Set oIE = CreateObject("InternetExplorer.Application")
oIE.Navigate ("https://gmail.com")
oIE.Visible = True
Dim user_Name, passWord
user_Name = "letsaladeen"
passWord = "raj8lm@TCS4U"
Web_Login user_Name,passWord
Count_Mail sender_Name
Read_Mail sender_Name
Write_Mail sender_Name

Function Web_Login(sUserName, sPassword)
   With Browser("title:=Gmail.*").Page("micclass:=Page")
        'Check if the UserName field exists
        If .WebEdit("html id:=Email").Exist(0) Then
            .WebEdit("html id:=Email").Set sUserName    'Set UserName
            .WebEdit("html id:=Passwd").SetSecure sPassword    'Set Password
            .WebButton("name:=Sign in").Click        'Click Submit
            .Sync
        End If

        'Check for Link Inbox(xyz)
        If .Link("innertext:=Inbox.*").Exist(15) Then 
			GMailLogin = True
			Reporter.ReportEvent micPass, "Login Successfull", "Login Details"
		End If

   End With
End Function
 @@ hightlight id_;_Browser("Google").Page("Google").WebEdit("q")_;_script infofile_;_ZIP::ssf1.xml_;_





