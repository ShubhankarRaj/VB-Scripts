'Msgbox Browser("MakeMyTrip, India's No").Page("MakeMyTrip, India's No").WebList("selnoOfAdults").GetROProperty("default value")
'Msgbox Browser("MakeMyTrip, India's No").Page("MakeMyTrip, India's No").WebList("selnoOfAdults").CheckProperty("selection", 6)
'
'' Solution to Question 15
'Browser("MakeMyTrip, India's No").Page("MakeMyTrip, India's No").WebButton("Search Flights").CaptureBitmap "D:\STUDY\VB Scpt\WebButton.bmp",True
'Msgbox ("Height = ") &Browser("MakeMyTrip, India's No").Page("MakeMyTrip, India's No").WebButton("Search Flights").GetROProperty("height") & vbNewLine & "Width = " &Browser("MakeMyTrip, India's No").Page("MakeMyTrip, India's No").WebButton("Search Flights").GetROProperty("width")


''' Solution to question 13 & 14
'Set oLink = Description.Create()
'oLink("micclass").value = "Link"
'
'set obj = Browser("MakeMyTrip, India's No").Page("MakeMyTrip, India's No").ChildObjects (oLink)
'linkCount = obj.count
'
'For i = 51 to 55
'	Msgbox obj(i).GetROProperty("text")
'Next
'
'For i = 0 to linkCount-1
'	If Instr(obj(i).GetROProperty("text"), "Manage Bookings") > 0 Then
'	Msgbox obj(i).GetROProperty("text")
'	Msgbox i
'	End If
'Next
'
'
''' Question 16
'Msgbox Browser("MakeMyTrip, India's No").Page("MakeMyTrip, India's No").GetROProperty("URL")
'
'
''' Question 17
'If Browser("MakeMyTrip, India's No").WinStatusBar("msctls_statusbar32").GetROProperty("text") = "Done" Then
'	Msgbox "The page is loaded"
'End If


'''Question 18
'Browser("MakeMyTrip, India's No").Refresh


''' Question 19
'Msgbox Browser("MakeMyTrip, India's No").GetROProperty("title")

''' Question 20
'Set oDesc = Description.Create()
'oDesc("micclass").value = "Frame"
'
'Set oFrame = Browser("MakeMyTrip, India's No").Page("MakeMyTrip, India's No").ChildObjects (oDesc)
'frameCount = oFrame.Count
'For i = 0 to frameCount-1
'oFrame(i).CaptureBitmap ("D:\STUDY\VB Scpt\Frame"&i&".bmp"), True
'Next


'' Question 21
' Using Start Transaction and End Transaction


'' Question 22 and 24
'x = TIMER
'SystemUtil.Run "iexplore.exe", "http://www.makemytrip.com/"
'Browser("MakeMyTrip, India's No").Page("MakeMyTrip, India's No").Sync
'y = TIMER
'Msgbox "Page load time = "&(y-x)

'' Question 23
Msgbox Browser("MakeMyTrip, India's No").Page("MakeMyTrip, India's No").WebElement("Get Started").Object.title

'' Question 25
'DataTable.ImportSheet "D:\STUDY\QTP Workshop Final\Topics.xls", Global
DataTable.ImportSheet "D:\STUDY\QTP Workshop Final\Topics.xls", 1, "Global"
Msgbox "Done"



