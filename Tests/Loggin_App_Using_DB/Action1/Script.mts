Option Explicit
Dim objConnection, objRecordSet

Set objConnection = CreateObject("Adodb.Connection")

Set objRecordSet = CreateObject("Adodb.RecordSet")

objConnection.Provider = ("Microsoft.Jet.OLEDB.4.0")

objConnection.Open "C:\.... DbPath"

objRecordSet.Open "Select * From Login",objConnection

Do until objRecordSet.EOF = True
	SystemUtil.Run "C:\.....ApplicationPath"
	Dialog("text:=Login").Activate
	Dialog("text:=Login").Winedit("attached text:=Agent Name:").Set objRecordSet.Fields("Agent")
	Dialog("text:=Login").Winedit("attached text:=Password:").Set objRecordSet.Fields("Password")
	Dialog("text:=Login").Winbutton("text:=OK","width:=60").Click
	Window("text:=Flight Reservation").Close
	objRecordSet.MoveNext
Loop

objRecordSet.Close
objConnection.Close

Set objRecordSet = Nothing
Set objConnection = Nothing