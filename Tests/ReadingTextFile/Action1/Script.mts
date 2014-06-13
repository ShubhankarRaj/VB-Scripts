Dim objFSO
Dim objStream
Dim txtString

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objStream = objFSO.OpenTextFile("someno.txt")

txtString = ""
Do While Not objStream.AtEndOfStream
	txtString = txtString & objStream.ReadLIne & vbNewLine
Loop

If txtString <> "" Then
	Msgbox txtString
Else
	Msgbox "The file is empty."
End If
