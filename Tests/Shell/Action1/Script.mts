Dim oShell

Set oShell = CreateObject (“WScript.shell”)

oShell.run “cmd /K CD C:\ & Dir”
Msgbox "Check the Window !!"
Set oShell = Nothing
