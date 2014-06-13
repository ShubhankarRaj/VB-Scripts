
'Dialog("Login").WinEdit("Agent Name:").Set "scott" @@ hightlight id_;_3410808_;_script infofile_;_ZIP::ssf6.xml_;_
'Dialog("Login").WinEdit("Agent Name:").Type  micTab  @@ hightlight id_;_3410808_;_script infofile_;_ZIP::ssf7.xml_;_
' @@ hightlight id_;_3148892_;_script infofile_;_ZIP::ssf8.xml_;_
'Dialog("Login").WinEdit("Password:").SetSecure "mercury" @@ hightlight id_;_3148892_;_script infofile_;_ZIP::ssf13.xml_;_
'Dialog("Login").WinEdit("Password:").Type  micReturn  @@ hightlight id_;_3148892_;_script infofile_;_ZIP::ssf14.xml_;_
'Window("Flight Reservation").WinButton("Button").Click @@ hightlight id_;_462202_;_script infofile_;_ZIP::ssf15.xml_;_
'Window("Flight Reservation").Dialog("Graph").Close
'Window("Flight Reservation").WinButton("Button_2").Click @@ hightlight id_;_5178734_;_script infofile_;_ZIP::ssf16.xml_;_
'z = Window("Notepad").WinEditor("Edit").GetVisibleText
'Window("Notepad").Close
'Msgbox z
'Window("Flight Reservation").Close
'
'SystemUtil.Run "C:\Users\Sony\Desktop\NewQTPNotepad.txt"
'Window("Notepad").Type z

Set outSheet = DataTable.AddSheet("Output")


Msgbox "Output sheet added?"









