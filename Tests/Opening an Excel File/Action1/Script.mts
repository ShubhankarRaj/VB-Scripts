		Dim ExclPath, oExcl, oWorkBook, i, charArray, charString, SyedPart
		ExclPath = "C:\Users\Sony\Desktop\New folder\Unlock\Expense Sheet_VGR.xlsx"
		Set oExcl = CreateObject("Excel.Application")
		charString="a b c d e f g h i j k l m n o p q r s t u v  w x y z A BC D E F G H I J K L M N O P Q R S T U  V W X Y Z 0 1 2 3 4 5 6 7 8 9 ! @ # $ % ^ & * ( ) - _ + = [ ] { } \ | ; : ' < > , . / ?"
		charArray = Split(charString," ")
		RajPart ="raj8"
		TestSetCount = UBound(charArray)
		For each x in CharArray
			char1 = x
			For each y in CharArray
				char2 = y
				For each z in CharArray
					char3 = z
					For each j in CharArray
						On Error Resume Next
						char4 = j
						SyedPart = (char1)&(char2)&(char3)&(char4)
						TotalPaswd = rajPart&SyedPart
						
		'				oExcl.Workbooks.Open (ExclPath), Password = TotalPaswd
						oExcl.Workbooks.Open ExclPath, , , ,TotalPaswd
		
'						If (Err.Number <>0 ) Then
'							Window("Microsoft Excel").Dialog("Microsoft Office Excel").WinButton("OK").Click
'							Window("Microsoft Excel").Close
'                        End If

						If (Err.Number = 0) Then
							MsgBox TotalPaswd
							Exit for
						End If

						On Error Goto 0
		'				Window("Microsoft Excel").Window("Password").WinObject("raj8vgr3").Click 75,11 @@ hightlight id_;_1116614_;_script infofile_;_ZIP::ssf4.xml_;_
		'				Window("Microsoft Excel").Window("Password").WinObject("raj8vgr3").Type TotalPaswd @@ hightlight id_;_1116614_;_script infofile_;_ZIP::ssf5.xml_;_
		'				Window("Microsoft Excel").Window("Password").WinObject("raj8vgr3").Type  micReturn  @@ hightlight id_;_1116614_;_script infofile_;_ZIP::ssf6.xml_;_
		'				Window("Microsoft Excel").Dialog("Microsoft Office Excel").WinButton("OK").Click @@ hightlight id_;_2361770_;_script infofile_;_ZIP::ssf7.xml_;_
		'				Window("Book1").WinObject("NetUIHWND").Click 745,37 @@ hightlight id_;_1968468_;_script infofile_;_ZIP::ssf8.xml_;_
		'				Window("Microsoft Excel").Close
		
						Next
				Next
			Next
		Next

oExcl.Close
Set oExcl = Nothing
























