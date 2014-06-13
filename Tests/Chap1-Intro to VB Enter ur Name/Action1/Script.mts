Dim a, b, c, count
count =0
Do

	a = "No"
	b = Inputbox ("Please!! Enter your name:")
    If b = "" Then
        Msgbox ("Please try to enter your name!!")
        a = "Yes"
	Else
		c = "Hello, "&b&" ! Great to see you."
	End If
	count = count + 1
	If count = 2  Then
			Msgbox ("What the FUCK r u doing ?? Enter the God Damn Name.. u asshole  !!")
	End If
Loop While a = "Yes"

Msgbox c




