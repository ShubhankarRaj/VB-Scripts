Dim lngAge
'Do
lngAge=InputBox("Please enter your age in years: ")
'If IsNumeric(lngAge) Then
	lngAge=CLng(lngAge)
	Msgbox lngAge
	lngAge=lngAge+50
	Msgbox lngAge
	MsgBox "In 50 years you would be "&lngAge&" years old."
'Else

	Msgbox "Sorry!! You have no entered a valid number."
'End If
'Loop until IsNumeric(lngAge)=TRUE






