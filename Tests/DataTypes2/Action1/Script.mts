Dim datBirth

datBirth=Inputbox ("Please Enter your Date of Birth")

If IsDate(datBirth) Then
	Msgbox "You were born on day " &Day(datBirth)& " of month "&Month(datBirth)& " of year "&Year(datBirth)& "."

Else
	Msgbox "Sorry !! The Date entered was not correct"
End If
