Dim Cntr, WrdLngth, WrdBldr

'WrdLngth = Len(" VB Script is going on Great !")
'Msgbox WrdLngth
'
'For Cntr = 1 to WrdLngth
'	'Msgbox Mid (" VB Script is going on Great !", Cntr, 1)
'	WrdBldr = WrdBldr & Mid (" VB Script is going on Great !", Cntr, 4)
'	Msgbox WrdBldr
'Next
'
'Msgbox WrdBldr

Dim InWrd
InWrd = Inputbox ("Type in the word whose length u want to know")
WrdLngth = Len(InWrd)

Msgbox Inwrd & " contains " & WrdLngth & " characters"

