
Dim SizeOfArray
Dim i,j, DonationAmount
Dim TotalAmount
String Side
ReDim Amount(0,0)
SizeOfArray = Inputbox("Kindly enter the size of the 2D Array in which the locality is arranged")
SizeOfArray = SizeOfArray-1
ReDim Amount(SizeOfArray, SizeOfArray)
For i = 0 to SizeOfArray
	For j = 0 to SizeOfArray
'		Msgbox ("Rows: " &i& ", Columns: " &j)
        DonationAmount = Inputbox("Kindly Enter the amount of money to be donated in this house")
		Amount(i,j) = DonationAmount
	Next
Next
For i = 0 to SizeOfArray
	For j = 0 to SizeOfArray
    Msgbox Amount(i, j)
	If i=j Then
		If i = SizeOfArray Then
			
		End If
	End If
'	If Amount(i,j)<Amount(i,j+1) && Then
'		 Side = "RIGHT"
     Next
Next
	











