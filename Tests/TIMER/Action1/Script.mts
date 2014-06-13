
Delay(15)

Sub Delay(Sec)
   StartTime = TIMER
   Do
	   Loop until TIMER - StartTime > Sec
End Sub