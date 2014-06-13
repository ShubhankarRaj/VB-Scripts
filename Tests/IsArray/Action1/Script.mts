
Option Explicit
Dim arrArray
Dim i
Dim sTmp
Dim charactr

arrArray=Array(81,21,32,45,73,58,123,125,66)

If IsArray(arrArray) Then
	For i=LBound(arrArray) To UBound(arrArray)
        msgbox arrArray(i)
		arrArray(i) = Chr (charactr)
	Next
End If

If IsArray(arrArray) Then
	sTmp=Join(arrArray, vbNullString)
	MsgBox sTmp
End If
Erase arrArray