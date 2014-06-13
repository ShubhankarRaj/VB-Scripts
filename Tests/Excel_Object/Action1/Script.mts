

Dim Excel_path, Excel_sheet, eRow, eColumn, Value_from_file

Excel_path="D:\STUDY\VB Scpt\Test.xlsx"
Excel_sheet="Sheet1"
eRow= 8
eColumn= 4


Function value_from_Excel(Excel_path, Excel_sheet, eRow, eColumn)
    On Error Resume Next
   Set Excel_object = createobject("Excel.Application")
	Excel_object.workbooks.open Excel_path
	Set My_sheet = Excel_object.sheets.item(Excel_sheet)
	Excel_value= My_sheet.cells(eRow,eColumn) 
	Excel_object.application.quit
	Set Excel_object = Nothing
    value_from_Excel=Excel_value
	On Error Goto 0
	
	If Err.Count>0 Then
		Msgbox "Error is Generated"
		Msgbox Err.Count()
	End If
	On Error Goto 0
End Function

Value_from_file = value_from_Excel(Excel_path,Excel_sheet,eRow,eColumn)
Msgbox Value_from_file



