Private Sub Form_Load()
	Set ExcelApp = CreateObject("Excel.Application")
	Set ExcelWB = ExcelApp.Workbooks.Add
	ExcelApp.Visible = True
	Set ExcelWS = ExcelWB.Worksheets("Sheet2")
	ExcelWS.Cells(1,1).Value = "Testing Testing"
	For rowCounter = 2 To 10
		ExcelWS.Cells(rowCounter,1).Value = "Using a loop to fill in rows of 1st Column"
	Next
	ExcelWS.Range("A15","F25").Value = "Using Range to fill in the cells"
	ExcelWS.Range(excelWS.Cells(1,4),excelWS.Cells(10,5)).Value = "Using Range with Cells() Objects"

	rowCount=ExcelWS.UsedRange.Rows.Count
	colCount=ExcelWS.UsedRange.Columns.Count

	Msgbox rowCount
	Msgbox colCount
	For each sheet in ExcelWB.Sheets
		If Not sheet.Name="Sheet2" Then
			sheet.Delete
		End If

	Next
	excelWB.SaveAs "D:\STUDY\UnderstandingExcelObject.xls"
	excelWB.Close
End Sub

Form_Load()


















