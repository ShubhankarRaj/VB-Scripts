Dim x,y
x = DataTable("Parameter1", dtGlobalSheet)
Msgbox x

y = DataTable("Parameter1", dtLocalSheet)
Msgbox y

z = DataTable("Parameter1", 2)
Msgbox z

a = DataTable.GlobalSheet.GetParameter("Parameter1").value
Msgbox a

b = DataTable.isSheetExist("Global")
Msgbox b

Msgbox Environment("TestIteration")


'************************************************

isDate(datBirth)
2de_dAY = Day(datBirth)
2de_montH = Month(datBirth)
2de_year = Year(datBirth)

'************************************************


Class LaunchApp
   Public Default Function Run()
		Set Run = Nothing
		With Browser("title:=.*Google.*")
			If Not.Exist(0) Then
				SystemUtil.Run "iexplore.exe", "http://gmail.com"
			End If

			If .WebEdit("name:=Email").Exist(10) Then
				Set Run = New LoginPage
			End If
		End with
   End Function
End Class

Class LoginPage
   Public Default Function Run(userName, Password)
	  With Browser("title:=.*Google.*")
		.WebEdit("name:=Email").Set userName
		.WebEdit("name:=Passwd").Set Password
		.WebButton("name:=Sign in").Click()
		.Sync
	  End with
   End Function
End Class


Set LoginPageObject = (New LaunchApp).Run()
 
Set InboxPageObject = LoginPageObject.Run("raj.shubhankar8055", "raj8lm@TCS4U")


'*******************************************************


Msgbox "An Error Occurred", 48
Msgbox "An Error Occurred", 2 
Msgbox " ", 4096


'*****************************************

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWB = ExcelApp.Workbooks.Add
ExcelApp.Visible = True
Set ExcelWS = ExcelWB.Worksheets("Sheet1")
ExcelWS.Cells(1,1).Value = "Testing Testing"
For rowCounter = 2 To 10
	ExcelWS.Cells(rowCounter,1).Value = "Using a loop to fill in rows of first column."
Next

ExcelWS.Range("A15","F25").Value = "Using Range to fill in the values."
ExcelWS.Range(ExcelWS.Cells(1,4), ExcelWS.Cells(5,5)) = "Using Range with Cell() objects"

rowCount = ExcelWS.UsedRange.Rows.Count
colCount = ExcelWS.UsedRange.Cols.Count

For each sheet in ExcelWB.Sheets
	If Not sheet.Name = "Sheet2" Then
		sheet.Delete
	End If
Next
ExcelWB.SaveAs "Test Results"
ExcelWB.Close

Set ExcelWB = Nothing
Set ExcelWS = Nothing
Set ExcelApp = Nothing

'*****************************************************
Dim ExcelWB, ExcelS, eRow, eCol, Val_frm_file


Function value_from_excel(ExcelPath, ExcelSheet, ExcelRow, ExcelColumn)
	On Error Resume Next
	Set Excel_obj = CreateObject("Excel.Application")
	Excel_obj.workbooks.Open ExcelPath
	Set My_Sheet = Excel_obj.sheets.item(ExcelSheet)
	Excel_Value = My_Sheet.cells(ExcelRow, ExcelColumn)
	Excel_obj.application.quit
	Set Excel_obj = Nothing
	value_from_excel = Excel_Value
	On Error Goto 0

	If Err.Count > 0 Then
		Msgbox "Error is generated."
	End If
End Function

excelValue = value_from_excel(Path, Sheet, Row, Col)
Msgbox excelValue

'****************************************************

Dim R, str
str = "1111111111"
Set R = New RegExp
R.pattern = "."
R.Global = True

Set Matches = R.Execute(str)
For each Match in Matches
	Msgbox (Match+1)&Match.Value
	Msgbox Matches.Count
Next

'******************************************

Dim objFSO
Dim objStream
Dim txtString

Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set objStream = objFSO.OpenTextFile("someno.txt")

txtString = ""
Do While Not objStream.AtEndOfStream
    txtString = txtString & objStream.ReadLine & vbNewLine
Loop

'***************************************

DomainName =
ProjectName = 
UserName =
Password = 
ServerName =

Set objTDConn = QCUtil.QCConnection

objTDConn.InitConnectionEx QC Servername
objTDConn.ConnectProjectEx DomainName, ProjectName, UserName, Password

If objTDConn.LoggedIn Then
	Connected Successfully
End If

'**********************************************************

If TeWindow("TeWindow").TeScreen("TeScreen").Exist(0) Then
	If TeWindow("TeWindow").TeScreen("TeScreen").TeField("TeField").Exist(0) Then
		If lcase(FieldName)<>"password" Then
			TeWindow("TeWindow").TeScreen("TeScreen").TeField(FieldName) Set NewText
		
		else
			TeWindow("TeWindow").TeScreen("TeScreen").TeField(FieldName).SetSecure Crypt.Encrypt(New Text)
		End If	
End If