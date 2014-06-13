 @@ hightlight id_;_197736_;_script infofile_;_ZIP::ssf1.xml_;_
Browser("QuickTest Professional").Page("QuickTest Professional").Sync @@ hightlight id_;_Browser("QuickTest Professional").Page("QuickTest Professional")_;_script infofile_;_ZIP::ssf3.xml_;_
Dim i,j
Dim rowCount, ColCount
Dim cellText, objTable

Set objTable = Browser("QuickTest Professional").Page("QuickTest Professional").WebTable("Mercury Application")

rowCount = objTable.RowCount
ColCount = objTable.ColumnCount(1)

Set outSheet = DataTable.AddSheet("Output")

For i = 1 to ColCount
	cellText = objTable.GetCellData(1,i)
	outSheet.AddParameter cellText,""
	'Msgbox cellText
Next

For i = 2 to rowCount
	outSheet.SetCurrentRow i-1
	ColCount = objTable.ColumnCount(i)
	For j =1 to ColCount
		cellText = objTable.GetCellData(i,j)
		'Msgbox cellText
		outSheet.GetParameter(j).value = cellText
	Next
Next
'Msgbox "Rows: "&rowCount&" and Columns : "&ColCount

a = DataTable.GetSheet("Output").GetParameter("Versions").ValuebyRow(2)
Msgbox a
Browser("QuickTest Professional").Close











