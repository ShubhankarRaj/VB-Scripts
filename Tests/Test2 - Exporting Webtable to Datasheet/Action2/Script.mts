Dim i, iCount

Set oGlobal = DataTable.GlobalSheet
iCount	= oGlobal.RowCount

For i = 1 to iCount
	oGlobal.SetCurrentRow i

	Msgbox DataTable("Versions")

Next