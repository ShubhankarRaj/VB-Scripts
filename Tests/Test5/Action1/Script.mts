DataTable.Import("C:\Users\Sony\Desktop\Magna\FIS India Pvt Ltd_ -TS_For the month of August _2012 -.xls")
DataTable.ImportSheet "C:\Users\Sony\Desktop\Magna\FIS India Pvt Ltd_ -TS_For the month of August _2012 -.xls", 1,"My Sheet"
Dim RowCount
'RowCount = DataTable.GetSheet("My Sheet").GetRowCount

Msgbox "Excel Sheet Imported"
Msgbox RowCount

















