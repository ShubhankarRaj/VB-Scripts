'Dim Uid, Pwd
'
'Uid = datatable.Value("userName", dtGlobalSheet)
'Pwd = datatable.Value("passWord", dtGlobalSheet)
'
'Browser("mail Yahoo").page("mail Yahoo").webEdit("User Name").Set "Uid"
'Browser("mail Yahoo").page("mail Yahoo").webEdit("User Name").Set "Pwd"
'
'DataTable.SetNextRow
'
'DataTable.AddSheet(Action2)

Variable = DataTable.AddSheet("My Sheet").AddParameter("Time","8.00")
Msgbox Variable

DataTable.DeleteSheet("My Sheet")

Wait 10
