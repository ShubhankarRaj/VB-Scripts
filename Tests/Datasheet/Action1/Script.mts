Dim x,y

x = DataTable("Parameter1",dtGlobalSheet)
Msgbox x
y = DataTable("Parameter1",dtLocalSheet)
Msgbox y
z = DataTable.Value("Parameter1", 2)
Msgbox z
a = DataTable.GlobalSheet.GetParameter("Parameter1").value
Msgbox a

b = isSheetExist("Global")
Msgbox b

STR = "Current QTP iteration: " & Environment ("TestIteration") & vbNewLine & DataTable ("Parameter1", dtGlobalSheet) & vbNewLine & DataTable ("Parameter2", dtGlobalSheet)
MsgBox STR












