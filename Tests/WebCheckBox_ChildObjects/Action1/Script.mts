'setting the test object as 'Passenger List' web table
Set TestObject = Browser("Nucleation").Page("Nucleation Page").WebTable("Passenger List")
'calling the function to uncheck the checkbox against 'Passenger 1'
Call WebTableCheckBoxONorOFF(TestObject, "Passenger 1", 2, 3, "OFF")
'calling the function to check the checkbox against 'Passenger 2'
Call WebTableCheckBoxONorOFF(TestObject, "Passenger 2", 2, 3, "ON")

'*********************************************************************************************************************************************************************************************************************
'Function Name                     : WebTableCheckBoxONorOFF(ByVal TestObject, ByVal SearchText, ByVal SearchTextInColumn, ByVal CheckBoxColumn, ByVal ONorOFF)
'Function Description             : Function to check or uncheck a checkbox within a web table based on value present in same or another column in the same row.
'Data Parameters                  : TestObject:- Specify the WebTable. eg: Browser("-----").Page("-----").WebTable("-----")
'                                             SearchText:- Specify the text to search in the WebTable. eg: "Passenger 1"
'                                             SearchTextInColumn:- Specify the column in which text is to be searched. eg: 2
'                                             CheckBoxColumn:- Specify the column where the check box resides in the WebTable. eg: 3
'                                             ONorOFF:- Specify the needed state of the checkbox. 'ON' to check and 'OFF' to uncheck
'Created by                           : Kannan S
'Email                                   : info@nucleation.in
'Creation date                       : 22-May-2013
'Website                               : www.nucleation.in
'THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
'LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
'Feel free to use the code as you wish but kindly keep this header section intact.
'Copyright © 2012 - 2013 Nucleation. All Rights Reserved.
'*********************************************************************************************************************************************************************************************************************
Function WebTableCheckBoxONorOFF(ByVal TestObject, ByVal SearchText, ByVal SearchTextInColumn, ByVal CheckBoxColumn, ByVal ONorOFF)
	'in case of any errors
   On Error Resume Next
   'checking whether the object exist
   ObjectExist = TestObject.Exist(10)
   If ObjectExist Then
	   'setting MatchFound as false
	   MatchFound = FALSE
	   'finding the total number of rows in the webtable
	   TotalRows = TestObject.RowCount
	   'looping in the webtable until a match is found
	   For CRow = 1 to TotalRows
		   'checking whether Searched Text is found
		   If SearchText = Trim(TestObject.GetCellData(CRow, SearchTextInColumn)) Then
			   'saving the row where the value is found to a variable
			   FoundRow = CRow
			   'setting MatchFound as true
			   MatchFound = TRUE
			   Exit For
		   End If
	   Next
	   'in case a match is found
	   If MatchFound Then
		   'setting the object web check box
			Set CheckBoxObject = TestObject.ChildItem(FoundRow, CheckBoxColumn, "WebCheckBox", 0)
			'interpreting specified state
			Select Case LCase(Trim(ONorOFF))
				Case "on"
					ONorOFF = 1
					CheckBoxObject.Set "On"
					specifiedText = "Checked"
				Case "off"
					ONorOFF = 0
					CheckBoxObject.Set "Off"
					specifiedText = "Unchecked"
				Case Else
					Reporter.ReportEvent micFail, "Incorrect argument in function call", "Argument 'ONorOFF' should be specified as 'ON' or 'OFF'. You specified the argument as '" & ONorOFF & "'."
					Exit Function
			End Select
			'checking the status of the checkbox
			ActualState = CheckBoxObject.GetROProperty("checked")
			'reporting pass or fail depending on the specified state
			If ActualState = ONorOFF Then
				Reporter.ReportEvent micPass, "CheckBox '" & specifiedText & "' as specified.", "CheckBox '" & CheckBoxObject.GetROProperty("name") & "' against '" & SearchText & "' was '" & specifiedText & "' as specified."
			Else
				Reporter.ReportEvent micFail, "CheckBox not '" & specifiedText & "' as specified.", "CheckBox '" & CheckBoxObject.GetROProperty("name") &  "' against '" & SearchText & "' was not '" & specifiedText & "' as specified."
			End If
			'in case match is not found
       ElseIf Not MatchFound Then
	 		'reporting error if specified text is not found in the webtable
			Reporter.ReportEvent micFail, "Cannot Find Specified Text", "Cannot find the text " & SearchText & " in the object specified in function call."
	   End If
   ElseIf Not ObjectExist Then
	 	'reporting error if the specified object is not found
		Reporter.ReportEvent micFail, "Cannot Find The Object", "Object specified in function call cannot be found."
   End If
End Function