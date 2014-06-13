' **********************
Dim x0
Set x = CreateObject("Quicktest.Application")
x.launch
x.showpanescreen "Activescreen", True
Wait 3
x.windowstate = "Minimized"
x.Visible = True
Set x = Nothing
'****************************

Msgbox "An Error Occurred", 48
Msgbox "An Error Occurred", 2 
Msgbox " ", 4096

'*****************************

Option Explicit
Dim x,y

x = InputBox("Please enter a number to divide with 100.")

If x<>0 Then
	y=x/100
	Msgbox "100 divided by "& x &" is: "& y &"."
Else
	Err.Raise vbObjectError +15000, "ERR_MSG_UGLY.VBS", "Hey Stupid!! You cant enter Zero."
End If

'*******************************

Option Explicit

Const FILE_NAME = "WSH_DEBUG_TEST_FILE.TXT"
Const COPY_SUFFIX="_COPY"
Const OVERWRITE_FILE=True

Dim objFSO
Dim strExtension
Dim blnFileExists
Dim strNewFileName
Dim strScriptPath

Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptPath = GetScriptPath()
blnFileExists = VerifyFile(strScriptPath, FILE_NAME)

If blnFileExists Then
	strExtension = GetExtension(FILE_NAME)
	strNewFileName = MakeNewFileName(FILE_NAME, strExtension, COPY_SUFFIX)
	CopyFile strScriptPath & FILE_NAME, strScriptPath & strNewFileName, OVERWRITE_FILE

Else
	On Error GoTo 0
	Err.Raise vbObjectError + 10000, "WSH_DEBUG_EXAMPLE.VBS", "Expected file" & FILE_NAME & " not found."
End If

'************************* Supporting Procedures and functions **************************************

Private Sub CopyFile(strFileName, strNewFileName, blnOverWrite)
   objFSO.CopyFile strFileName, strNewFileName, blnOverwrite
End Sub

Private Function GetExtension(strFileName)
   GetExtension = objFSO.GetExtensionName(strFileName)
End Function

Private Function GetScriptPath
   Dim strPath
   strPath = objFSO.GetAbsolutePathName(ScriptFullName)
   strPath = Left(strPath, Len(strPath) - Len(objFSO.GetFileName(strPath)))
End Function

Private Function VerifyFile(strPath, StrFileName)
   VerifyFile = objFSO.FileExists(strFileName & strFileName)
End Function

Private Function MakeNewFileName(strFileName, strExtension, strSuffix)
   MakeNewFileName = Left(strFileName, Len(strFileName) - (1+Len(strExtension))) & strSuffix & "." & strExtension
End Function

