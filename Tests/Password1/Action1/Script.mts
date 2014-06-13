Function GetPassword( myPrompt )
' This function uses Internet Explorer to
' create a dialog and prompt for a password.

    Dim obj_IE
'  Create an IE object
    Set obj_IE = CreateObject( “InternetExplorer.Application” )
'   specify  the IE  settings
    obj_IE.Navigate “about:blank”
    obj_IE.Document.Title = “Password”
    obj_IE.ToolBar        = False
    obj_IE.Resizable      = False
    obj_IE.StatusBar      = False
    obj_IE.Width          = 300
    obj_IE.Height         = 180
'    Center the dialog window on the screen
    With obj_IE.Document.ParentWindow.Screen
        obj_IE.Left = (.AvailWidth  – obj_IE.Width ) \ 2
        obj_IE.Top  = (.Availheight – obj_IE.Height) \ 2
    End With
  
'     Insert the HTML code to prompt for a password
    obj_IE.Document.Body.InnerHTML = “<DIV align=”"center”"><P>” & myPrompt _
                                  & “</P>” & vbCrLf _
                                  & “<P><INPUT TYPE=”"password”" SIZE=”"20″” ” _
                                  & “ID=”"Password”"></P>” & vbCrLf _
                                  & “<P><INPUT TYPE=”"hidden”" ID=”"OK”" ” _
                                  & “NAME=”"OK”" VALUE=”"0″”>” _
                                  & “<INPUT TYPE=”"submit”" VALUE=”" OK “” ” _
                                  & “OnClick=”"VBScript:OK.Value=1″”></P></DIV>”
'     Make the window visible
    obj_IE.Visible = True
'     Wait till the OK button has been clicked
    Do While obj_IE.Document.All.OK.Value = 0
        WScript.Sleep 200
    Loop
'     Read the password from the dialog window
    GetPassword = obj_IE.Document.All.Password.Value
'     Close and release the object
    obj_IE.Quit
    Set obj_IE = Nothing
End Function



strPw = GetPassword( “Please enter your password:” )
msgbox  “Your password is: ” & strPw

