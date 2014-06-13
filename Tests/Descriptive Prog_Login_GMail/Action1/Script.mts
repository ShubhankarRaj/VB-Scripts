Class LaunchApp
    Public Default Function Run()
        'Return Nothing if the initialization fails
        Set Run = Nothing
 
        With Browser("title:=.*Google.*")
            If Not .Exist(0) Then
                SystemUtil.Run "iexplore.exe", "http://gmail.com"
            End If
 
            If .WebEdit("name:=Email").Exist(10) Then
                'Return the LoginPage PageObject if UserName field is found
                Set Run = New LoginPage
            End If
        End With
    End Function
End Class

 
'Login process
Class LoginPage
    Public Default Function Run(userName, password)
        With Browser("title:=.*Google.*")
            .WebEdit("name:=Email").Set userName
            .WebEdit("name:=Passwd").Set password
            .WebButton("name:=Sign in").Click()
            .Sync
 
            If InStr(.GetROProperty("title"), "Find a Flight") > 0 Then
                'If Welcome page appears = User has logged in
                'Return FindAFlightPage Object
                Set Run = New FindAFlightPage
            Else
                'If the Welcome page failed to appear = Login failed
                'Stay at the Login Page
                Set Run = Me
            End If
        End With
    End Function    
End Class
 
'Contains methods for the Inbox page
Class FindAFlightPage
    'returns SelectAFlightPage
    Public Function GotoSelectFlightsPage()
        'code
    End Function
 
    ' other methods
End Class

Set LoginPageObject = (New LaunchApp).Run()
 
Set InboxPageObject = LoginPageObject.Run("raj.shubhankar8055", "raj8lm@TCS4U")


