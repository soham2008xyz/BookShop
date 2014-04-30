Attribute VB_Name = "modUser"
Public username As String, password As String
Public lastLogin As Date
Public loggedIn As Boolean

Public Sub initUser()
    username = ""
    password = ""
    loggedIn = False
End Sub

Public Function hashPwd(pwd As String) As String
    Dim i As Integer
    
End Function

Public Sub usrLogin(usr As String, pwd As String)
    If Not loggedIn Then
        username = usr
        password = pwd
        loggedIn = True
        lastLogin = Time
        Debug.Print username & " logged in at " & lastLogin
    Else
        Debug.Print "Already Logged in!" & vbCrLf & "Username: " & username & vbCrLf & "Password: " & password
    End If
End Sub

Public Sub usrLogout()
    If loggedIn Then
        Debug.Print username & " logged out at " & Time
        username = ""
        password = ""
        loggedIn = False
    Else
        Debug.Print "Already Logged out!"
    End If
End Sub

