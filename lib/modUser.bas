Attribute VB_Name = "modUser"
Public username As String, password As String, usertype As String
Public lastLogin As Date
Public loggedIn As Boolean

Public Sub initUser()
    username = ""
    password = ""
    usertype = "USER"
    loggedIn = False
End Sub

Public Function hashPwd(pwd As String) As String
    Dim i As Integer
    
End Function

Public Sub usrLogin(usr As String, pwd As String, ut As String)
    If Not loggedIn Then
        username = usr
        password = pwd
        loggedIn = True
        usertype = ut
        lastLogin = Time
        Debug.Print username & " logged in at " & lastLogin
    Else
        Debug.Print username & " is already logged in! Please logout first."
    End If
End Sub

Public Sub usrLogout()
    If loggedIn Then
        Debug.Print username & " logged out at " & Time
        username = ""
        password = ""
        usertype = "USER"
        loggedIn = False
    Else
        Debug.Print "No user is logged in! Please login first."
    End If
End Sub

