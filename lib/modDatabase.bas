Attribute VB_Name = "modDatabase"
Option Explicit

Public conn As ADODB.Connection
Public dbOpen As Boolean

Public Sub initDB()
StartConn:
    dbOpen = False
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\BookDB.mdb;Persist Security Info=False"
    conn.CursorLocation = adUseClient
    conn.Open
    If Not conn.State = adStateOpen Then
        Select Case MsgBox("There was an error opening the databse! Please exit and restart the program. Alternately, you can try to connect again.", vbCritical + vbApplicationModal + vbRetryCancel + vbDefaultButton1, "Database Error")
        Case vbRetry
            GoTo StartConn
        Case vbCancel
            End
        End Select
    Else
        dbOpen = True
    End If
End Sub

Public Sub closeDB()
    conn.Close
    dbOpen = False
End Sub
