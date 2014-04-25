VERSION 5.00
Begin VB.Form LoginView 
   Caption         =   "Login"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   3720
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5040
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtUsername 
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   2040
      Width           =   2895
   End
   Begin VB.PictureBox ShopLogo 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9120
      ScaleHeight     =   495
      ScaleWidth      =   735
      TabIndex        =   0
      Tag             =   "no_resize"
      ToolTipText     =   "Go to Main Screen"
      Top             =   240
      Width           =   735
   End
   Begin VB.Label cmdCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5040
      MousePointer    =   10  'Up Arrow
      TabIndex        =   8
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label cmdLogin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2640
      MousePointer    =   10  'Up Arrow
      TabIndex        =   7
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label lblUsername 
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label MessageBar 
      Alignment       =   2  'Center
      Caption         =   "Enter your username and password and click ""Login"""
      BeginProperty Font 
         Name            =   "Roboto Condensed Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   9135
   End
   Begin VB.Label OptionIcon 
      AutoSize        =   -1  'True
      Caption         =   "//"
      BeginProperty Font 
         Name            =   "Roboto Thin"
         Size            =   12
         Charset         =   0
         Weight          =   250
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   180
   End
   Begin VB.Label OptionLabel 
      AutoSize        =   -1  'True
      Caption         =   "Login to Application:"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   1950
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   360
      X2              =   9720
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label ShopName 
      AutoSize        =   -1  'True
      Caption         =   "Students Book House"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "LoginView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Dim Token As Long
Dim C As Long
Dim i As Integer
 
Dim conn As ADODB.Connection
Dim login As ADODB.Recordset

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdLogin.BackColor = &H8000000D
cmdCancel.BackColor = &HC0&
cmdLogin.FontItalic = False
cmdCancel.FontItalic = True

End Sub

Private Sub cmdLogin_Click()
'MessageBar.Visible = False
Timer1.Enabled = False

If txtUsername.Text = "" Then
    MessageBar.Caption = "ERROR: Username cannot be empty!"
    'MessageBar.Visible = True
    txtUsername.SetFocus
    Timer1.Enabled = True
    Exit Sub
ElseIf txtPassword.Text = "" Then
    MessageBar.Caption = "ERROR: Password cannot be empty!"
    'MessageBar.Visible = True
    txtPassword.SetFocus
    Timer1.Enabled = True
    Exit Sub
Else
StartConn:
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
    End If
    
    Set login = New ADODB.Recordset
    login.CursorType = adOpenDynamic
    login.CursorLocation = adUseClient
    login.LockType = adLockOptimistic
    login.Open "Select * from Users where UserName='" & txtUsername.Text & "'", conn, login.CursorType, login.LockType, adCmdUnknown
    
    If login.EOF Then
        MessageBar.Caption = "ERROR: No such user exists! Please check for spelling errors."
        'MessageBar.Visible = True
        txtUsername.SetFocus
        Timer1.Enabled = True
        Exit Sub
    Else
        If login.Fields("Password") = txtPassword.Text Then
            Me.Hide
            HomeView.Show
        Else
            MessageBar.Caption = "ERROR: Wrong password! Please check for spelling/capitalization errors."
            'MessageBar.Visible = True
            txtPassword.SetFocus
            Timer1.Enabled = True
            Exit Sub
        End If
    End If
End If

End Sub

Private Sub cmdLogin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdLogin.BackColor = &HC0&
cmdCancel.BackColor = &H8000000D
cmdLogin.FontItalic = True
cmdCancel.FontItalic = False

End Sub

Private Sub Form_Load()
Token = InitGDIPlus
C = Me.BackColor
If C < 0 Then C = GetSysColor(C - &H80000000)
 
ShopLogo.Picture = LoadPictureGDIPlus(App.Path & "\Images\logo.png", 35, 35, C, True)
cmdLogin.BackColor = &H8000000D
cmdCancel.BackColor = &H8000000D

i = 0
Timer1.Enabled = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdLogin.BackColor = &H8000000D
cmdCancel.BackColor = &H8000000D
cmdLogin.FontItalic = False
cmdCancel.FontItalic = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
FreeGDIPlus Token
End Sub

Private Sub Timer1_Timer()
MessageBar.FontBold = Not MessageBar.FontBold

End Sub
