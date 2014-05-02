VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form LoginView 
   AutoRedraw      =   -1  'True
   Caption         =   "Login"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   Picture         =   "LoginForm.frx":0000
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
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   4080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1366
      ScreenHeightDT  =   768
      ScreenWidthDT   =   1366
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   5010
      FormWidthDT     =   10380
      FormScaleHeightDT=   4425
      FormScaleWidthDT=   10140
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
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
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   360
      X2              =   9720
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label ShopName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
Dim Token As Long
Dim C As Long
Dim exitVal As Integer

Dim login As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdLogin.BackColor = &H8000000D
    cmdCancel.BackColor = &HC0&
    cmdLogin.FontItalic = False
    cmdCancel.FontItalic = True
End Sub

Private Sub cmdLogin_Click()
    Timer1.Enabled = False

    If txtUsername.Text = "" Then
        MessageBar.Caption = "ERROR: Username cannot be empty!"
        txtUsername.SetFocus
        Timer1.Enabled = True
        Exit Sub
    ElseIf txtPassword.Text = "" Then
        MessageBar.Caption = "ERROR: Password cannot be empty!"
        txtPassword.SetFocus
        Timer1.Enabled = True
        Exit Sub
    Else
        Set login = New ADODB.Recordset
        login.CursorType = adOpenDynamic
        login.CursorLocation = adUseClient
        login.LockType = adLockOptimistic
        login.Open "Select * from Users where UserName='" & txtUsername.Text & "'", conn, login.CursorType, login.LockType, adCmdUnknown
        
        If login.EOF Then
            MessageBar.Caption = "ERROR: No such user exists! Please check for spelling errors."
            txtUsername.SetFocus
            Timer1.Enabled = True
            Exit Sub
        Else
            If login.Fields("Password") = txtPassword.Text Then
                Me.Hide
                HomeView.Show
                usrLogin login.Fields("Username"), login.Fields("Password")
                txtUsername.Text = ""
                txtPassword.Text = ""
            Else
                MessageBar.Caption = "ERROR: Wrong password! Please check for spelling/capitalization errors."
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

Private Sub Form_Activate()
    txtUsername.SetFocus
    exitVal = vbNo
End Sub

Private Sub Form_Initialize()
    Token = InitGDIPlus
    C = Me.BackColor
    If C < 0 Then C = GetSysColor(C - &H80000000)
End Sub

Private Sub Form_Load()
    ShopLogo.Picture = LoadPictureGDIPlus(App.Path & "\Images\logo.png", 100, 80, &HADADAD, True)

    cmdLogin.BackColor = &H8000000D
    cmdCancel.BackColor = &H8000000D

    Timer1.Enabled = False
    exitVal = vbNo
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdLogin.BackColor = &H8000000D
    cmdCancel.BackColor = &H8000000D
    cmdLogin.FontItalic = False
    cmdCancel.FontItalic = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    exitVal = MsgBox("Are you sure you want to exit the application?", vbYesNo + vbDefaultButton2 + vbQuestion + vbApplicationModal, "Confirm Exit")
    If exitVal = vbYes Then
        FreeGDIPlus Token
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub Timer1_Timer()
    MessageBar.FontBold = Not MessageBar.FontBold
End Sub
