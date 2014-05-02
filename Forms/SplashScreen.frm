VERSION 5.00
Begin VB.Form SplashView 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8370
   FillColor       =   &H8000000D&
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7680
      Top             =   3360
   End
   Begin VB.Label MessageBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Starting application..."
      BeginProperty Font 
         Name            =   "Roboto Condensed Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   6615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "xBookShop"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   20.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1290
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1290
      Left            =   1440
      TabIndex        =   1
      Top             =   2400
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1290
      Left            =   960
      TabIndex        =   0
      Top             =   2400
      Width           =   210
   End
End
Attribute VB_Name = "SplashView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dShort As Integer
Dim dLong, dLong2, dLong3 As Integer
Dim animCount, animCount2, animCount3 As Integer
Dim doneAnim, animStarted2, animStarted3 As Boolean
Dim loadTime As Long

Private Sub Form_Load()
    doneAnim = False
    animCount = 0
    dShort = 100
    dLong = 300
    dLong2 = 300
    dLong3 = 300

    animStarted2 = False
    animStarted3 = False
    Timer1.Enabled = True
    Label3.Left = 0
    Label3.Visible = False
    Label2.Left = 0
    Label2.Visible = False
    Label1.Left = 0
    Label1.Visible = False
    Label4.FontBold = True

    loadTime = 0
    initUser
    initDB
End Sub

Private Sub Timer1_Timer()
    If animCount = 0 Then
        Label1.Visible = True
        Label1.Left = 0
        animCount = animCount + 1
    End If
    If animCount > 0 And animCount <= 8 Then
        Label1.Left = Label1.Left + dLong
        animCount = animCount + 1
        dLong = dLong + 10
    End If
    If animCount > 8 And animCount <= 25 Then
        Label1.Left = Label1.Left + dShort
        animCount = animCount + 1
    End If
    If animCount > 25 And animCount <= 38 Then
        dLong = dLong - 10
        Label1.Left = Label1.Left + dLong
        animCount = animCount + 1
    End If
    If animCount > 38 Then
        Label1.Visible = False
        Label1.Left = 0
        animCount = 0
        dLong = 300
    End If

    animCount2 = (animCount - 7)
    If animCount2 < 0 Then animCount2 = animCount2 + 39
    
    If animCount2 = 0 Then
        Label2.Visible = True
        Label2.Left = 0
        dLong2 = 300
        animStarted2 = True
    End If
    If animCount2 > 0 And animCount2 <= 8 And animStarted2 Then
        Label2.Left = Label2.Left + dLong2
        dLong2 = dLong2 + 20
    End If
    If animCount2 > 8 And animCount2 <= 25 And animStarted2 Then
        Label2.Left = Label2.Left + dShort
    End If
    If animCount2 > 25 And animCount2 <= 38 And animStarted2 Then
        dLong2 = dLong2 - 10
        Label2.Left = Label2.Left + dLong2
    End If
    If animCount2 > 38 Then
        Label2.Visible = False
        Label2.Left = 0
    End If

    animCount3 = (animCount - 16)
    If animCount3 < 0 Then animCount3 = animCount3 + 39

    If animCount3 = 0 Then
        Label3.Visible = True
        Label3.Left = 0
        dLong3 = 300
        animStarted3 = True
    End If
    If animCount3 > 0 And animCount3 <= 8 And animStarted3 Then
        Label3.Left = Label3.Left + dLong3
        dLong3 = dLong3 + 5
    End If
    If animCount3 > 8 And animCount3 <= 25 And animStarted3 Then
        Label3.Left = Label3.Left + dShort
    End If
    If animCount3 > 25 And animCount3 <= 38 And animStarted3 Then
        dLong3 = dLong3 - 5
        Label3.Left = Label3.Left + dLong3
    End If
    If animCount3 > 38 Then
        Label3.Visible = False
        Label3.Left = 0
    End If

    loadTime = loadTime + Timer1.Interval
    If loadTime = 1000 Then MessageBar.Caption = "Loading resources..."
    If loadTime = 2500 Then MessageBar.Caption = "Initializing database connection..."
    If loadTime = 4500 Then MessageBar.Caption = "All done!"
    
    If loadTime > 5000 Then
        Timer1.Enabled = False
        Unload Me
        LoginView.Show
    End If
End Sub
