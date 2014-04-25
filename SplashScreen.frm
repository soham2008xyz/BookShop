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
         Name            =   "Roboto Light"
         Size            =   27.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   300
      TabIndex        =   2
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   27.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   900
      TabIndex        =   1
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   27.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   2760
      Width           =   135
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

Private Sub Form_Load()
doneAnim = False
animCount = 0
'animCount2 = 0
'animCount3 = 0

dShort = 100
dLong = 300
dLong2 = 300
dLong3 = 300

animStarted2 = False
animStarted3 = False
Timer1.Enabled = True
Label3.left = 0
Label3.Visible = False
Label2.left = 0
Label2.Visible = False
Label1.left = 0
Label1.Visible = False
Label4.FontBold = True

End Sub

Private Sub Timer1_Timer()
'If Not doneAnim Then
    If animCount = 0 Then
        'Timer1.Enabled = True
        Label1.Visible = True
        Label1.left = 0
        'Label2.Visible = True
        'Label2.left = Label1.left + 10
        animCount = animCount + 1
    End If
    If animCount > 0 And animCount <= 8 Then
        Label1.left = Label1.left + dLong
        animCount = animCount + 1
        dLong = dLong + 10
    End If
    If animCount > 8 And animCount <= 25 Then
        Label1.left = Label1.left + dShort
        animCount = animCount + 1
    End If
    If animCount > 25 And animCount <= 38 Then
        dLong = dLong - 10
        Label1.left = Label1.left + dLong
        animCount = animCount + 1
    End If
    If animCount > 38 Then
        'Timer1.Enabled = False
        'doneAnim = True
        Label1.Visible = False
        Label1.left = 0
        animCount = 0
        dLong = 300
    End If
'End If

animCount2 = (animCount - 7)
If animCount2 < 0 Then animCount2 = animCount2 + 39
'Label2.Caption = animCount2

'If Not doneAnim Then
    If animCount2 = 0 Then
        'Timer1.Enabled = True
        Label2.Visible = True
        Label2.left = 0
        dLong2 = 300
        animStarted2 = True
        'Label2.Visible = True
        'Label2.left = Label1.left + 50
        'animCount = animCount + 1
    End If
    If animCount2 > 0 And animCount2 <= 8 And animStarted2 Then
        Label2.left = Label2.left + dLong2
        'animCount = animCount + 1
        dLong2 = dLong2 + 20
    End If
    If animCount2 > 8 And animCount2 <= 25 And animStarted2 Then
        Label2.left = Label2.left + dShort
        'animCount = animCount + 1
    End If
    If animCount2 > 25 And animCount2 <= 38 And animStarted2 Then
        dLong2 = dLong2 - 10
        Label2.left = Label2.left + dLong2
        'animCount = animCount + 1
    End If
    If animCount2 > 38 Then
        'Timer1.Enabled = False
        'doneAnim = True
        Label2.Visible = False
        Label2.left = 0
        'animCount = 0
        'dLong2 = 300
    End If
'End If

animCount3 = (animCount - 16)
If animCount3 < 0 Then animCount3 = animCount3 + 39

'If Not doneAnim Then
    If animCount3 = 0 Then
        'Timer1.Enabled = True
        Label3.Visible = True
        Label3.left = 0
        dLong3 = 300
        animStarted3 = True
        'Label2.Visible = True
        'Label2.left = Label1.left + 50
        'animCount = animCount + 1
    End If
    If animCount3 > 0 And animCount3 <= 8 And animStarted3 Then
        Label3.left = Label3.left + dLong3
        'animCount = animCount + 1
        dLong3 = dLong3 + 5
    End If
    If animCount3 > 8 And animCount3 <= 25 And animStarted3 Then
        Label3.left = Label3.left + dShort
        'animCount = animCount + 1
    End If
    If animCount3 > 25 And animCount3 <= 38 And animStarted3 Then
        dLong3 = dLong3 - 5
        Label3.left = Label3.left + dLong3
        'animCount = animCount + 1
    End If
    If animCount3 > 38 Then
        'Timer1.Enabled = False
        'doneAnim = True
        Label3.Visible = False
        Label3.left = 0
        'animCount = 0
        'dLong2 = 300
    End If
'End If

End Sub
