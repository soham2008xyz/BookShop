VERSION 5.00
Begin VB.Form WelcomeForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   3600
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "About this software"
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   2880
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Go to Inventory"
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Go to Billing "
      CausesValidation=   0   'False
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Picture         =   "Welcome.frx":0000
      TabIndex        =   1
      Top             =   1440
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9840
      ScaleHeight     =   495
      ScaleWidth      =   735
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   240
      X2              =   10575
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Please select an option:"
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Students Book House"
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   18
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3330
   End
End
Attribute VB_Name = "WelcomeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Dim flag As Integer

Private Sub Form_Initialize()
 flag = 0
 Dim Token As Long
 Dim C As Long
    
 C = Me.BackColor
 If C < 0 Then C = GetSysColor(C - &H80000000)

 Token = InitGDIPlus
 Picture1.Picture = LoadPictureGDIPlus(App.Path & "\Images\logo.png", 35, 35, C, False)
 
 
 FreeGDIPlus Token
End Sub

Private Sub Form_Unload(Cancel As Integer)
HomeView.Show

End Sub

Private Sub Timer1_Timer()
If Me.WindowState = vbMinimized And flag = 0 Then
MsgBox "Minimized"
flag = 1
End If
If Me.WindowState <> vbMinimized Then
flag = 0
End If
End Sub
