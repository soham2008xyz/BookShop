VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form BillingView 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame PreviewFrame 
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   9960
      TabIndex        =   5
      Top             =   1560
      Width           =   9735
   End
   Begin VB.Frame BillingFrame 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   9255
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   720
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   5520
         Width           =   5175
      End
      Begin VB.ListBox List1 
         Columns         =   1
         BeginProperty Font 
            Name            =   "Roboto Light"
            Size            =   12
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4155
         ItemData        =   "BillingView.frx":0000
         Left            =   360
         List            =   "BillingView.frx":0002
         TabIndex        =   6
         Top             =   960
         Width           =   5295
      End
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   4080
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1366
      ScreenHeightDT  =   768
      ScreenWidthDT   =   1366
      FormHeightDT    =   6585
      FormWidthDT     =   10590
      FormScaleHeightDT=   6000
      FormScaleWidthDT=   10350
   End
   Begin VB.PictureBox ShopLogo 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   19080
      ScaleHeight     =   495
      ScaleWidth      =   735
      TabIndex        =   0
      Tag             =   "no_resize"
      ToolTipText     =   "Go to Main Screen"
      Top             =   240
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   480
      X2              =   19680
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
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   3375
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
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   180
   End
   Begin VB.Label OptionLabel 
      AutoSize        =   -1  'True
      Caption         =   "Please select an option:"
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
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   2325
   End
End
Attribute VB_Name = "BillingView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Option Explicit
Dim Token As Long
Dim C As Long
Dim exitVal As Integer

Private Sub DataList1_Click()
'Debug.Print DataList1.SelectedItem
'Debug.Print Adodc1.Recordset.AbsolutePosition
Adodc1.Recordset.AbsolutePosition = DataList1.SelectedItem
Debug.Print "Selected: " & Adodc1.Recordset.Fields("BookName") & " by " & Adodc1.Recordset.Fields("AuthorName")

End Sub

Private Sub Form_Load()
Token = InitGDIPlus
C = Me.BackColor
If C < 0 Then C = GetSysColor(C - &H80000000)
 
ShopLogo.Picture = LoadPictureGDIPlus(App.Path & "\Images\logo.png", Me.Width / 592, Me.Height / 318, C, False)
End Sub

Private Sub Form_Resize()
ShopLogo.Picture = LoadPictureGDIPlus(App.Path & "\Images\logo.png", Me.Width / 592, Me.Height / 318, C, False)
'Cls
'Print Me.Width
'Print Me.Height
'Print Me.Width / 35
'Print Me.Height / 35
End Sub

Private Sub Form_Unload(Cancel As Integer)
FreeGDIPlus Token
End Sub

