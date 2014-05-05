VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form PrintView 
   AutoRedraw      =   -1  'True
   Caption         =   "Print Bill"
   ClientHeight    =   9165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   Picture         =   "PrintBill.frx":0000
   ScaleHeight     =   9165
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCustomer 
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
      Left            =   2640
      TabIndex        =   9
      Top             =   1920
      Width           =   5295
   End
   Begin VB.TextBox txtBill 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3120
      Width           =   9615
   End
   Begin VB.PictureBox ShopLogo 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9360
      ScaleHeight     =   495
      ScaleWidth      =   735
      TabIndex        =   0
      Tag             =   "no_resize"
      ToolTipText     =   "Students Book House"
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
      FormHeightDT    =   9750
      FormWidthDT     =   10710
      FormScaleHeightDT=   9165
      FormScaleWidthDT=   10470
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label cmdPrint 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Print"
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
      Left            =   2880
      MousePointer    =   10  'Up Arrow
      TabIndex        =   15
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Label cmdProceed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proceed >>"
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
      Left            =   5400
      MousePointer    =   10  'Up Arrow
      TabIndex        =   14
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Label txtAmt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   13
      Top             =   7680
      Width           =   60
   End
   Begin VB.Label lblAmt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "// Total Bill Amount:"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   12
      Top             =   7680
      Width           =   2115
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   360
      X2              =   9960
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblBill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "// Bill Details:"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   11
      Top             =   2640
      Width           =   1410
   End
   Begin VB.Label lblCustomer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name:"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   10
      Top             =   2040
      Width           =   1755
   End
   Begin VB.Label txtTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "09:57:01"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9120
      TabIndex        =   8
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8400
      TabIndex        =   7
      Top             =   1440
      Width           =   570
   End
   Begin VB.Label txtDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "02-05-2014"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   1440
      Width           =   1230
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   540
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
      TabIndex        =   3
      Top             =   240
      Width           =   3375
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
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   180
   End
   Begin VB.Label OptionLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   60
   End
End
Attribute VB_Name = "PrintView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Token As Long
Dim C As Long

Private Sub cmdPrint_Click()
    cmdPrint.Visible = False
    cmdProceed.Visible = False
    PrintForm
    cmdPrint.Visible = True
    cmdProceed.Visible = True
End Sub

Private Sub cmdProceed_Click()
    Unload Me
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPrint.BackColor = &H80&
    cmdProceed.BackColor = &H8000000D
    cmdPrint.FontItalic = True
    cmdProceed.FontItalic = False
End Sub

Private Sub cmdProceed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPrint.BackColor = &H8000000D
    cmdProceed.BackColor = &H80&
    cmdPrint.FontItalic = False
    cmdProceed.FontItalic = True
End Sub

Private Sub Form_Load()
    Token = InitGDIPlus
    C = Me.BackColor
    If C < 0 Then C = GetSysColor(C - &H80000000)
    
    ShopLogo.Picture = LoadPictureGDIPlus(App.Path & "\Images\logo.png", 100, 80, &HADADAD, True)
    
    txtBill.Text = "Item No." & vbTab & "Name" & vbTab & vbTab & vbTab & vbTab & vbTab & "Qty." & vbTab & "Rate" & vbTab & "Price" & vbNewLine
    txtBill.Text = txtBill.Text & String$(70, "=") & vbNewLine
    txtDate.Caption = DateValue(Now)
    txtTime.Caption = TimeValue(Now)
    lblDate.FontBold = True
    lblTime.FontBold = True
    lblBill.FontBold = True
    lblAmt.FontBold = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPrint.BackColor = &H8000000D
    cmdProceed.BackColor = &H8000000D
    cmdPrint.FontItalic = False
    cmdProceed.FontItalic = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case MsgBox("Go back to home screen?", vbApplicationModal + vbYesNo + vbQuestion + vbDefaultButton1, "Sure to proceed?")
        Case vbYes
            Unload Me
            FreeGDIPlus Token
            HomeView.Show
        Case vbNo
            Cancel = 1
    End Select
End Sub

