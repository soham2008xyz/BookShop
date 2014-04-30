VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form HomeView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Home"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "WelcomeForm.frx":0000
   ScaleHeight     =   6105
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusView 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   5730
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10319
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   706
            MinWidth        =   706
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox ExitIcon 
      Height          =   1695
      Left            =   6960
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   1635
      ScaleWidth      =   2835
      TabIndex        =   9
      Top             =   3600
      Width           =   2895
      Begin VB.Label ExitOption 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "// Logout"
         BeginProperty Font 
            Name            =   "Roboto Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   0
         TabIndex        =   15
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.PictureBox ReportsIcon 
      Height          =   1695
      Left            =   6960
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   1635
      ScaleWidth      =   2835
      TabIndex        =   8
      Top             =   1560
      Width           =   2895
      Begin VB.Label ReportsOption 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "// View Reports"
         BeginProperty Font 
            Name            =   "Roboto Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   0
         TabIndex        =   12
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.PictureBox AboutIcon 
      Height          =   1695
      Left            =   3720
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   1635
      ScaleWidth      =   2835
      TabIndex        =   7
      Top             =   3600
      Width           =   2895
      Begin VB.Label AboutOption 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "// About Application"
         BeginProperty Font 
            Name            =   "Roboto Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   0
         TabIndex        =   14
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.PictureBox CatalogIcon 
      Height          =   1695
      Left            =   480
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   1635
      ScaleWidth      =   2835
      TabIndex        =   5
      Top             =   3600
      Width           =   2895
      Begin VB.Label CatalogOption 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "// View Catalog"
         BeginProperty Font 
            Name            =   "Roboto Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   495
         Left            =   0
         TabIndex        =   13
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.PictureBox BillingIcon 
      Height          =   1695
      Left            =   480
      MouseIcon       =   "WelcomeForm.frx":1CCA
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   1635
      ScaleWidth      =   2895
      TabIndex        =   4
      ToolTipText     =   "Go to Billing Screen"
      Top             =   1560
      Width           =   2955
      Begin VB.Label BillingOption 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "// Go to Billing"
         BeginProperty Font 
            Name            =   "Roboto Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   0
         TabIndex        =   10
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.PictureBox ShopLogo 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9240
      ScaleHeight     =   495
      ScaleWidth      =   735
      TabIndex        =   0
      Tag             =   "no_resize"
      ToolTipText     =   "Go to Main Screen"
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox InventoryIcon 
      Height          =   1695
      Left            =   3720
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   1635
      ScaleWidth      =   2835
      TabIndex        =   6
      Top             =   1560
      Width           =   2895
      Begin VB.Label InventoryOption 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "// Go to Inventory"
         BeginProperty Font 
            Name            =   "Roboto Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   1200
         Width           =   2895
      End
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
      TabIndex        =   3
      Top             =   960
      Width           =   180
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
      TabIndex        =   2
      Top             =   240
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   480
      X2              =   9840
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "HomeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Dim Token As Long
Dim C As Long
Dim exitVal As Integer

Private Sub BillingIcon_Click()
Me.Hide
WelcomeForm.Show
End Sub

Private Sub ExitIcon_Click()
'exitVal = MsgBox("Are you sure you want to log out?", vbYesNo + vbDefaultButton2 + vbInformation + vbApplicationModal, "Confirm Exit")
'If exitVal = vbYes Then
'    usrLogout
'    LoginView.Show
'    Unload Me
'End If
Unload Me
End Sub

Private Sub ExitOption_Click()
'exitVal = MsgBox("Are you sure you want to log out?", vbYesNo + vbDefaultButton2 + vbInformation + vbApplicationModal, "Confirm Exit")
'If exitVal = vbYes Then End
'    usrLogout
'    LoginView.Show
'    Unload Me
'End If
Unload Me
End Sub

Private Sub Form_Activate()
'StatusView.Style = sbrSimple
'StatusView.SimpleText = "Hello " & username & "!"
StatusView.Panels(1).Text = "Hello " & username & "!"
StatusView.Panels(3).Text = "Last logged in at " & lastLogin
StatusView.Panels(2).Picture = LoadPictureGDIPlus(App.Path & "\Images\clock.png", 30, 20, C, True)

exitVal = vbNo
End Sub

Private Sub Form_Load()
Token = InitGDIPlus
C = Me.BackColor
If C < 0 Then C = GetSysColor(C - &H80000000)
 
ShopLogo.Picture = LoadPictureGDIPlus(App.Path & "\Images\logo.png", 35, 35, C, True)
BillingIcon.Picture = LoadPictureGDIPlus(App.Path & "\Images\billing.png", 200, 270, vbWhite, True)
InventoryIcon.Picture = LoadPictureGDIPlus(App.Path & "\Images\inventory.jpg", 200, 270, C, True)
ReportsIcon.Picture = LoadPictureGDIPlus(App.Path & "\Images\reports.jpg", 200, 270, C, True)
CatalogIcon.Picture = LoadPictureGDIPlus(App.Path & "\Images\catalog.jpg", 200, 270, C, True)
AboutIcon.Picture = LoadPictureGDIPlus(App.Path & "\Images\about.png", 200, 270, C, True)
ExitIcon.Picture = LoadPictureGDIPlus(App.Path & "\Images\exit.jpg", 200, 270, C, True)

BillingOption.BackColor = &H8000000D
InventoryOption.BackColor = &H8000000D
ReportsOption.BackColor = &H8000000D
CatalogOption.BackColor = &H8000000D
AboutOption.BackColor = &H8000000D
ExitOption.BackColor = &H8000000D

exitVal = vbNo

End Sub

Private Sub Form_Unload(Cancel As Integer)
exitVal = MsgBox("Are you sure you want to log out?", vbYesNo + vbDefaultButton2 + vbInformation + vbApplicationModal, "Confirm Exit")
If exitVal = vbYes Then
    usrLogout
    LoginView.Show
    FreeGDIPlus Token
    Unload Me
Else
    Cancel = 1
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'ShopLogo.Picture = LoadPictureGDIPlus(App.Path & "\Images\logo.png", 35, 35, C, True)

BillingOption.BackColor = &H8000000D
InventoryOption.BackColor = &H8000000D
ReportsOption.BackColor = &H8000000D
CatalogOption.BackColor = &H8000000D
AboutOption.BackColor = &H8000000D
ExitOption.BackColor = &H8000000D

BillingOption.FontItalic = False
InventoryOption.FontItalic = False
ReportsOption.FontItalic = False
CatalogOption.FontItalic = False
AboutOption.FontItalic = False
ExitOption.FontItalic = False

End Sub

Private Sub BillingOption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
BillingOption.BackColor = &H80&
InventoryOption.BackColor = &H8000000D
ReportsOption.BackColor = &H8000000D
CatalogOption.BackColor = &H8000000D
AboutOption.BackColor = &H8000000D
ExitOption.BackColor = &H8000000D

BillingOption.FontItalic = True
InventoryOption.FontItalic = False
ReportsOption.FontItalic = False
CatalogOption.FontItalic = False
AboutOption.FontItalic = False
ExitOption.FontItalic = False

End Sub

Private Sub BillingIcon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
BillingOption.BackColor = &H80&
InventoryOption.BackColor = &H8000000D
ReportsOption.BackColor = &H8000000D
CatalogOption.BackColor = &H8000000D
AboutOption.BackColor = &H8000000D
ExitOption.BackColor = &H8000000D

BillingOption.FontItalic = True
InventoryOption.FontItalic = False
ReportsOption.FontItalic = False
CatalogOption.FontItalic = False
AboutOption.FontItalic = False
ExitOption.FontItalic = False

End Sub

Private Sub InventoryOption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
BillingOption.BackColor = &H8000000D
InventoryOption.BackColor = &H80&
ReportsOption.BackColor = &H8000000D
CatalogOption.BackColor = &H8000000D
AboutOption.BackColor = &H8000000D
ExitOption.BackColor = &H8000000D

BillingOption.FontItalic = False
InventoryOption.FontItalic = True
ReportsOption.FontItalic = False
CatalogOption.FontItalic = False
AboutOption.FontItalic = False
ExitOption.FontItalic = False

End Sub

Private Sub InventoryIcon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
BillingOption.BackColor = &H8000000D
InventoryOption.BackColor = &H80&
ReportsOption.BackColor = &H8000000D
CatalogOption.BackColor = &H8000000D
AboutOption.BackColor = &H8000000D
ExitOption.BackColor = &H8000000D

BillingOption.FontItalic = False
InventoryOption.FontItalic = True
ReportsOption.FontItalic = False
CatalogOption.FontItalic = False
AboutOption.FontItalic = False
ExitOption.FontItalic = False

End Sub

Private Sub ReportsOption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
BillingOption.BackColor = &H8000000D
InventoryOption.BackColor = &H8000000D
ReportsOption.BackColor = &H80&
CatalogOption.BackColor = &H8000000D
AboutOption.BackColor = &H8000000D
ExitOption.BackColor = &H8000000D

BillingOption.FontItalic = False
InventoryOption.FontItalic = False
ReportsOption.FontItalic = True
CatalogOption.FontItalic = False
AboutOption.FontItalic = False
ExitOption.FontItalic = False

End Sub

Private Sub ReportsIcon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
BillingOption.BackColor = &H8000000D
InventoryOption.BackColor = &H8000000D
ReportsOption.BackColor = &H80&
CatalogOption.BackColor = &H8000000D
AboutOption.BackColor = &H8000000D
ExitOption.BackColor = &H8000000D

BillingOption.FontItalic = False
InventoryOption.FontItalic = False
ReportsOption.FontItalic = True
CatalogOption.FontItalic = False
AboutOption.FontItalic = False
ExitOption.FontItalic = False

End Sub

Private Sub CatalogOption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
BillingOption.BackColor = &H8000000D
InventoryOption.BackColor = &H8000000D
ReportsOption.BackColor = &H8000000D
CatalogOption.BackColor = &H80&
AboutOption.BackColor = &H8000000D
ExitOption.BackColor = &H8000000D

BillingOption.FontItalic = False
InventoryOption.FontItalic = False
ReportsOption.FontItalic = False
CatalogOption.FontItalic = True
AboutOption.FontItalic = False
ExitOption.FontItalic = False

End Sub

Private Sub CatalogIcon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
BillingOption.BackColor = &H8000000D
InventoryOption.BackColor = &H8000000D
ReportsOption.BackColor = &H8000000D
CatalogOption.BackColor = &H80&
AboutOption.BackColor = &H8000000D
ExitOption.BackColor = &H8000000D

BillingOption.FontItalic = False
InventoryOption.FontItalic = False
ReportsOption.FontItalic = False
CatalogOption.FontItalic = True
AboutOption.FontItalic = False
ExitOption.FontItalic = False

End Sub

Private Sub AboutOption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
BillingOption.BackColor = &H8000000D
InventoryOption.BackColor = &H8000000D
ReportsOption.BackColor = &H8000000D
CatalogOption.BackColor = &H8000000D
AboutOption.BackColor = &H80&
ExitOption.BackColor = &H8000000D

BillingOption.FontItalic = False
InventoryOption.FontItalic = False
ReportsOption.FontItalic = False
CatalogOption.FontItalic = False
AboutOption.FontItalic = True
ExitOption.FontItalic = False

End Sub

Private Sub AboutIcon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
BillingOption.BackColor = &H8000000D
InventoryOption.BackColor = &H8000000D
ReportsOption.BackColor = &H8000000D
CatalogOption.BackColor = &H8000000D
AboutOption.BackColor = &H80&
ExitOption.BackColor = &H8000000D

BillingOption.FontItalic = False
InventoryOption.FontItalic = False
ReportsOption.FontItalic = False
CatalogOption.FontItalic = False
AboutOption.FontItalic = True
ExitOption.FontItalic = False

End Sub

Private Sub ExitOption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
BillingOption.BackColor = &H8000000D
InventoryOption.BackColor = &H8000000D
ReportsOption.BackColor = &H8000000D
CatalogOption.BackColor = &H8000000D
AboutOption.BackColor = &H8000000D
ExitOption.BackColor = &H80&

BillingOption.FontItalic = False
InventoryOption.FontItalic = False
ReportsOption.FontItalic = False
CatalogOption.FontItalic = False
AboutOption.FontItalic = False
ExitOption.FontItalic = True

End Sub

Private Sub ExitIcon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
BillingOption.BackColor = &H8000000D
InventoryOption.BackColor = &H8000000D
ReportsOption.BackColor = &H8000000D
CatalogOption.BackColor = &H8000000D
AboutOption.BackColor = &H8000000D
ExitOption.BackColor = &H80&

BillingOption.FontItalic = False
InventoryOption.FontItalic = False
ReportsOption.FontItalic = False
CatalogOption.FontItalic = False
AboutOption.FontItalic = False
ExitOption.FontItalic = True

End Sub

'Private Sub ShopLogo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'ShopLogo.Picture = LoadPictureGDIPlus(App.Path & "\Images\logo.png", 35, 35, vbBlue, True)
'End Sub
