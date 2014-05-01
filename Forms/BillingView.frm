VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form BillingView 
   Caption         =   "Form1"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   15150
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc BookList 
      Height          =   330
      Left            =   4800
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Book List"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
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
      Height          =   4695
      Left            =   6960
      TabIndex        =   5
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Frame BookFrame 
      Caption         =   "Select Book"
      BeginProperty Font 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   600
      TabIndex        =   4
      Top             =   1440
      Width           =   6015
      Begin MSDataListLib.DataList DataList1 
         Bindings        =   "BillingView.frx":0000
         Height          =   3570
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   6297
         _Version        =   393216
         ListField       =   "BOOKNAME"
         BoundColumn     =   ""
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
      FormHeightDT    =   8970
      FormWidthDT     =   15390
      FormScaleHeightDT=   8385
      FormScaleWidthDT=   15150
   End
   Begin VB.PictureBox ShopLogo 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   14040
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
      X2              =   14760
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
Option Explicit
Dim Token As Long
Dim C As Long
Dim exitVal As Integer

    
Private Sub DataList1_Click()
BookList.Recordset.AbsolutePosition = DataList1.SelectedItem
Debug.Print "Selected: " & BookList.Recordset.Fields("BOOKNAME") & " by " & BookList.Recordset.Fields("AUTHORNAME")

End Sub

Private Sub Form_Initialize()
    Token = InitGDIPlus
    C = Me.BackColor
    If C < 0 Then C = GetSysColor(C - &H80000000)

    BookList.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\BookDB.mdb;Persist Security Info=False"
    BookList.CursorLocation = adUseClient
    BookList.CursorType = adOpenDynamic
    BookList.CommandType = adCmdTable
    BookList.RecordSource = "BookList"
    BookList.Refresh
    
End Sub

Private Sub Form_Load()
    ShopLogo.Picture = LoadPictureGDIPlus(App.Path & "\Images\logo.png", Me.Width / 592, Me.Height / 318, C, False)
End Sub

Private Sub Form_Resize()
    ShopLogo.Picture = LoadPictureGDIPlus(App.Path & "\Images\logo.png", Me.Width / 592, Me.Height / 318, C, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
FreeGDIPlus Token
End Sub

