VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form InventoryView 
   AutoRedraw      =   -1  'True
   Caption         =   "View Inventory"
   ClientHeight    =   8580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15030
   LinkTopic       =   "Form1"
   Picture         =   "InventoryView.frx":0000
   ScaleHeight     =   8580
   ScaleWidth      =   15030
   StartUpPosition =   3  'Windows Default
   Begin BookShop.Frameset DetailsFrame 
      Height          =   4935
      Left            =   7680
      TabIndex        =   6
      Top             =   2760
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8705
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   192
      ForeColorDisabled=   13977088
      XP7_BorderColor =   -2147483635
      Caption         =   "// Book Details"
      Transparent     =   -1  'True
      Begin VB.TextBox txtCategory 
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
         Left            =   2520
         TabIndex        =   28
         Top             =   3480
         Width           =   4095
      End
      Begin VB.TextBox txtPrice 
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
         Left            =   2520
         TabIndex        =   27
         Top             =   2880
         Width           =   4095
      End
      Begin VB.TextBox txtQty 
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
         Left            =   2520
         TabIndex        =   26
         Top             =   2280
         Width           =   4095
      End
      Begin VB.TextBox txtISBN 
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
         Left            =   2520
         TabIndex        =   25
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox txtAuthorName 
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
         Left            =   2520
         TabIndex        =   24
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox txtBookName 
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
         Left            =   2520
         TabIndex        =   23
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label cmdDelete 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Delete"
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
         Left            =   2760
         MousePointer    =   10  'Up Arrow
         TabIndex        =   33
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label lblCategory 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category Name:"
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
         Left            =   360
         TabIndex        =   22
         Top             =   3600
         Width           =   1680
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price per Unit:"
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
         Left            =   360
         TabIndex        =   21
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity in Stock:"
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
         Left            =   360
         TabIndex        =   20
         Top             =   2400
         Width           =   1860
      End
      Begin VB.Label lblISBN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ISBN:"
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
         Left            =   360
         TabIndex        =   19
         Top             =   1800
         Width           =   570
      End
      Begin VB.Label lblAuthorName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author Name:"
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
         Left            =   360
         TabIndex        =   18
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label lblBookName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Name:"
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
         Left            =   360
         TabIndex        =   17
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label cmdCancel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         Enabled         =   0   'False
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
         TabIndex        =   12
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label cmdSave 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Save"
         Enabled         =   0   'False
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
         Left            =   4200
         MousePointer    =   10  'Up Arrow
         TabIndex        =   11
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label cmdEdit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Edit"
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
         Left            =   1560
         MousePointer    =   10  'Up Arrow
         TabIndex        =   10
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label cmdAdd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add"
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
         Left            =   240
         MousePointer    =   10  'Up Arrow
         TabIndex        =   9
         Top             =   4200
         Width           =   1095
      End
   End
   Begin BookShop.Frameset ListFrame 
      Height          =   4935
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8705
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   192
      ForeColorDisabled=   13977088
      XP7_BorderColor =   -2147483635
      Caption         =   "// Query results"
      Transparent     =   -1  'True
      Begin VB.ListBox BookResults 
         BeginProperty Font 
            Name            =   "Roboto Light"
            Size            =   12
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3525
         ItemData        =   "InventoryView.frx":46D2
         Left            =   240
         List            =   "InventoryView.frx":46D4
         TabIndex        =   7
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label cmdLast 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ">>"
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
         Left            =   5160
         MousePointer    =   10  'Up Arrow
         TabIndex        =   16
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label cmdNext 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ">"
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
         Left            =   3720
         MousePointer    =   10  'Up Arrow
         TabIndex        =   15
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label cmdPrev 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<"
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
         Left            =   2280
         MousePointer    =   10  'Up Arrow
         TabIndex        =   14
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label cmdFirst 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<<"
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
         Left            =   840
         MousePointer    =   10  'Up Arrow
         TabIndex        =   13
         Top             =   4200
         Width           =   1095
      End
   End
   Begin BookShop.Frameset SearchFrame 
      Height          =   1215
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   2143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   192
      ForeColorDisabled=   13977088
      XP7_BorderColor =   -2147483635
      Caption         =   "// Filter by Name"
      Transparent     =   -1  'True
      Begin VB.TextBox txtSearch 
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
         Left            =   2160
         TabIndex        =   8
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label cmdBack 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<< Back"
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
         Left            =   12000
         MousePointer    =   10  'Up Arrow
         TabIndex        =   32
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label cmdReset 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reset"
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
         Left            =   9720
         MousePointer    =   10  'Up Arrow
         TabIndex        =   31
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label cmdFilter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Filter"
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
         Left            =   7440
         MousePointer    =   10  'Up Arrow
         TabIndex        =   30
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblSearch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text to Search:"
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
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   1560
      End
   End
   Begin VB.PictureBox ShopLogo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   14040
      ScaleHeight     =   495
      ScaleWidth      =   735
      TabIndex        =   0
      Tag             =   "no_resize"
      ToolTipText     =   "Go to Main Screen"
      Top             =   120
      Width           =   735
   End
   Begin MSAdodcLib.Adodc BookList 
      Height          =   330
      Left            =   4680
      Top             =   240
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
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   3960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1366
      ScreenHeightDT  =   768
      ScreenWidthDT   =   1366
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   9165
      FormWidthDT     =   15270
      FormScaleHeightDT=   8580
      FormScaleWidthDT=   15030
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin MSComctlLib.StatusBar StatusView 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   34
      Top             =   8160
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18891
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
      Top             =   120
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
      Top             =   840
      Width           =   180
   End
   Begin VB.Label OptionLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book Inventory"
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
      Top             =   840
      Width           =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   360
      X2              =   14640
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "InventoryView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Token As Long
Dim C As Long
Dim LastI As Integer
Dim flagNew As Boolean

Private Sub BookResults_Click()
    BookList.Recordset.AbsolutePosition = BookResults.ListIndex + 1
    txtBookName.Text = BookList.Recordset.Fields("BOOKNAME")
    txtAuthorName.Text = BookList.Recordset.Fields("AUTHORNAME")
    txtISBN.Text = BookList.Recordset.Fields("ISBN")
    txtCategory.Text = BookList.Recordset.Fields("CATEGORY")
    txtQty.Text = BookList.Recordset.Fields("QTY")
    txtPrice.Text = BookList.Recordset.Fields("MRP")
    StatusView.Panels(1).Text = "Book " & (BookResults.ListIndex + 1) & " of " & BookResults.ListCount & " selected"
End Sub

Private Sub BookResults_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub cmdAdd_Click()
    txtBookName.Text = ""
    txtAuthorName.Text = ""
    txtISBN.Text = ""
    txtCategory.Text = ""
    txtQty.Text = ""
    txtPrice.Text = ""
    
    txtBookName.Enabled = True
    txtAuthorName.Enabled = True
    txtISBN.Enabled = True
    txtCategory.Enabled = True
    txtQty.Enabled = True
    txtPrice.Enabled = True
    
    cmdAdd.Enabled = False
    cmdAdd.BackColor = QBColor(8)
    cmdEdit.Enabled = False
    cmdEdit.BackColor = QBColor(8)
    cmdDelete.Enabled = False
    cmdDelete.BackColor = QBColor(8)
    cmdSave.Enabled = True
    cmdSave.BackColor = &H8000000D
    cmdCancel.Enabled = True
    cmdCancel.BackColor = &H8000000D
    
    flagNew = True
    BookList.Recordset.AddNew
    txtBookName.SetFocus
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H80&
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = True
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H80&
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = True
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub cmdCancel_Click()
    If flagNew Then
        BookList.Recordset.Cancel
        BookList.Recordset.CancelUpdate
        BookResults.ListIndex = 0
        flagNew = False
    End If

    cmdAdd.Enabled = True
    cmdAdd.BackColor = &H8000000D
    cmdEdit.Enabled = True
    cmdEdit.BackColor = &H8000000D
    cmdDelete.Enabled = True
    cmdDelete.BackColor = &H8000000D
    cmdSave.Enabled = False
    cmdSave.BackColor = QBColor(8)
    cmdCancel.Enabled = False
    cmdCancel.BackColor = QBColor(8)
    
    txtBookName.Enabled = False
    txtAuthorName.Enabled = False
    txtISBN.Enabled = False
    txtCategory.Enabled = False
    txtQty.Enabled = False
    txtPrice.Enabled = False
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H80&
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = True
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub cmdDelete_Click()
    Dim index As Integer
    Select Case MsgBox("Are you sure you want to delete this book from the inventory?", vbApplicationModal + vbYesNo + vbQuestion + vbDefaultButton1, "Confirm delete")
        Case vbYes
            index = BookResults.ListIndex - 1
            BookList.Recordset.Delete
            BookList.Recordset.Update
            BookResults.RemoveItem BookResults.ListIndex
            If index >= 0 Then
                BookResults.ListIndex = index
            Else
                BookResults.ListIndex = 0
            End If
    End Select
    StatusView.Panels(2).Text = BookResults.ListCount & " books in database"
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H80&
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = True
End Sub

Private Sub cmdEdit_Click()
    txtBookName.Enabled = True
    txtAuthorName.Enabled = True
    txtISBN.Enabled = True
    txtCategory.Enabled = True
    txtQty.Enabled = True
    txtPrice.Enabled = True
    
    cmdAdd.Enabled = False
    cmdAdd.BackColor = QBColor(8)
    cmdEdit.Enabled = False
    cmdEdit.BackColor = QBColor(8)
    cmdDelete.Enabled = False
    cmdDelete.BackColor = QBColor(8)
    cmdSave.Enabled = True
    cmdSave.BackColor = &H8000000D
    cmdCancel.Enabled = True
    cmdCancel.BackColor = &H8000000D
    
    flagNew = False
    txtBookName.SetFocus
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H80&
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = True
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub cmdFilter_Click()
    If txtSearch.Text = "" Then
        MsgBox "Enter text to search for!", vbApplicationModal + vbOKOnly + vbInformation, "Enter search term"
        txtSearch.SetFocus
        Exit Sub
    End If
    
    Dim i As Integer
    For i = LastI To BookResults.ListCount - 1
        If InStr(1, BookResults.List(i), txtSearch.Text, vbTextCompare) > 0 Then
            BookResults.ListIndex = i
            LastI = i
            Exit Sub
        End If
    Next i
    MsgBox "'" & txtSearch.Text & "' not found in the records!", vbApplicationModal + vbOKOnly + vbInformation, "Not found"
End Sub

Private Sub cmdFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H80&
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = True
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub cmdFirst_Click()
    BookResults.ListIndex = 0
End Sub

Private Sub cmdFirst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H80&
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = True
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub cmdLast_Click()
    BookResults.ListIndex = BookResults.ListCount - 1
End Sub

Private Sub cmdLast_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H80&
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = True
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub cmdNext_Click()
    If BookResults.ListIndex = BookResults.ListCount - 1 Then
        MsgBox "No records after this!", vbApplicationModal + vbOKOnly + vbInformation, "End of file reached"
    Else
        BookResults.ListIndex = BookResults.ListIndex + 1
    End If
End Sub

Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H80&
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = True
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub cmdPrev_Click()
    If BookResults.ListIndex = 0 Then
        MsgBox "No records before this!", vbApplicationModal + vbOKOnly + vbInformation, "Start of file reached"
    Else
        BookResults.ListIndex = BookResults.ListIndex - 1
    End If
End Sub

Private Sub cmdPrev_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H80&
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = True
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub cmdReset_Click()
    LastI = 0
    BookResults.ListIndex = 0
    txtSearch.Text = ""
End Sub

Private Sub cmdReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H80&
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = True
    cmdDelete.FontItalic = False
End Sub

Private Sub cmdSave_Click()
    If txtBookName.Text = "" Then
        MsgBox "Book name cannot be blank!", vbApplicationModal + vbExclamation + vbOKOnly, "Error"
        txtBookName.SetFocus
        Exit Sub
    End If
    If txtAuthorName.Text = "" Then
        MsgBox "Author name cannot be blank!", vbApplicationModal + vbExclamation + vbOKOnly, "Error"
        txtAuthorName.SetFocus
        Exit Sub
    End If
    If txtISBN.Text = "" Then
        MsgBox "ISBN cannot be blank!", vbApplicationModal + vbExclamation + vbOKOnly, "Error"
        txtISBN.SetFocus
        Exit Sub
    End If
    If txtCategory.Text = "" Then
        MsgBox "Category cannot be blank!", vbApplicationModal + vbExclamation + vbOKOnly, "Error"
        txtCategory.SetFocus
        Exit Sub
    End If
    If txtQty.Text = "" Then
        MsgBox "Quantity cannot be blank!", vbApplicationModal + vbExclamation + vbOKOnly, "Error"
        txtQty.SetFocus
        Exit Sub
    End If
    If txtPrice.Text = "" Then
        MsgBox "Price cannot be blank!", vbApplicationModal + vbExclamation + vbOKOnly, "Error"
        txtPrice.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtQty.Text) Then
        MsgBox "Quantity must be a number!", vbApplicationModal + vbExclamation + vbOKOnly, "Error"
        txtQty.Text = ""
        txtQty.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtPrice.Text) Then
        MsgBox "Price must be a number!", vbApplicationModal + vbExclamation + vbOKOnly, "Error"
        txtPrice.Text = ""
        txtPrice.SetFocus
        Exit Sub
    End If
    BookList.Recordset.Fields("BOOKNAME") = txtBookName.Text
    BookList.Recordset.Fields("AUTHORNAME") = txtAuthorName.Text
    BookList.Recordset.Fields("ISBN") = txtISBN.Text
    BookList.Recordset.Fields("CATEGORY") = txtCategory.Text
    BookList.Recordset.Fields("QTY") = txtQty.Text
    BookList.Recordset.Fields("MRP") = txtPrice.Text
    BookList.Recordset.Update
    
    If flagNew Then
        flagNew = False
        BookResults.AddItem txtBookName.Text
        BookResults.ListIndex = BookResults.ListCount - 1
    End If
    
    cmdAdd.Enabled = True
    cmdAdd.BackColor = &H8000000D
    cmdEdit.Enabled = True
    cmdEdit.BackColor = &H8000000D
    cmdDelete.Enabled = True
    cmdDelete.BackColor = &H8000000D
    cmdSave.Enabled = False
    cmdSave.BackColor = QBColor(8)
    cmdCancel.Enabled = False
    cmdCancel.BackColor = QBColor(8)
    
    txtBookName.Enabled = False
    txtAuthorName.Enabled = False
    txtISBN.Enabled = False
    txtCategory.Enabled = False
    txtQty.Enabled = False
    txtPrice.Enabled = False
    StatusView.Panels(2).Text = BookResults.ListCount & " books in database"
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H80&
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = True
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub Form_Activate()
    StatusView.Panels(2).Text = BookResults.ListCount & " books in database"
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
    ShopLogo.Picture = LoadPictureGDIPlus(App.Path & "\Images\logo.png", 100, 80, &HADADAD, True)
    
    SearchFrame.AutoReDraw = True
    SearchFrame.ReDraw
    ListFrame.AutoReDraw = True
    DetailsFrame.ReDraw
    DetailsFrame.AutoReDraw = True
    DetailsFrame.ReDraw
    Me.Width = Me.Width + 20
    
    txtBookName.Enabled = False
    txtAuthorName.Enabled = False
    txtISBN.Enabled = False
    txtCategory.Enabled = False
    txtQty.Enabled = False
    txtPrice.Enabled = False
    
    cmdAdd.Enabled = True
    cmdAdd.BackColor = &H8000000D
    cmdEdit.Enabled = True
    cmdEdit.BackColor = &H8000000D
    cmdDelete.Enabled = True
    cmdDelete.BackColor = &H8000000D
    cmdSave.Enabled = False
    cmdSave.BackColor = QBColor(8)
    cmdCancel.Enabled = False
    cmdCancel.BackColor = QBColor(8)
    
    BookList.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\BookDB.mdb;Persist Security Info=False"
    BookList.CursorLocation = adUseClient
    BookList.CursorType = adOpenDynamic
    BookList.CommandType = adCmdTable
    BookList.RecordSource = "BookList"
    BookList.Refresh
    
    BookList.Recordset.MoveFirst
    While Not BookList.Recordset.EOF
        BookResults.AddItem BookList.Recordset.Fields("BOOKNAME")
        BookList.Recordset.MoveNext
    Wend
    
    BookResults.ListIndex = 0
    BookList.Recordset.AbsolutePosition = BookResults.ListIndex + 1
    LastI = 0
    flagNew = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub Form_Resize()
    SearchFrame.ReDraw
    ListFrame.ReDraw
    DetailsFrame.ReDraw
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case MsgBox("Go back to home screen?", vbApplicationModal + vbYesNo + vbQuestion + vbDefaultButton1, "Sure to exit?")
        Case vbYes
            FreeGDIPlus Token
            HomeView.Show
            Unload Me
        Case vbNo
            Cancel = 1
    End Select
End Sub

Private Sub lblAuthorName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub lblBookName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub lblCategory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub lblISBN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub lblPrice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub lblQty_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub lblSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub txtAuthorName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub txtBookName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub txtCategory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub txtISBN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub txtPrice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub txtQty_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub

Private Sub txtSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled Then cmdAdd.BackColor = &H8000000D
    If cmdEdit.Enabled Then cmdEdit.BackColor = &H8000000D
    If cmdSave.Enabled Then cmdSave.BackColor = &H8000000D
    If cmdCancel.Enabled Then cmdCancel.BackColor = &H8000000D
    cmdNext.BackColor = &H8000000D
    cmdPrev.BackColor = &H8000000D
    cmdFirst.BackColor = &H8000000D
    cmdLast.BackColor = &H8000000D
    cmdBack.BackColor = &H8000000D
    cmdFilter.BackColor = &H8000000D
    cmdReset.BackColor = &H8000000D
    If cmdDelete.Enabled Then cmdDelete.BackColor = &H8000000D
    cmdAdd.FontItalic = False
    cmdEdit.FontItalic = False
    cmdSave.FontItalic = False
    cmdCancel.FontItalic = False
    cmdNext.FontBold = False
    cmdPrev.FontBold = False
    cmdFirst.FontBold = False
    cmdLast.FontBold = False
    cmdBack.FontItalic = False
    cmdFilter.FontItalic = False
    cmdReset.FontItalic = False
    cmdDelete.FontItalic = False
End Sub
