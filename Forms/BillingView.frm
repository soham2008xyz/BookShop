VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form BillingView 
   AutoRedraw      =   -1  'True
   Caption         =   "Billing"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   Picture         =   "BillingView.frx":0000
   ScaleHeight     =   9030
   ScaleWidth      =   15150
   StartUpPosition =   2  'CenterScreen
   Begin BookShop.Frameset OptionsFrame 
      Height          =   1575
      Left            =   7800
      TabIndex        =   14
      Top             =   1560
      Width           =   6975
      _extentx        =   12303
      _extenty        =   2778
      font            =   "BillingView.frx":46D2
      caption         =   "// Add to Cart"
      forecolor       =   192
      forecolordisabled=   13977088
      xp7_bordercolor =   -2147483635
      transparent     =   -1  'True
      Begin VB.ComboBox BookQty 
         BeginProperty Font 
            Name            =   "Roboto Light"
            Size            =   12
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2520
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label MessageBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please add items to your cart."
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
         Left            =   360
         TabIndex        =   18
         Top             =   1080
         Width           =   6255
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
         Left            =   4200
         MousePointer    =   10  'Up Arrow
         TabIndex        =   17
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblSelectQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Quantity:"
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
         Left            =   600
         TabIndex        =   15
         Top             =   480
         Width           =   1650
      End
   End
   Begin BookShop.Frameset BillFrame 
      Height          =   5175
      Left            =   7800
      TabIndex        =   13
      Top             =   3360
      Width           =   6975
      _extentx        =   12303
      _extenty        =   9128
      font            =   "BillingView.frx":46FE
      caption         =   "// Cart Details"
      forecolor       =   192
      forecolordisabled=   13977088
      xp7_bordercolor =   -2147483635
      transparent     =   -1  'True
      Begin VB.ListBox BillList 
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
         ItemData        =   "BillingView.frx":472A
         Left            =   240
         List            =   "BillingView.frx":472C
         TabIndex        =   19
         Top             =   480
         Width           =   6495
      End
      Begin VB.Label cmdProceed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proceed >>"
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
         Left            =   3720
         MousePointer    =   10  'Up Arrow
         TabIndex        =   21
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label cmdDelete 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Delete"
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
         Left            =   1200
         MousePointer    =   10  'Up Arrow
         TabIndex        =   20
         Top             =   4320
         Width           =   2175
      End
   End
   Begin BookShop.Frameset PriceFrame 
      Height          =   2175
      Left            =   480
      TabIndex        =   6
      Top             =   6360
      Width           =   7095
      _extentx        =   12515
      _extenty        =   3836
      font            =   "BillingView.frx":472E
      caption         =   "// Book Details"
      forecolor       =   192
      xp7_bordercolor =   -2147483635
      transparent     =   -1  'True
      Begin VB.Label txtPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select a book from the list"
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
         Left            =   1920
         TabIndex        =   12
         Top             =   1680
         Width           =   2745
      End
      Begin VB.Label txtQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select a book from the list"
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
         Left            =   1920
         TabIndex        =   11
         Top             =   1080
         Width           =   2745
      End
      Begin VB.Label txtAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select a book from the list"
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
         Left            =   1920
         TabIndex        =   10
         Top             =   480
         Width           =   2745
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price:"
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
         TabIndex        =   9
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label lblQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty. in Stock:"
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
         TabIndex        =   8
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label lblAuthor 
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
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1440
      End
   End
   Begin BookShop.Frameset BookFrame 
      Height          =   4575
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   7095
      _extentx        =   12515
      _extenty        =   8070
      font            =   "BillingView.frx":475A
      framestyle      =   13
      caption         =   "// Select Book"
      forecolor       =   192
      forecolordisabled=   0
      xp7_bordercolor =   -2147483635
      transparent     =   -1  'True
      Begin MSDataListLib.DataList BookContainer 
         Bindings        =   "BillingView.frx":4786
         Height          =   3840
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6773
         _Version        =   393216
         ListField       =   "BOOKNAME"
         BoundColumn     =   ""
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
   End
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
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   9615
      FormWidthDT     =   15390
      FormScaleHeightDT=   9030
      FormScaleWidthDT=   15150
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
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
      Top             =   240
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   480
      X2              =   14760
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
      Left            =   480
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
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   180
   End
   Begin VB.Label OptionLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please select books for billing:"
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
      Width           =   2940
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
Dim billCount As Integer

Private Sub BookContainer_Click()
    If BookContainer.SelectedItem = Null Then
        Exit Sub
    Else
        BookList.Recordset.AbsolutePosition = BookContainer.SelectedItem
        txtAuthor.Caption = BookList.Recordset.Fields("AUTHORNAME")
        txtQty.Caption = BookList.Recordset.Fields("QTY")
        txtPrice.Caption = "Rs. " & BookList.Recordset.Fields("MRP")
        
        BookQty.Clear
        BookQty.Enabled = False
        Dim q, i As Integer
        q = Val(BookList.Recordset.Fields("QTY"))
        If (q > 0) Then
            For i = 1 To q
                BookQty.AddItem i
            Next i
            BookQty.Enabled = True
            txtQty.ForeColor = &H0&
        Else
            BookQty.AddItem 0
            txtQty.ForeColor = &HC0&
        End If
    End If
End Sub

Private Sub cmdAdd_Click()
    If Val(BookQty.Text) >= 1 Then
        If Val(BookQty.Text) <= Val(BookList.Recordset.Fields("QTY")) Then
            Select Case MsgBox("Available in stock = " & BookList.Recordset.Fields("QTY") & vbNewLine & "Remaining in stock = " & CStr(Int(Val(BookList.Recordset.Fields("QTY"))) - (BookQty.ListIndex + 1)) & vbNewLine & "Add to cart?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton1, "Confirm add")
                Case vbYes
                    BookList.Recordset.Fields("QTY") = (Int(Val(BookList.Recordset.Fields("QTY"))) - (BookQty.ListIndex + 1))
                    BookList.Recordset.Update
                    txtQty.Caption = BookList.Recordset.Fields("QTY")
                    MessageBar.Caption = "'" & BookList.Recordset.Fields("BOOKNAME") & "' added to cart!"
                    billCount = billCount + 1
                    BillList.AddItem BookList.Recordset.Fields("BOOKNAME") & " - Rs. " & BookList.Recordset.Fields("MRP") & " x " & (BookQty.ListIndex + 1)
                    BillList.ItemData(billCount - 1) = Val(BookList.Recordset.Fields("ID"))
                    
                    cmdDelete.Enabled = True
                    cmdProceed.Enabled = True
                    cmdDelete.BackColor = &H8000000D
                    cmdProceed.BackColor = &H8000000D
                    
                    BookQty.Clear
                    BookQty.Enabled = False
                    Dim q, i As Integer
                    q = Val(BookList.Recordset.Fields("QTY"))
                    If (q > 0) Then
                        For i = 1 To q
                            BookQty.AddItem i
                        Next i
                        BookQty.Enabled = True
                    Else
                        BookQty.AddItem 0
                        txtQty.ForeColor = &HC0&
                    End If
                Case vbNo
                    MessageBar.Caption = "Cart not updated! Please review cart."
            End Select
        Else
            MsgBox "The quantity is more than available stock!", vbApplicationModal + vbOKOnly + vbExclamation, "Please select valid quantity"
        End If
    Else
        MsgBox "Please select a valid quantity first!", vbApplicationModal + vbOKOnly + vbExclamation, "Please select quantity"
    End If
End Sub

Private Sub cmdDelete_Click()
    If BillList.ListCount <= 0 Then
        MsgBox "Your cart is empty. Add items to cart first!", vbApplicationModal + vbOKOnly + vbExclamation, "No items to delete"
    Else
        If BillList.ListIndex < 0 Then
            MsgBox "Please select the item you want to delete!", vbApplicationModal + vbOKOnly + vbExclamation, "Select item to delete"
        Else
            Dim pos As Integer
            pos = InStrRev(BillList.Text, "x", , vbTextCompare)
            
            Select Case MsgBox("Are you sure you want to remove '" & Left$(BillList.Text, (InStrRev(BillList.Text, "-", , vbTextCompare) - 2)) & "' from your cart?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton1, "Confirm delete")
                Case vbYes
                    conn.Execute "UPDATE BookList SET  QTY = QTY + " & Mid$(BillList.Text, pos + 2) & " WHERE ID = " & BillList.ItemData(BillList.ListIndex) & ";"
                    BookList.RecordSource = "BookList"
                    BookList.Refresh
                    
                    MessageBar.Caption = "'" & Left$(BillList.Text, (InStrRev(BillList.Text, "-", , vbTextCompare) - 2)) & "' removed from cart."
                    BillList.RemoveItem BillList.ListIndex
                    billCount = billCount - 1
                    If BillList.ListCount <= 0 Then
                        cmdDelete.Enabled = False
                        cmdProceed.Enabled = False
                        cmdDelete.BackColor = QBColor(8)
                        cmdProceed.BackColor = QBColor(8)
                    End If
                Case vbNo
                    MessageBar.Caption = "Item not removed! Please review cart."
            End Select
        End If
    End If
End Sub

Private Sub cmdProceed_Click()
    Select Case MsgBox("Proceed to payment?", vbApplicationModal + vbYesNo + vbQuestion + vbDefaultButton1, "Are you sure?")
        Case vbYes
            Dim i As Integer
            Dim amt As Double
            amt = 0
            For i = 0 To (BillList.ListCount - 1)
                PrintView.txtBill.Text = PrintView.txtBill.Text & (i + 1) & vbTab
                If Len(Left$(BillList.List(i), (InStrRev(BillList.List(i), "-", , vbTextCompare) - 2))) > 32 Then
                    PrintView.txtBill.Text = PrintView.txtBill.Text & Left$(Left$(BillList.List(i), (InStrRev(BillList.List(i), "-", , vbTextCompare) - 2)), 32) & "..." & vbTab & vbTab
                Else
                    PrintView.txtBill.Text = PrintView.txtBill.Text & Left$(BillList.List(i), (InStrRev(BillList.List(i), "-", , vbTextCompare) - 2)) & String$(42 - Len(Left$(BillList.List(i), (InStrRev(BillList.List(i), "-", , vbTextCompare) - 2))), " ") & vbTab & vbTab
                End If
                PrintView.txtBill.Text = PrintView.txtBill.Text & Mid$(BillList.List(i), InStrRev(BillList.List(i), "x", , vbTextCompare) + 2) & vbTab & Mid$(BillList.List(i), InStrRev(BillList.List(i), ".", , vbTextCompare) + 2, (InStrRev(BillList.List(i), "x", , vbTextCompare) - 2) - (InStrRev(BillList.List(i), ".", , vbTextCompare) + 2) + 1) & vbTab & Val(Mid$(BillList.List(i), InStrRev(BillList.List(i), "x", , vbTextCompare) + 2)) * Val(Mid$(BillList.List(i), InStrRev(BillList.List(i), ".", , vbTextCompare) + 2, (InStrRev(BillList.List(i), "x", , vbTextCompare) - 2) - (InStrRev(BillList.List(i), ".", , vbTextCompare) + 2) + 1)) & vbNewLine
                amt = amt + (Val(Mid$(BillList.List(i), InStrRev(BillList.List(i), "x", , vbTextCompare) + 2)) * Val(Mid$(BillList.List(i), InStrRev(BillList.List(i), ".", , vbTextCompare) + 2, (InStrRev(BillList.List(i), "x", , vbTextCompare) - 2) - (InStrRev(BillList.List(i), ".", , vbTextCompare) + 2) + 1)))
            Next i
            PrintView.OptionLabel.Caption = "Bill for " & billCount & " item(s)"
            PrintView.txtAmt.Caption = "Rs. " & amt
            billCount = 0
            BillList.Clear
            cmdDelete.Enabled = False
            cmdProceed.Enabled = False
            cmdDelete.BackColor = QBColor(8)
            cmdProceed.BackColor = QBColor(8)
            Me.Hide
            PrintView.Show
        Case vbNo
            Exit Sub
    End Select
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
   BookFrame.AutoReDraw = True
   BookFrame.ReDraw
   PriceFrame.AutoReDraw = True
   PriceFrame.ReDraw
   OptionsFrame.AutoReDraw = True
   OptionsFrame.ReDraw
   BillFrame.AutoReDraw = True
   BillFrame.ReDraw
   
   BookQty.Clear
   BookQty.AddItem 0
   BookQty.Enabled = False
   cmdDelete.Enabled = False
   cmdProceed.Enabled = False
   cmdDelete.BackColor = QBColor(8)
   cmdProceed.BackColor = QBColor(8)
   MessageBar.Caption = "Please add items to your cart."
   
   exitVal = vbNo
   billCount = 0
   
   Me.Width = Me.Width + 10
End Sub

Private Sub Form_Resize()
    BookFrame.ReDraw
    PriceFrame.ReDraw
    OptionsFrame.ReDraw
    BillFrame.ReDraw
End Sub

Private Sub Form_Unload(Cancel As Integer)
    exitVal = MsgBox("Go back to main screen?", vbYesNo + vbDefaultButton2 + vbQuestion + vbApplicationModal, "Confirm Exit")
    If exitVal = vbYes Then
        HomeView.Show
        FreeGDIPlus Token
        Books.Close
        Unload Me
    Else
        Cancel = 1
    End If
End Sub

