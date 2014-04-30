VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox InventoryIcon 
      Height          =   1695
      Left            =   3720
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   1635
      ScaleWidth      =   2835
      TabIndex        =   11
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
         TabIndex        =   12
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
      TabIndex        =   10
      Tag             =   "no_resize"
      ToolTipText     =   "Go to Main Screen"
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox BillingIcon 
      Height          =   1695
      Left            =   480
      MouseIcon       =   "BillingView.frx":0000
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   1635
      ScaleWidth      =   2895
      TabIndex        =   8
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
         TabIndex        =   9
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
      TabIndex        =   6
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
         TabIndex        =   7
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
      TabIndex        =   4
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
         TabIndex        =   5
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
      TabIndex        =   2
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
         TabIndex        =   3
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.PictureBox ExitIcon 
      Height          =   1695
      Left            =   6960
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   1635
      ScaleWidth      =   2835
      TabIndex        =   0
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
         TabIndex        =   1
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   480
      X2              =   9840
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   960
      Width           =   2325
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

