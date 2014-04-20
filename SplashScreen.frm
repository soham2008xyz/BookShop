VERSION 5.00
Begin VB.Form Form1 
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
      Interval        =   50
      Left            =   1080
      Top             =   1440
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "..."
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
      Height          =   615
      Left            =   1100
      TabIndex        =   0
      Top             =   2520
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dShort As Integer
Dim dLong As Integer

Dim animCount As Integer
Dim doneAnim As Boolean

Private Sub Form_Load()
doneAnim = False
animCount = 0

dShort = 100
dLong = 300
End Sub

Private Sub Timer1_Timer()
If Not doneAnim Then
    If animCount = 0 Then
        Label1.Visible = True
        Label1.left = 250
        animCount = animCount + 1
    ElseIf animCount <= 20 Then
        Label1.left = Label1.left + dShort
        animCount = animCount + 1
        dShort = dShort + 1
    ElseIf animCount <= 30 Then
        Label1.left = Label1.left + dLong
        animCount = animCount + 1
    ElseIf animCont <= 50 Then
        Label1.left = Label1.left + dShort
        animCount = animCount + 1
        dShort = dShort - 1
    Else
        'Timer1.Enabled = False
        Label1.Visible = False
        animCount = 0
    End If
End If
End Sub
