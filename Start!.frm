VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000E&
   Caption         =   "AWord·开始"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13440
   ForeColor       =   &H80000008&
   Icon            =   "Start!.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   7935
   ScaleWidth      =   13440
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "    跳过>>"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1920
      Left            =   120
      Picture         =   "Start!.frx":19D02
      Top             =   2400
      Width           =   3225
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   120
      Picture         =   "Start!.frx":1AA08
      Top             =   120
      Width           =   3225
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Form1.Show
Unload Me
End Sub
Private Sub Image2_Click()
OpenFileWindow.Show
End Sub
Private Sub Label1_Click()
Label1.ForeColor = &H8000000D
Form1.Show
Unload Me
End Sub
Private Sub Form_Resize()
Label1.Top = Form3.Height - 1000
End Sub
Private Sub Form_Load()
Label1.ForeColor = vbBlack
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.ForeColor = &H8000000D Then
Label1.ForeColor = vbBlack
End If
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &H8000000D
End Sub
