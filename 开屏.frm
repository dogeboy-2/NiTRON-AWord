VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4005
   ClientLeft      =   7545
   ClientTop       =   3390
   ClientWidth     =   7200
   ClipControls    =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "开屏.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7200
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   4290
      Left            =   0
      TabIndex        =   0
      Top             =   -240
      Width           =   7200
      Begin VB.Timer Timer4 
         Left            =   5040
         Top             =   2640
      End
      Begin VB.Timer Timer3 
         Left            =   6120
         Top             =   1680
      End
      Begin VB.Timer Timer2 
         Left            =   4800
         Top             =   1320
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1800
         Left            =   2640
         Picture         =   "开屏.frx":19D02
         ScaleHeight     =   1264.2
         ScaleMode       =   0  'User
         ScaleWidth      =   1264.2
         TabIndex        =   4
         Top             =   960
         Width           =   1800
      End
      Begin VB.Timer Timer1 
         Left            =   6120
         Top             =   3360
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AuroraWord"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2310
         TabIndex        =   3
         Top             =   2760
         Width           =   2580
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "AuroraStudio 极光软创"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "正在加载 AuroraWord.."
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
         Left            =   240
         TabIndex        =   1
         Top             =   3720
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Timer1.Enabled = True
    Timer1.Interval = 800
End Sub
Private Sub Timer1_Timer()
Label1.Caption = "正在加载 AuroraWord..."
Timer4.Enabled = True
Timer4.Interval = 300
Timer1.Enabled = False
End Sub
Private Sub Timer2_Timer()
Label1.Caption = "加载完成"
Timer3.Enabled = True
Timer3.Interval = 100
End Sub
Private Sub Timer3_Timer()
Unload Me
Form3.Show
End Sub
Private Sub Timer4_Timer()
Label1.Caption = "正在加载 AuroraWord....."
Timer2.Enabled = True
Timer2.Interval = 400
End Sub
