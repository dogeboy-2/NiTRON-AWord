VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000E&
   Caption         =   "AWord"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17565
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "文件Form.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8070
   ScaleWidth      =   17565
   Begin VB.PictureBox Picture7 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   3120
      ScaleHeight     =   7095
      ScaleWidth      =   15855
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   15855
      Begin VB.TextBox Text5 
         Height          =   465
         Left            =   2520
         TabIndex        =   35
         Top             =   2640
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1800
         Left            =   360
         Picture         =   "文件Form.frx":19D02
         ScaleHeight     =   1264.2
         ScaleMode       =   0  'User
         ScaleWidth      =   1264.2
         TabIndex        =   26
         Top             =   1080
         Width           =   1800
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000E&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   840
         Left            =   2590
         TabIndex        =   34
         ToolTipText     =   "WYSISWYG!"
         Top             =   1440
         Width           =   490
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "What_Damon"
         Height          =   495
         Left            =   10080
         TabIndex        =   30
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Redmountain2018"
         Height          =   495
         Left            =   10080
         TabIndex        =   29
         Top             =   1080
         Width           =   2085
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "苏方华-RTC"
         Height          =   615
         Left            =   10080
         TabIndex        =   28
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "鸣谢：（以首字母为序）"
         Height          =   495
         Left            =   9360
         TabIndex        =   27
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "  Preview 3 (UI and function packs)"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   2880
         TabIndex        =   25
         Top             =   3960
         Width           =   7215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "当前版本："
         Height          =   2055
         Left            =   2640
         TabIndex        =   24
         Top             =   3240
         Width           =   4215
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AuroraStudio 极光软创"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   2520
         TabIndex        =   23
         Top             =   480
         Width           =   4650
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H8000000E&
         Caption         =   "uroraWord"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   840
         Left            =   3120
         TabIndex        =   22
         ToolTipText     =   "WYSIWYG!"
         Top             =   1440
         Width           =   5445
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   3360
      ScaleHeight     =   5655
      ScaleWidth      =   8775
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CommandButton Command4 
         Caption         =   "取消"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         TabIndex        =   19
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "确定"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         TabIndex        =   18
         Top             =   5040
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   17
         Text            =   ".doc"
         Top             =   4440
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   4440
         Width           =   5535
      End
      Begin VB.FileListBox File2 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3405
         Left            =   2880
         TabIndex        =   15
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   120
         Width           =   5535
      End
      Begin VB.DriveListBox Drive2 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   2415
      End
      Begin VB.DirListBox Dir2 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3510
         Left            =   0
         TabIndex        =   12
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   3000
      ScaleHeight     =   7335
      ScaleWidth      =   9495
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   9495
      Begin VB.CommandButton Command2 
         Caption         =   "取消"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         TabIndex        =   9
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "确定"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         TabIndex        =   8
         Top             =   5640
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   7
         Text            =   "*.fns"
         Top             =   4800
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   4800
         Width           =   5535
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3345
         Left            =   3120
         TabIndex        =   5
         Top             =   960
         Width           =   5295
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3510
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   5535
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "注意：如果您输入的文件不存在，我们会在当前目录下自动为您创建使用该名字的新的空白文件。"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   10
         Top             =   5640
         Width           =   4455
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   100
      Picture         =   "文件Form.frx":1BC64
      ScaleHeight     =   10455
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   0
      Width           =   2500
      Begin VB.Image Label4 
         Height          =   360
         Left            =   120
         Picture         =   "文件Form.frx":1EC73
         Top             =   9650
         Width           =   360
      End
      Begin VB.Image Image4 
         Height          =   360
         Left            =   480
         Picture         =   "文件Form.frx":1F1AF
         Top             =   4410
         Width           =   360
      End
      Begin VB.Image Image3 
         Height          =   360
         Left            =   480
         Picture         =   "文件Form.frx":1F74E
         Top             =   2960
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   480
         Picture         =   "文件Form.frx":1FCA9
         Top             =   2205
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   480
         Picture         =   "文件Form.frx":201C9
         Top             =   1403
         Width           =   360
      End
      Begin VB.Image Picture2 
         Height          =   480
         Left            =   240
         Picture         =   "文件Form.frx":206F2
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Picture3 
         BackStyle       =   0  'Transparent
         Caption         =   "新建"
         ForeColor       =   &H80000007&
         Height          =   615
         Left            =   1060
         TabIndex        =   33
         Top             =   1475
         Width           =   975
      End
      Begin VB.Label Picture4 
         BackStyle       =   0  'Transparent
         Caption         =   "打开"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1060
         TabIndex        =   32
         Top             =   2250
         Width           =   870
      End
      Begin VB.Label Command5 
         BackStyle       =   0  'Transparent
         Caption         =   "保存"
         ForeColor       =   &H80000007&
         Height          =   495
         Left            =   1060
         TabIndex        =   31
         Top             =   3000
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000016&
         FillColor       =   &H80000016&
         Height          =   460
         Left            =   220
         Top             =   1200
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000016&
         X1              =   120
         X2              =   2280
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "另存为"
         Height          =   495
         Left            =   1065
         TabIndex        =   20
         Top             =   4440
         Width           =   1935
      End
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   2640
      X2              =   2640
      Y1              =   0
      Y2              =   8040
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileName As String
Dim Wertern As String
Dim Northern As String
Dim fuck As String
Dim Xa As String, Ya As String, FClick As String
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xa = X
Ya = Y
FClick = "Yes"
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xa = ""
Ya = ""
FClick = "No"
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FClick = "Yes" Then
Me.Left = Me.Left - Xa + X
Me.Top = Me.Top - Ya + Y
End If
End Sub
Private Sub Command5_Click()
If Not Shape1.Visible = True Then
Shape1.Visible = True
Shape1.Top = Command5.Top
Else
Shape1.Top = Command5.Top
End If
If Not Form1.RichTextBox1.FileName = "" And Not Form1.RichTextBox1.FileName = "未命名" Then
Form1.RichTextBox1.SaveFile Form1.RichTextBox1.FileName
Form1.Caption = "AWord:" & Form1.RichTextBox1.FileName
Else
Picture5.Visible = False
Picture6.Visible = True
Picture7.Visible = False
End If
End Sub
Private Sub Form_Resize()
    Picture1.Height = Form2.Height
    Line3.Y2 = Form2.Height
    Label4.Top = Me.Height - Label4.Height - Label4.Height - Label4.Height
End Sub
Private Sub Image1_Click()
Picture3_Click
End Sub
Private Sub Image2_Click()
Picture4_Click
End Sub
Private Sub Image3_Click()
Command5_Click
End Sub

Private Sub Image5_Click()

End Sub

Private Sub Label11_Click()
Text5.Visible = True
End Sub

Private Sub Label2_Click()
If Not Shape1.Visible = True Then
Shape1.Visible = True
Shape1.Top = Label2.Top
Else
Shape1.Top = Label2.Top
End If
Picture5.Visible = False
Picture6.Visible = True
Picture7.Visible = False
End Sub
Private Sub Label4_Click()
If Not Shape1.Visible = True Then
Shape1.Visible = True
Shape1.Top = Label4.Top
Else
Shape1.Top = Label4.Top
End If
Picture7.Visible = Not Picture7.Visible
Picture5.Visible = False
Picture6.Visible = False
End Sub

Private Sub Picture2_Click()
Unload Me
Form1.Show
End Sub
Private Sub Picture3_Click()
If Not Shape1.Visible = True Then
Shape1.Visible = True
Shape1.Top = Picture3.Top
Else
Shape1.Top = Picture3.Top
End If
Form1.RichTextBox1.text = "" '清空文本框
FileName = "未命名"
Form1.Caption = "AWord:" & FileName
Unload Me
Form1.Show
End Sub
Private Sub Picture4_Click()
If Not Shape1.Visible = True Then
Shape1.Visible = True
Shape1.Top = Picture4.Top
Else
Shape1.Top = Picture4.Top
End If
Picture6.Visible = False
Picture5.Visible = True
Picture7.Visible = False
Form1.Caption = "AWord:" & Form1.RichTextBox1.FileName
End Sub
Private Sub Combo1_Click()
File1.Pattern = Combo1.text
End Sub
Private Sub Command1_Click()
If Not Text2.text = "" And Not Text2.text = " " And Not File1.FileName = "" And Not File1.FileName = " " Then
Wertern = File1.Path & "\" & Text2.text
Dim Mazmun As String
Dim strFileName As String
strFileName = Wertern
Dim A
Set A = CreateObject("ADODB.Stream")
A.Charset = "utf-8"
A.open
A.LoadFromFile strFileName
Mazmun = A.ReadText
A.Close
Dim stm
Set stm = CreateObject("adodb.stream")
stm.Type = 2
stm.Mode = 3
stm.Charset = "gb2312"
stm.open
stm.WriteText Mazmun
stm.SaveToFile strFileName, 2
stm.flush
stm.Close
Set stm = Nothing
Form1.RichTextBox1.FileName = Wertern
Form1.RichTextBox1.LoadFile Form1.RichTextBox1.FileName
Form1.Caption = "AWord:" + Form1.RichTextBox1.FileName
Form1.Show
Unload Me
Else
MsgBox "不能不选择文件或文件名为空", vbInformation, "提示"
End If
End Sub
Private Sub Command2_Click()
Unload Me
Form1.Show
End Sub
Private Sub Dir1_Click()
File1.Path = Dir1.Path
Text1.text = Dir1.Path & "\"
End Sub
Private Sub Drive1_change()
Dir1.Path = Drive1.Drive
End Sub
Private Sub File1_Click()
Text2.text = File1.FileName
End Sub
Private Sub File1_DblClick()
If Not Text2.text = "" And Not Text2.text = " " And Not File1.FileName = "" And Not File1.FileName = " " Then
Wertern = File1.Path & "\" & Text2.text
Dim Mazmun As String
Dim strFileName As String
strFileName = Wertern
Dim A
Set A = CreateObject("ADODB.Stream")
A.Charset = "utf-8"
A.open
A.LoadFromFile strFileName
Mazmun = A.ReadText
A.Close
Dim stm
Set stm = CreateObject("adodb.stream")
stm.Type = 2
stm.Mode = 3
stm.Charset = "gb2312"
stm.open
stm.WriteText Mazmun
stm.SaveToFile strFileName, 2
stm.flush
stm.Close
Set stm = Nothing
Form1.RichTextBox1.FileName = Wertern
Form1.RichTextBox1.LoadFile Form1.RichTextBox1.FileName
Form1.Caption = "AWord:" + Form1.RichTextBox1.FileName
Form1.Show
Unload Me
Else
MsgBox "不能不选择文件或文件名为空", vbInformation, "提示"
End If
End Sub
Private Sub Form_Load()
Text1.text = Dir1.Path & "\"
With Combo1
    .AddItem "*.doc"
    .AddItem "*.rtf"
    .AddItem "*.fns"
    .AddItem "*.*"
End With
    File1.Pattern = Combo1.text
    Dir1.Path = Drive1.Drive
    Text3.text = Dir2.Path & "\"
With Combo2
    .AddItem ".doc"
    .AddItem ".rtf"
    .AddItem ".fns"
    .AddItem " "
End With
    File2.Pattern = Combo2.text
    Dir2.Path = Drive2.Drive
        If Form1.WindowState = 2 Then
        Me.WindowState = 2
    Else
        If Form1.WindowState = 0 Then
            Me.WindowState = 0
            Me.Left = Form1.Left
            Me.Width = Form1.Width
            Me.Top = Form1.Top
            Me.Height = Form1.Height
        End If
    End If
End Sub
Private Sub Combo2_Click()
File2.Pattern = Combo2.text
End Sub
Private Sub Command3_Click()
Northern = File2.Path + "\" + Text4.text + Combo2.text
Form1.RichTextBox1.SaveFile Northern
Form1.RichTextBox1.FileName = Northern
Form1.Caption = "AWord:" + Form1.RichTextBox1.FileName
Form1.Show
Form2.BackColor = vbBlack
Unload Form2
End Sub
Private Sub Command4_Click()
Form1.Show
Unload Me
End Sub
Private Sub Dir2_Click()
File2.Path = Dir2.Path
Text3.text = Dir2.Path & "\"
End Sub
Private Sub Drive2_change()
Dir2.Path = Drive2.Drive
End Sub
Private Sub File2_Click()
Text4.text = File2.FileName
End Sub
Private Sub File2_DblClick()
Northern = File2.Path + "\" + Text4.text + Combo2.text
Form1.RichTextBox1.SaveFile Northern
Form1.RichTextBox1.FileName = Northern
Form1.Show
Form1.Caption = "AWord:" + Form1.RichTextBox1.FileName
Unload Me
End Sub
Private Sub Me_Unload(Cancel As Integer)
Form1.Show
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
If Text5.text = "AuroraStudio" Then
MsgBox "Copyright AuroraStudio 2019-2023"
End If
If Text5.text = "INTRON" Then
MsgBox "Copyright INTRON 软件组织 2021-2022→Copyright AuroraStudio 2019-2023"
Text5.text = "AuroraStudio"
End If
If Text5.text = "Locker" Then
MsgBox "Locker 是 Winfans最烂的产品，没有之一"
End If
If Text5.text = "NetExplore" Then
MsgBox "沉舟侧畔千帆过，病树前头万木春......无需为曾经的 Internet Explorer 和 NetExplore 感到悲伤，因为它们归根结底都是历史的产物，自然也要回到历史当中去。——苏方华-RTC"
End If
End If
End Sub
