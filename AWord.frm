VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   6615
   ClientLeft      =   225
   ClientTop       =   765
   ClientWidth     =   12510
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "AWord.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   12510
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Left            =   4080
      Top             =   1920
   End
   Begin RichTextLib.RichTextBox RTB2 
      Height          =   135
      Left            =   1800
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   238
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"AWord.frx":19D02
   End
   Begin VB.CommandButton Command17 
      Caption         =   "打印"
      Height          =   495
      Left            =   9480
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5680
      TabIndex        =   3
      Text            =   "14"
      Top             =   240
      Width           =   1095
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
      Left            =   3880
      TabIndex        =   2
      Text            =   "微软雅黑"
      Top             =   240
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1695
      Left            =   3120
      TabIndex        =   1
      Top             =   3360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2990
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"AWord.frx":19D84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2470
      Picture         =   "AWord.frx":19E15
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   0
      ToolTipText     =   "粘贴"
      Top             =   360
      Width           =   615
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000000&
      X1              =   120
      X2              =   2400
      Y1              =   2040
      Y2              =   1920
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "段落"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7960
      TabIndex        =   7
      Top             =   1350
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "字体"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5200
      TabIndex        =   6
      Top             =   1350
      Width           =   615
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000010&
      X1              =   9260
      X2              =   9260
      Y1              =   120
      Y2              =   1680
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   7000
      X2              =   7000
      Y1              =   120
      Y2              =   1680
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "剪贴板"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2750
      TabIndex        =   5
      Top             =   1350
      Width           =   615
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   3760
      X2              =   3760
      Y1              =   120
      Y2              =   1680
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "撤销"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1510
      TabIndex        =   4
      Top             =   1350
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   2200
      X2              =   2200
      Y1              =   120
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   1200
      X2              =   1200
      Y1              =   120
      Y2              =   1680
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   6490
      Picture         =   "AWord.frx":1A52E
      Top             =   741
      Width           =   330
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   555
      Left            =   5900
      Picture         =   "AWord.frx":1A8EE
      Top             =   720
      Width           =   465
   End
   Begin VB.Image Command6 
      Height          =   600
      Left            =   320
      Picture         =   "AWord.frx":1ACBE
      ToolTipText     =   "保存"
      Top             =   480
      Width           =   600
   End
   Begin VB.Image Command14 
      Height          =   360
      Left            =   1495
      Picture         =   "AWord.frx":1B4B1
      ToolTipText     =   "重做"
      Top             =   760
      Width           =   360
   End
   Begin VB.Image Command13 
      Height          =   360
      Left            =   1495
      Picture         =   "AWord.frx":1B9B2
      ToolTipText     =   "撤消"
      Top             =   195
      Width           =   360
   End
   Begin VB.Image Command16 
      Height          =   510
      Left            =   8580
      Picture         =   "AWord.frx":1BEA1
      ToolTipText     =   "固定内容"
      Top             =   765
      Width           =   405
   End
   Begin VB.Image Command15 
      Height          =   465
      Left            =   7840
      Picture         =   "AWord.frx":1C327
      ToolTipText     =   "项目符号"
      Top             =   760
      Width           =   510
   End
   Begin VB.Image Command12 
      Height          =   555
      Left            =   7140
      Picture         =   "AWord.frx":1C7E5
      ToolTipText     =   "查找内容"
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Command11 
      Height          =   465
      Left            =   8500
      Picture         =   "AWord.frx":1CC24
      ToolTipText     =   "右对齐"
      Top             =   180
      Width           =   465
   End
   Begin VB.Image Command10 
      Height          =   495
      Left            =   7840
      Picture         =   "AWord.frx":1D018
      ToolTipText     =   "居中"
      Top             =   150
      Width           =   495
   End
   Begin VB.Image Command9 
      Height          =   465
      Left            =   7220
      Picture         =   "AWord.frx":1D438
      ToolTipText     =   "左对齐"
      Top             =   165
      Width           =   420
   End
   Begin VB.Image Command8 
      Height          =   315
      Left            =   5430
      Picture         =   "AWord.frx":1D827
      ToolTipText     =   "删除线"
      Top             =   840
      Width           =   420
   End
   Begin VB.Image Command7 
      Height          =   510
      Left            =   4920
      Picture         =   "AWord.frx":1DC35
      ToolTipText     =   "下划线"
      Top             =   760
      Width           =   375
   End
   Begin VB.Image Command5 
      Height          =   360
      Left            =   4410
      Picture         =   "AWord.frx":1DFA6
      ToolTipText     =   "斜体"
      Top             =   840
      Width           =   360
   End
   Begin VB.Image Command4 
      Height          =   360
      Left            =   3940
      Picture         =   "AWord.frx":1E3D5
      Top             =   840
      Width           =   360
   End
   Begin VB.Image Command2 
      Height          =   360
      Left            =   3200
      Picture         =   "AWord.frx":1E8DC
      Top             =   220
      Width           =   360
   End
   Begin VB.Image Command3 
      Height          =   360
      Left            =   3200
      Picture         =   "AWord.frx":1EDF5
      ToolTipText     =   "剪切"
      Top             =   780
      Width           =   360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑"
      Begin VB.Menu mnuCopy 
         Caption         =   "复制"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCut 
         Caption         =   "剪贴"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "粘贴"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep 
         Caption         =   "-"
      End
      Begin VB.Menu Bold 
         Caption         =   "加粗"
         Shortcut        =   ^B
      End
      Begin VB.Menu what 
         Caption         =   "倾斜"
         Shortcut        =   ^I
      End
      Begin VB.Menu UnderLine 
         Caption         =   "下划线"
         Shortcut        =   ^U
      End
      Begin VB.Menu Delete 
         Caption         =   "删除线"
      End
      Begin VB.Menu SD 
         Caption         =   "-"
      End
      Begin VB.Menu onleft 
         Caption         =   "左对齐"
         Shortcut        =   ^L
      End
      Begin VB.Menu oncenter 
         Caption         =   "居中"
         Shortcut        =   ^E
      End
      Begin VB.Menu onright 
         Caption         =   "右对齐"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSEP2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSelectAll 
         Caption         =   "全选"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "搜索"
      Begin VB.Menu mnuFind 
         Caption         =   "查找"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindOn 
         Caption         =   "查找下一个"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'声明查找变量
Dim sFind As String
'声明文件类型
Dim FileType, FiType As String
Dim FileName As String
Option Explicit
Dim bChg As Boolean '记录富文本框内容是否发生变化
Dim UndoNum As Long
Dim n As Long, chazhao As String
 Private TargetPosition As Integer
      Public Ask As Boolean
Private Sub Bold_Click()
On Error Resume Next
RichTextBox1.SelBold = Not RichTextBox1.SelBold
End Sub
Private Sub Command1_Click()
Call mnuPaste_Click
End Sub
Private Sub Command10_Click()
Call oncenter_Click
End Sub
Private Sub Command11_Click()
Call onright_Click
End Sub
Private Sub Command12_Click()
Call mnuFind_Click
End Sub
Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(text), wait
   Set WshShell = Nothing
End Sub
Private Sub Command13_Click()
On Error Resume Next
RichTextBox1.SetFocus
Sendkeys ("^Z")
End Sub
Private Sub Command14_Click()
On Error Resume Next
RichTextBox1.SetFocus
Sendkeys ("^Y")
End Sub
Private Sub Command15_Click()
On Error Resume Next
RichTextBox1.SelBullet = Not RichTextBox1.SelBullet
End Sub
Private Sub Command16_Click()
On Error Resume Next
RichTextBox1.SelProtected = Not RichTextBox1.SelProtected
End Sub

Private Sub Command2_Click()
Call mnuCopy_Click
End Sub
Private Sub Command3_Click()
Call mnuCut_Click
End Sub
Private Sub Command4_Click()
Call Bold_Click
End Sub
Private Sub Command5_Click()
Call what_Click
End Sub
Private Sub Command6_Click()
If Not RichTextBox1.FileName = "" And Not RichTextBox1.FileName = "未命名" Then
RichTextBox1.SaveFile RichTextBox1.FileName
Me.Caption = "AWord:" & RichTextBox1.FileName
Else
SaveFileWindow.Show
End If
Ask = False
End Sub
Private Sub Command7_Click()
Call UnderLine_Click
End Sub
Private Sub Command8_Click()
Call Delete_Click
End Sub
Private Sub Command9_Click()
Call onleft_Click
End Sub
Private Sub Delete_Click()
On Error Resume Next
RichTextBox1.SelStrikeThru = Not RichTextBox1.SelStrikeThru
End Sub
'设置编辑框的位置和大小
Private Sub Form_Resize()
On Error Resume Next '出错处理
RichTextBox1.Top = 1900
RichTextBox1.Left = 100
RichTextBox1.Height = ScaleHeight
RichTextBox1.Width = ScaleWidth - 20
Line6.X1 = 10
Line6.X2 = ScaleWidth - 20
Line6.Y1 = RichTextBox1.Top - 20
Line6.Y2 = RichTextBox1.Top - 20
End Sub
Private Sub Image1_Click()
On Error Resume Next
RichTextBox1.SetFocus
Sendkeys "^(=)"
RichTextBox1.Refresh
End Sub
Private Sub Image2_Click()
On Error Resume Next
RichTextBox1.SetFocus
Sendkeys "^(+=)"
RichTextBox1.Refresh
End Sub
Private Sub mnuFile_Click()
Form2.Show
End Sub
'新建文件
Private Sub mnuNew_Click()
RichTextBox1.text = "" '清空文本框
FileName = "未命名"
Me.Caption = "AWord:" & FileName
End Sub
'打开文件
Private Sub mnuOpen_Click()
OpenFileWindow.Show
Me.Caption = "AWord:" & RichTextBox1.FileName
End Sub
'保存文件
Private Sub mnuSave_Click()
If Not RichTextBox1.FileName = "" And Not RichTextBox1.FileName = "未命名" Then
RichTextBox1.SaveFile RichTextBox1.FileName
Me.Caption = "AWord:" & RichTextBox1.FileName
Else
SaveFileWindow.Show
End If
End Sub
'退出
Private Sub mnuExit_Click()
End
End Sub
'复制
Private Sub mnuCopy_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText
End Sub
'剪切
Private Sub mnuCut_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText
RichTextBox1.SelText = ""
End Sub
Private Sub mnuSaveAs_Click()
SaveFileWindow.Show
End Sub
Private Sub mnuFind_Click()
 Dim s As String
s = RichTextBox1.text
chazhao = InputBox("")
If chazhao = "" Then
MsgBox "查找内容为空，请输入有效的查找内容"
Exit Sub
End If
n = InStr(s, chazhao)
If n > 0 Then
RichTextBox1.SelStart = n - 1
RichTextBox1.SelLength = Len(chazhao)
RichTextBox1.SetFocus
Else
MsgBox "此文件中不包含所查找的指定内容"
End If
End Sub
'全选
Private Sub mnuSelectAll_Click()
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = Len(RichTextBox1.text)
End Sub
'粘贴
Private Sub mnuPaste_Click()
On Error Resume Next
RichTextBox1.SelText = Clipboard.GetText
End Sub
'继续查找
Private Sub mnuFindOn_Click()
Dim s As String, linshi As Long
If chazhao = "" Then
chazhao = InputBox("")
If chazhao = "" Then
MsgBox "查找内容为空，请输入有效的查找内容"
Exit Sub
End If
End If
If n = 0 Then
s = RichTextBox1.text
Else
s = Mid(RichTextBox1.text, n + Len(chazhao), Len(RichTextBox1.text) - n - Len(chazhao))
End If
linshi = n
n = InStr(s, chazhao)
If n > 0 Then
n = n + linshi
RichTextBox1.SelStart = n - 1
RichTextBox1.SelLength = Len(chazhao)
RichTextBox1.SetFocus
Else
If linshi <> 0 Then
MsgBox "已搜索全部"
End If
End If
End Sub
Private Sub oncenter_Click()
If RichTextBox1.SelProtected = False Then
RichTextBox1.SelAlignment = rtfCenter
Else
MsgBox "要进行居中的文本处于锁定状态！"
End If
End Sub
Private Sub onleft_Click()
If RichTextBox1.SelProtected = False Then
RichTextBox1.SelAlignment = rtfLeft
Else
MsgBox "要进行左对齐的文本处于锁定状态！"
End If
End Sub
Private Sub onright_Click()
If RichTextBox1.SelProtected = False Then
RichTextBox1.SelAlignment = rtfRight
Else
MsgBox "要进行右对齐的文本处于锁定状态！"
End If
End Sub
'设置弹出式菜单（即在编辑框中单击鼠标右键时弹出的动态菜单）
Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuEdit, vbPopupMenuLeftAlign
Else
Exit Sub
End If
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer) '回车键，需要改按钮
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
    RichTextBox1.SelFontName = (Combo1.text)
    End If
End Sub




Private Sub Timer2_Timer()
If Not RichTextBox1.FileName = "" And Not RichTextBox1.FileName = "未命名" Then
Me.Caption = "AWord:" & RichTextBox1.FileName & "- 已保存"
Else
Exit Sub
End If
End Sub

Private Sub UnderLine_Click()
On Error Resume Next
RichTextBox1.SelUnderline = Not RichTextBox1.SelUnderline
End Sub
Private Sub what_Click()
On Error Resume Next
RichTextBox1.SelItalic = Not RichTextBox1.SelItalic
End Sub
Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyShift Then
RichTextBox1.SelFontName = (Combo1.text)
RichTextBox1.Refresh
RichTextBox1.SelFontName = (Combo2.text)
RichTextBox1.Refresh
End If
End Sub
Private Sub Form_Load()
RichTextBox1.SelFontName = (Combo1.text)
RichTextBox1.Refresh
RichTextBox1.SelFontSize = (Combo2.text)
RichTextBox1.Refresh
With Combo1
.AddItem "等线"
.AddItem "等线 Light"
.AddItem "楷体"
.AddItem "仿宋"
.AddItem "宋体"
.AddItem "黑体"
.AddItem "微软雅黑"
.AddItem "微软雅黑 Light"
.AddItem "新宋体"
.AddItem "Segoe UI"
.AddItem "Agency FB"
.AddItem "Bahnschrift"
.AddItem "Bauhaus 93"
.AddItem "Bell MT"
.AddItem "Berlin Sans FB"
.AddItem "Cambria"
End With
Dim webnet As String
webnet = VBA.Command
If Not webnet = "" Then
RichTextBox1.LoadFile webnet
End If
Ask = False
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer) '回车键，需要改按钮
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
    RichTextBox1.SelFontSize = (Combo2.text)
    End If
End Sub
Private Sub RichTextBox1_Change()
On Error Resume Next
RichTextBox1.SelFontName = (Combo1.text)
RichTextBox1.Refresh
RichTextBox1.SelFontSize = (Combo2.text)
RichTextBox1.Refresh
Combo1.text = RichTextBox1.SelFontName
Combo2.text = RichTextBox1.SelFontSize
Ask = True
RTB2.TextRTF = RichTextBox1.TextRTF
RTB2.FileName = RichTextBox1.FileName
If Not RichTextBox1.FileName = "" And Not RichTextBox1.FileName = "未命名" Then
Me.Caption = "AWord:" & RichTextBox1.FileName & "- 正在保存"
RTB2.SaveFile RichTextBox1.FileName
Timer2.Enabled = True
Timer2.Interval = 1000
Ask = False
Else
Exit Sub
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Ask = True Then
Dim Flag As Integer, MsgStr As String
MsgStr = "文件已经改变，是否要保存？" '提示语
Flag = MsgBox(MsgStr, vbYesNoCancel, "提示") '给予提示
If Flag = vbYes Then
Command6_Click
Exit Sub
End If
If Flag = vbNo Then
Exit Sub
End If
If Flag = vbCancel Then
Cancel = True
End If
End If
End Sub
