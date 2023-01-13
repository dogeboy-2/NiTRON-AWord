VERSION 5.00
Begin VB.Form SaveFileWindow 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "±£´æÎÄ¼þ"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8550
   Icon            =   "OpenFileWindow - ¸±±¾.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ËùÓÐÕßÖÐÐÄ
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   5535
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      TabIndex        =   5
      Top             =   120
      Width           =   5535
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Text            =   ".doc"
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "È¡Ïû"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "È·¶¨"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   5160
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   3000
      TabIndex        =   1
      Top             =   720
      Width           =   5295
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3510
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   8640
      X2              =   0
      Y1              =   5040
      Y2              =   5040
   End
End
Attribute VB_Name = "SaveFileWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Northern As String
Private Sub Combo1_Click()
File1.Pattern = Combo1.text
End Sub
Private Sub Command1_Click()
Northern = SaveFileWindow.File1.Path + "\" + SaveFileWindow.Text2.text + SaveFileWindow.Combo1.text
Form1.RichTextBox1.SaveFile Northern
Form1.RichTextBox1.FileName = Northern
Form1.Caption = "AWord:" + Form1.RichTextBox1.FileName
Form1.Show
Unload Me
Unload SaveFileWindow
End Sub
Private Sub Command2_Click()
Unload Me
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
Northern = OpenFileWindow.File1.Path + "\" + SaveFileWindow.Text2.text + SaveFileWindow.Combo1.text
Form1.RichTextBox1.SaveFile Northern
Form1.RichTextBox1.FileName = Northern
Form1.Show
Form1.Caption = "AWord:" + Form1.RichTextBox1.FileName
Unload Me
End Sub
Private Sub Form_Load()
Text1.text = Dir1.Path & "\"
With Combo1
    .AddItem ".doc"
    .AddItem ".rtf"
    .AddItem ".fns"
    .AddItem " "
End With
    File1.Pattern = Combo1.text
    Dir1.Path = Drive1.Drive
End Sub

