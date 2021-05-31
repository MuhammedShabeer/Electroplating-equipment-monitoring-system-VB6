VERSION 5.00
Begin VB.Form LU_lock 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Process imformation"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   Icon            =   "LU_lock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6480
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "change "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comfir"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   4815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "LU_lock.frx":044A
      Top             =   360
      Width           =   6135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "process imformationt"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      TabIndex        =   1
      Top             =   45
      Width           =   2700
   End
End
Attribute VB_Name = "LU_lock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim INTEXT As String
  Open (App.Path & "\style\lc.log") For Output As #1
  INTEXT = Text1.Text
  Print #1, INTEXT
  Close #1
End Sub

Private Sub Command2_Click()
Dim aiput As Integer
Dim aiput_pass As String
 aiput = MsgBox("你要改变当前的流程信息吗", 32, "提示")
 If aiput = 1 Then
  aiput_pass = InputBox("请输入密吗", "密码", "")
   If Trim(aiput_pass) = "27721165" Then
     Text1.Locked = False
   Else
     MsgBox "密码错误！", 16, "密码"
   End If
 End If
End Sub

Private Sub Command3_Click()
 Unload Me
End Sub

Private Sub Form_Load()
rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
Dim INTEXT As String
Open (App.Path & "\style\lc.log") For Input As #1
  Do While Not EOF(1)
   Line Input #1, INTEXT
   Me.Text1.Text = Me.Text1.Text + INTEXT + Chr(13) + Chr(10)
   Loop
      Close #1
End Sub
