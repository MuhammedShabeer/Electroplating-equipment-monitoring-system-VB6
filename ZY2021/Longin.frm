VERSION 5.00
Begin VB.Form Longin 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "中新电镀设备制造有限公司"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000A&
   Icon            =   "Longin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer4 
      Interval        =   200
      Left            =   2280
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1080
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "中新电镀设备公司"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   2.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   45
      Left            =   2160
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   50
      Picture         =   "Longin.frx":0442
      Stretch         =   -1  'True
      Top             =   80
      Width           =   4575
   End
End
Attribute VB_Name = "Longin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Integer

Private Sub Form_Load()
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
  Label1.Visible = True
  Label1.Font.Size = 2
  b = 0
End Sub
Private Sub Timer1_Timer()
  Image1.Picture = LoadPicture(App.Path & "\zx\SSL11763.jpg")
  Timer2.Enabled = True
  Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
  Timer3.Enabled = True
  Image1.Picture = LoadPicture(App.Path & "\zx\SSL11758.jpg")
End Sub
Private Sub Timer3_Timer()
  welcome.Show
  Unload Me
End Sub
Private Sub Timer4_Timer()
 b = b + 1
 a = a + 3
 If b < 10 Then Label1.Font.Size = a
End Sub

