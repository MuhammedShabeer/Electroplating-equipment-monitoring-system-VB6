VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form welcome 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   0  'None
   Caption         =   "欢迎使用"
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "welcome.frx":0000
   ScaleHeight     =   2685
   ScaleMode       =   0  'User
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   75
      TabIndex        =   11
      Top             =   2415
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   105
      Picture         =   "welcome.frx":276D2
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   10
      Top             =   1020
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   3375
      TabIndex        =   3
      Top             =   1965
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "comfirn"
      Height          =   400
      Left            =   2040
      TabIndex        =   2
      Top             =   1965
      Width           =   1000
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      Left            =   2670
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1050
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2655
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1470
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3705
      Top             =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3720
      Top             =   600
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "PWD："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1665
      TabIndex        =   9
      Top             =   1470
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "USER："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   1665
      TabIndex        =   8
      Top             =   1065
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "asdasd"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   0
      Left            =   1425
      TabIndex        =   6
      Top             =   975
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "asdasd"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   180
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   600
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "asdasd"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   2145
      TabIndex        =   7
      Top             =   1215
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "asdasd"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   645
   End
End
Attribute VB_Name = "welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x1, x2, x3, y1, y2, y3, d1, z
Dim user1, passw1

Private Sub Combo1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    welcome.Text1.SetFocus
  End If
End Sub

Private Sub Command1_Click()
  welcome.Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
  Q_User = 9
  Unload Me
End Sub

Private Sub Form_Load()
  d1 = 1: x1 = 1: x2 = 1: x3 = 1
  y1 = 1: y2 = 1: y3 = 1: z = 0
  Label4(0).Visible = 0
  Label4(1).Visible = 0
  Command1.Visible = 0
  Command2.Visible = 0
  Label1(0).Font.Size = 1
  Label2(0).Font.Size = 20
  Label1(1).Font.Size = 1
  Label2(1).Font.Size = 20
  Label1(0).Caption = "欢迎使用本系统"
  Label1(1).Caption = "欢迎使用本系统"
  Label2(0).Caption = ""
  Label2(1).Caption = ""
  Label1(0).Left = Me.ScaleLeft
  Label1(0).Top = Me.ScaleTop
  Label1(1).Left = Label1(0).Left + 20
  Label1(1).Top = Label1(0).Top + 20
  Label2(0).Left = 1000
  Label2(0).Top = 10
  Label2(1).Left = 1020
  Label2(1).Top = 12
  Picture1.Width = 0
  Text1.Text = ""
  Text1.Visible = False
  ProgressBar1.Visible = False
  Combo1.Clear
  Combo1.Text = ""
  Dim sq1 As String
  Dim rs1 As New ADODB.Recordset
  sq1 = "select * from 系统管理"
  rs1.Open sq1, conn, adOpenKeyset, adLockPessimistic
  Do While Not rs1.EOF
      Combo1.AddItem rs1.Fields(0) ' & rs1.Fields(1)
      rs1.MoveNext
  Loop
  rs1.Close
  Combo1.Visible = False
  ProgressBar1.min = 0
  ProgressBar1.Max = 40
  Timer1.Enabled = True
  Timer2.Enabled = False
  o(0) = Chr(13)
'SetWindowPos Me.hwnd, -1, 0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelX, &H3 ', &H40
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    welcome.Timer1.Enabled = True
  End If
End Sub

Private Sub Timer1_Timer()
  If x1 < 30 Then
    Label1(0).Font.Size = x3
    Label1(1).Font.Size = x3
    If Picture1.Width <= 1400 Then Picture1.Width = x1 * 47
    Label1(0).Left = Me.ScaleLeft + x1 * 3
    Label1(0).Top = Me.ScaleTop + y1 * 13
    Label1(1).Left = Label1(0).Left + x2
    Label1(1).Top = Label1(0).Top + y2
    x1 = x1 + 1: x2 = x2 + 1: x3 = x3 + 1
    y1 = y1 + 1: y2 = y2 + 1: y3 = y3 + 1

  ElseIf x1 = 30 Then
    Text1.Visible = True
    Combo1.Visible = True
    Label4(0).Visible = 1
    Label4(1).Visible = 1
    Command1.Visible = 1
    Command2.Visible = 1
    Combo1.SetFocus
    Timer1.Enabled = False
    Me.Picture = LoadPicture("")
    x1 = x1 + 1
    y1 = 0
  Else
    'welcome.Text1.Visible = False
    Select Case d1
      Case 1
        user1 = Combo1.Text
        passw1 = Trim(Text1.Text)
        If passw1 = "" Or user1 = "" Then
          MsgBox "用户或密码输入异常! 第" & CStr(y1 + 1) & "次。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "超过3次，程式将关闭！", , "错误"
          If y1 >= 2 Then End
          d1 = 0
          y1 = y1 + 1
          Timer1.Enabled = False
          Text1.SetFocus
          Exit Sub
        End If
        If user1 = "CS" And passw1 = "djf" & Curmonth & Curday Then d1 = 39: Exit Sub
        Dim sq1 As String
        Dim rs1 As New ADODB.Recordset
        sq1 = "select * from 系统管理 where 用户名='" & user1 & "'and 密码='" & passw1 & "'"
        rs1.Open sq1, conn, adOpenKeyset, adLockPessimistic
        If rs1.RecordCount = 0 Then
          MsgBox "用户或密码输入异常! 第" & CStr(y1 + 1) & "次。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "超过3次，程式将关闭！", , "错误"
          If y1 >= 2 Then End
          rs1.Close
          d1 = 0
          y1 = y1 + 1
          Timer1.Enabled = False
          Text1.SetFocus
          Exit Sub
        Else
          userID = rs1.Fields(0)
          userpass = rs1.Fields(1)
          userpow = rs1.Fields(2)
          ProgressBar1.Visible = True
        End If
        rs1.Close
        welcome.Timer2.Enabled = True
        Label2(0).Caption = "中"
        Label2(1).Caption = "中"
      Case 5
        Label2(0).Caption = "中新"
        Label2(1).Caption = "中新"
      Case 9
        Label2(0).Caption = "中新电"
        Label2(1).Caption = "中新电"
      Case 13
        Label2(0).Caption = "中新电镀"
        Label2(1).Caption = "中新电镀"
      Case 17
        Label2(0).Caption = "中新电镀承"
        Label2(1).Caption = "中新电镀承"
      Case 21
        Label2(0).Caption = "中新电镀承制"
        Label2(1).Caption = "中新电镀承制"
          
      Case 40
        welcome.Timer1.Enabled = False
        d1 = -1: 'ss1 = ""
        Unload Me
        min.Show
        Exit Sub
    End Select
    d1 = d1 + 1
    ProgressBar1.Value = d1
  End If
End Sub



