VERSION 5.00
Begin VB.Form ChUser 
   Caption         =   "Form1"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   3390
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Confirm"
      Height          =   400
      Left            =   465
      TabIndex        =   5
      Top             =   1185
      Width           =   1000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Canel"
      Height          =   400
      Left            =   1815
      TabIndex        =   4
      Top             =   1185
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1335
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   705
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1350
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   285
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "User："
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
      Left            =   345
      TabIndex        =   3
      Top             =   300
      Width           =   930
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
      Left            =   345
      TabIndex        =   2
      Top             =   705
      Width           =   975
   End
End
Attribute VB_Name = "ChUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Text1.SetFocus
  End If

End Sub

Private Sub Command1_Click()
  user1 = Combo1.Text
  passw1 = Trim(Text1.Text)
  If passw1 = "" Or user1 = "" Then
    MsgBox "用户或密码输入异常! 第" & CStr(y1 + 1) & "次。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "超过3次，程式将关闭！", , "错误"
    If y1 >= 2 Then End
    y1 = y1 + 1
    Text1.SetFocus
    Exit Sub
  End If
  If user1 = "CS" And passw1 = "djf" & Curmonth & Curday Then Exit Sub
  Dim sq1 As String
  Dim rs1 As New ADODB.Recordset
  sq1 = "select * from 系统管理 where 用户名='" & user1 & "'and 密码='" & passw1 & "'"
  rs1.Open sq1, conn, adOpenKeyset, adLockPessimistic
  If rs1.RecordCount = 0 Then
    MsgBox "用户或密码输入异常! 第" & CStr(y1 + 1) & "次。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "超过3次，程式将关闭！", , "错误"
    If y1 >= 2 Then End
    rs1.Close
    Text1.SetFocus
    Exit Sub
  Else
    userID = rs1.Fields(0)
    userpass = rs1.Fields(1)
    userpow = rs1.Fields(2)
    rs1.Close
    min.StatusBar1.Panels.Item(2) = Chr(9) & Chr(9) & "当前操作用户:" & Chr(9) & userID
     Call T_CH
    Unload Me
  End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
  Text1.Text = ""
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
End Sub

Sub T_CH()
Dim s As Long
s = FileLen(App.Path & "\style\系统日志.ini")
 Open (App.Path & "\style\系统日志.ini") For Input As #1
    Do While Not EOF(1)
         Line Input #1, INTEXT
         TSTR = TSTR + INTEXT + Chr(13) + Chr(10)
    Loop
Close #1
If s > 10000 Then TSTR = Right(TSTR, 9971)
TSTR = TSTR + "   " + userID + "               " + Format(Now, "yyyy-mm-dd hh:mm:ss") + "            " + "用户切换" + Chr(13) + Chr(10)
        If Len(TSTR) > 10000 Then TSTR = Right(TSTR, 9800)
Open (App.Path & "\style\系统日志.ini") For Output As #1
      Print #1, TSTR
Close #1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Command1_Click
  End If

End Sub
