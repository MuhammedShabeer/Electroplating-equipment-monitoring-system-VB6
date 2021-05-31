VERSION 5.00
Begin VB.Form Add_user 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add user"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Add_user.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Add_user.frx":0442
      Left            =   2040
      List            =   "Add_user.frx":044C
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   2040
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2040
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   240
      Top             =   600
   End
   Begin VB.CommandButton Command2 
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
      Left            =   2760
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Confirm"
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
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   2055
      Left            =   240
      TabIndex        =   3
      Top             =   300
      Width           =   4215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Confirm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Authority"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   0
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enter poswod"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "New name"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "在此键入新用户及权限，只有管理员才可以操作！"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3960
   End
End
Attribute VB_Name = "Add_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
 For i = 0 To 2
   Text1(i).Text = ""
 Next i
 Combo1.Text = ""
 Text1(1).PasswordChar = "*"
 Text1(2).PasswordChar = "*"
 Timer1.Enabled = True
  End Sub

Private Sub Timer1_Timer()
  Text1(0).SetFocus
  Timer1.Enabled = False
End Sub
Private Sub Command1_Click()
 Dim sq1 As String
 Dim rs_add As New ADODB.Recordset
 On Error GoTo error_add
 If Trim(Text1(0).Text) = "" Then
    MsgBox "请必须输入用户", 48, "提示"
    Exit Sub
    Text1(0).SetFocus
 Else
    sq1 = "select * from 系统管理"
    rs_add.Open sq1, conn, adOpenKeyset, adLockPessimistic
    While (rs_add.EOF = False)
       If Trim(rs_add.Fields(0)) = Trim(Text1(0).Text) Then
          MsgBox "已有此用户！", 48, "提示"
          Text1(0).SetFocus
          For i = 0 To 2
           Text1(i).Text = ""
          Next i
          Combo1.Text = ""
          Exit Sub
        Else
         rs_add.MoveNext
        End If
    Wend

  If Trim(Text1(1)) <> Trim(Text1(2)) Then
     MsgBox "两次密码不致！", 48, "提示"
     Text1(1).SetFocus
     For i = 1 To 2
       Text1(i).Text = ""
     Next i
     Combo1.Text = ""
     Exit Sub
  ElseIf Trim(Combo1.Text) <> "系统管理" And Trim(Combo1.Text) <> "受限用户" Then
     MsgBox "请选择权限!", 48, "提示"
     Combo1.SetFocus
     Combo1.Text = ""
     Exit Sub
  ElseIf userpow <> "系统管理" Then
     MsgBox "你没有权限创建新的用户,请进入管理员模式!", 48, "提示"
     Exit Sub
  Else
     rs_add.AddNew
       rs_add.Fields(0) = Text1(0).Text
       rs_add.Fields(1) = Text1(1).Text
       rs_add.Fields(2) = Combo1.Text
       rs_add.Update
       rs_add.Close
    
     MsgBox "添加新用户成功！", 48, "提示"
  End If
 End If
error_add: MsgBox Err.Description
End Sub
Private Sub Command2_Click()
 Unload Me
End Sub
