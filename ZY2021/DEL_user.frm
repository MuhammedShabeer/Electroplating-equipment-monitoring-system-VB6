VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form DEL_user 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   Caption         =   "User registration"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "DEL_user.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1440
      Top             =   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "User loading"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Canel"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "User Delete"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Register user"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1695
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2990
         _Version        =   393216
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         TextStyle       =   4
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   2280
      Width           =   2535
   End
End
Attribute VB_Name = "DEL_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Dim A_USER As Integer
 Dim A_USER_V As Variant
 A_USER_V = MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.RowSel, 0)
 A_USER = MsgBox("你要删除的用户:" & A_USER_V, vbYesNo, "提示")
 If A_USER = 6 Then
 If A_USER_V = userID Then MsgBox "你不能删除您自己", 48, "提示": Exit Sub
 Dim sq1 As String
 Dim rs_del As New ADODB.Recordset
 sq1 = "select * from 系统管理"
 rs_del.Open sq1, conn, adOpenKeyset, adLockPessimistic
 rs_del.MoveFirst
 While (Not rs_del.EOF)
  If rs_del.Fields(0) = MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.RowSel, 0) Then
    rs_del.Delete
    rs_del.Update
    rs_del.Close
    Call RS_DEL_UP
    MsgBox "用户:" & A_USER_V & "已成功删除", 48, "提示"
    Exit Sub
  Else
    rs_del.MoveNext
  End If
 Wend
 End If
End Sub

Private Sub Command3_Click()
 Unload Me
End Sub

Private Sub Form_Load()
   rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
Dim i%
Dim a As Integer
   MSFlexGrid1.ColWidth(i) = 1300
   MSFlexGrid1.Rows = 1
   MSFlexGrid1.Cols = 3
 For i = 0 To 2
   MSFlexGrid1.ColAlignment(i) = 4
   MSFlexGrid1.ColWidth(0) = 1250
   MSFlexGrid1.ColWidth(1) = 1250
   MSFlexGrid1.ColWidth(2) = 1250
   MSFlexGrid1.Col = i
   MSFlexGrid1.TextMatrix(0, 0) = "用户名"
   MSFlexGrid1.TextMatrix(0, 1) = "密码"
   MSFlexGrid1.TextMatrix(0, 2) = "权限"
 Next i
End Sub

Private Sub Command4_Click()
Call RS_DEL_UP
End Sub
  
Private Sub RS_DEL_UP()
Dim sq1 As String
Dim rs_del As New ADODB.Recordset
Dim a As Integer
 sq1 = "select * from 系统管理"
 rs_del.Open sq1, conn, adOpenKeyset, adLockPessimistic
 rs_del.MoveLast
 MSFlexGrid1.Rows = rs_del.RecordCount + 1
 MSFlexGrid1.Cols = rs_del.Fields.Count
 Label1.Caption = "当前有" & rs_del.RecordCount & "个用户"
 a = MSFlexGrid1.Rows - 1
 rs_del.MoveFirst
 If (Not rs_del.EOF) Then
  For i = 1 To a
  MSFlexGrid1.TextMatrix(i, 0) = rs_del.Fields(0)
  MSFlexGrid1.TextMatrix(i, 1) = rs_del.Fields(1)
  MSFlexGrid1.TextMatrix(i, 2) = rs_del.Fields(2)
  rs_del.MoveNext
  Next
  End If
  rs_del.Close
End Sub
