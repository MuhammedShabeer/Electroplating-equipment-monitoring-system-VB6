VERSION 5.00
Begin VB.Form T_lock 
   BackColor       =   &H00C0FFC0&
   Caption         =   "操作记录"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   Icon            =   "T_lock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   8160
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "oprator item record"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Item edit record"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Feeding record"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "system login record"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   7575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   7935
   End
End
Attribute VB_Name = "T_lock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Call com_01

End Sub
Sub com_01()
Me.Caption = "系统日志"
 Text1.Text = ""
Dim INTEXT As String
Open (App.Path & "\style\系统日志.ini") For Input As #1
          Do While Not EOF(1)
               Line Input #1, INTEXT
               Text1.Text = Text1.Text + INTEXT + Chr(13) + Chr(10)
          Loop
      Close #1
End Sub
Private Sub Command2_Click()
Me.Caption = "投料记录"
Text1.Text = ""
 Dim INTEXT As String
Open (App.Path & "\style\投料.ini") For Input As #3
          Do While Not EOF(3)
               Line Input #3, INTEXT
               Text1.Text = Text1.Text + INTEXT + Chr(13) + Chr(10)
          Loop
      Close #3
End Sub

Private Sub Command3_Click()
Me.Caption = "料号编辑记录"
Text1.Text = ""
 Dim INTEXT As String
Open (App.Path & "\style\料号编辑.ini") For Input As #2
          Do While Not EOF(2)
               Line Input #2, INTEXT
               Text1.Text = Text1.Text + INTEXT + Chr(13) + Chr(10)
          Loop
      Close #2
End Sub

Private Sub Command4_Click()
Me.Caption = "料号使用操作记录"
 Text1.Text = ""
 Dim INTEXT As String
Open (App.Path & "\style\使用料号.ini") For Input As #4
          Do While Not EOF(4)
               Line Input #4, INTEXT
               Text1.Text = Text1.Text + INTEXT + Chr(13) + Chr(10)
          Loop
      Close #4
End Sub

Private Sub Command5_Click()
 Unload Me
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim INTEXT As String
Dim TSTR As String
Dim a_input As Integer
If userpow <> "系统管理" Then MsgBox "你没有权限删除记录", 48, "提示": Exit Sub
a_input = InputBox("请输入密码", "密码", "")
If a_input <> userpass Then MsgBox "密码错误", 16, "提示": Exit Sub
      Text1.Text = ""
      TSTR = Text1.Text
If Me.Caption = "系统日志" Then
      Open (App.Path & "\STYLE\系统日志.ini") For Output As #1
      Print #1, TSTR
      Close #1
ElseIf Me.Caption = "料号编辑记录" Then
     Open (App.Path & "\style\料号编辑.ini") For Output As #2
      Print #2, TSTR
      Close #2
ElseIf Me.Caption = "投料记录" Then
     Open (App.Path & "\STYLE\投料.ini") For Output As #3
      Print #3, TSTR
      Close #3
ElseIf Me.Caption = "料号使用操作记录" Then
     Open (App.Path & "\style\使用料号.ini") For Output As #4
      Print #4, TSTR
      Close #4
End If
End Sub

Private Sub Form_Load()
rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
  Me.Left = (Screen.Width - Me.Width) / 2
   Me.Top = (Screen.Height - Me.Height) / 2
   Call com_01
End Sub
