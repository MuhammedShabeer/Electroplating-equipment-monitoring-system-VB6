VERSION 5.00
Begin VB.Form T_lock 
   BackColor       =   &H00C0FFC0&
   Caption         =   "������¼"
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
   StartUpPosition =   3  '����ȱʡ
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
Me.Caption = "ϵͳ��־"
 Text1.Text = ""
Dim INTEXT As String
Open (App.Path & "\style\ϵͳ��־.ini") For Input As #1
          Do While Not EOF(1)
               Line Input #1, INTEXT
               Text1.Text = Text1.Text + INTEXT + Chr(13) + Chr(10)
          Loop
      Close #1
End Sub
Private Sub Command2_Click()
Me.Caption = "Ͷ�ϼ�¼"
Text1.Text = ""
 Dim INTEXT As String
Open (App.Path & "\style\Ͷ��.ini") For Input As #3
          Do While Not EOF(3)
               Line Input #3, INTEXT
               Text1.Text = Text1.Text + INTEXT + Chr(13) + Chr(10)
          Loop
      Close #3
End Sub

Private Sub Command3_Click()
Me.Caption = "�Ϻű༭��¼"
Text1.Text = ""
 Dim INTEXT As String
Open (App.Path & "\style\�Ϻű༭.ini") For Input As #2
          Do While Not EOF(2)
               Line Input #2, INTEXT
               Text1.Text = Text1.Text + INTEXT + Chr(13) + Chr(10)
          Loop
      Close #2
End Sub

Private Sub Command4_Click()
Me.Caption = "�Ϻ�ʹ�ò�����¼"
 Text1.Text = ""
 Dim INTEXT As String
Open (App.Path & "\style\ʹ���Ϻ�.ini") For Input As #4
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
If userpow <> "ϵͳ����" Then MsgBox "��û��Ȩ��ɾ����¼", 48, "��ʾ": Exit Sub
a_input = InputBox("����������", "����", "")
If a_input <> userpass Then MsgBox "�������", 16, "��ʾ": Exit Sub
      Text1.Text = ""
      TSTR = Text1.Text
If Me.Caption = "ϵͳ��־" Then
      Open (App.Path & "\STYLE\ϵͳ��־.ini") For Output As #1
      Print #1, TSTR
      Close #1
ElseIf Me.Caption = "�Ϻű༭��¼" Then
     Open (App.Path & "\style\�Ϻű༭.ini") For Output As #2
      Print #2, TSTR
      Close #2
ElseIf Me.Caption = "Ͷ�ϼ�¼" Then
     Open (App.Path & "\STYLE\Ͷ��.ini") For Output As #3
      Print #3, TSTR
      Close #3
ElseIf Me.Caption = "�Ϻ�ʹ�ò�����¼" Then
     Open (App.Path & "\style\ʹ���Ϻ�.ini") For Output As #4
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
