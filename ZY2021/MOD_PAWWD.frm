VERSION 5.00
Begin VB.Form MOD_PAWWD 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�޸�����"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "MOD_PAWWD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comfirm"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   2040
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   2175
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "user"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "old PWD"
         BeginProperty Font 
            Name            =   "����"
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
         TabIndex        =   9
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "New PWD"
         BeginProperty Font 
            Name            =   "����"
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
         TabIndex        =   8
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Comfirm"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   1050
      End
   End
End
Attribute VB_Name = "MOD_PAWWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Form_Load()
'rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
Dim i As Integer
 For i = 0 To 3
   Text1(i) = ""
 Next
 For i = 1 To 3
   Text1(i).PasswordChar = "*"
 Next
End Sub

Private Sub Command1_Click()
 Dim sq1 As String
 Dim rs_mod As New ADODB.Recordset
 If Trim(Text1(0)) = "" Then MsgBox "��������Ҫ����������û�", 48, "��ʾ": Text1(0).SetFocus: Exit Sub
 If Trim(Text1(1)) = "" Then MsgBox "������ԭ������", 48, "��ʾ": Text1(1).SetFocus: Exit Sub: Exit Sub
 If Trim(Text1(2)) = "" Then MsgBox "������������", 48, "��ʾ": Text1(2).SetFocus: Exit Sub
 If Trim(Text1(3)) = "" Then MsgBox "���ٴ�����������", 48, "��ʾ": Text1(3).SetFocus: Exit Sub
 sq1 = "Select * from ϵͳ���� where �û��� = '" & Text1(0).Text & "'"
    rs_mod.Open sq1, conn, adOpenForwardOnly, adLockReadOnly
    If rs_mod.EOF = True Then
      MsgBox "�Բ���,û����Ҫ���ĵ��û�", 48, "��ʾ"
      For i = 0 To 3
       Text1(i) = ""
      Next
      Exit Sub
    Else
      If Trim(rs_mod.Fields(1)) <> Trim(Text1(1).Text) Then
         MsgBox "ԭ�����벻��ȷ", 48, "��ʾ"
         For i = 2 To 3
           Text1(i) = ""
         Next
         Exit Sub
      ElseIf Trim(Text1(2).Text) <> Trim(Text1(3).Text) Then
         MsgBox "�������벻һ�£�", 48, "��ʾ"
           Text1(2).Text = ""
           Text1(3).Text = ""
           Exit Sub
      ElseIf userpow <> "ϵͳ����" Then
          MsgBox "��û��Ȩ���޸�����,��������Աģʽ!", 48, "��ʾ"
          Exit Sub
      Else
       sq1 = "Update ϵͳ���� set ����='" & Text1(2).Text & "' where �û��� = '" & Text1(0).Text & "'"
        conn.Execute sq1
         'rs_mod.Fields(0) = Text1(0).Text
         'rs_mod.Fields(1) = Text1(2).Text
         'rs_mod.Update
         rs_mod.Close
         MsgBox "�޸ĳɹ�!", 48, "��ʾ"
         For i = 0 To 3
           Text1(i) = ""
         Next
      End If
    End If
End Sub


