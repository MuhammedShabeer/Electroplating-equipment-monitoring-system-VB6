VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form LAO_eidt 
   BackColor       =   &H00C0FFC0&
   Caption         =   "�Ϻű༭"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7065
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6360
      Top             =   480
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Delet"
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
      Left            =   5235
      TabIndex        =   4
      ToolTipText     =   "ɾ��ѡȡ���Ϻ�"
      Top             =   4860
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
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
      Left            =   5235
      TabIndex        =   5
      ToolTipText     =   "�˳��Ϻű༭"
      Top             =   5340
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Export"
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
      Left            =   5235
      TabIndex        =   3
      ToolTipText     =   "�ɵ���ѡȡ�ϺŵĲ���"
      Top             =   4380
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2205
      Left            =   195
      TabIndex        =   29
      Top             =   4260
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3889
      _Version        =   393216
      ForeColor       =   -2147483642
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorBkg    =   12648384
      GridColor       =   0
      GridColorFixed  =   -2147483642
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2835
      TabIndex        =   14
      Text            =   "Text4"
      ToolTipText     =   "�ڴ������ѯ�ؼ���,�������ѯ���е��Ϻ�"
      Top             =   3375
      Width           =   2295
   End
   Begin VB.TextBox Text2 
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
      Left            =   2400
      TabIndex        =   6
      Text            =   "0"
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Inquire"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5235
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "�������ȫ����ѯ�����밴��Ӧ���ϺŲ�ѯ"
      Top             =   3300
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Canel"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   915
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3300
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comfire"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   75
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3300
      Width           =   855
   End
   Begin VB.TextBox text1 
      Alignment       =   2  'Center
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
      Index           =   6
      Left            =   4335
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "0000"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox text1 
      Alignment       =   2  'Center
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
      Index           =   5
      Left            =   4335
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "0000"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox text1 
      Alignment       =   2  'Center
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
      Index           =   1
      Left            =   4320
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "0000"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox text1 
      Alignment       =   2  'Center
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
      Index           =   0
      Left            =   2520
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "0000"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox text1 
      Alignment       =   2  'Center
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
      Index           =   4
      Left            =   2535
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "0000"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox text1 
      Alignment       =   2  'Center
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
      Index           =   3
      Left            =   2535
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0000"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox text1 
      Alignment       =   2  'Center
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
      Index           =   2
      Left            =   2535
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "0000"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   435
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3420
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1755
      MaskColor       =   &H80000004&
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "��ȡ���Ե���������"
      Top             =   3300
      Width           =   960
   End
   Begin VB.Line Line15 
      BorderColor     =   &H000080FF&
      X1              =   345
      X2              =   5880
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Select Item  change Item save Item to database"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   2
      Left            =   315
      TabIndex        =   28
      Top             =   3960
      Width           =   5520
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   4395
      TabIndex        =   27
      Top             =   3960
      Width           =   105
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   75
      X2              =   6195
      Y1              =   3900
      Y2              =   3900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Edit the item here. The item number < 15 characters!"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   45
      Width           =   6240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Index           =   1
      Left            =   2595
      TabIndex        =   24
      Top             =   3420
      Width           =   105
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000000C0&
      X1              =   360
      X2              =   5880
      Y1              =   400
      Y2              =   400
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000000FF&
      X1              =   5880
      X2              =   5880
      Y1              =   400
      Y2              =   1320
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000000FF&
      X1              =   360
      X2              =   5880
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000000FF&
      X1              =   360
      X2              =   360
      Y1              =   1320
      Y2              =   400
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      X1              =   4020
      X2              =   4020
      Y1              =   900
      Y2              =   3180
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   2280
      X2              =   2280
      Y1              =   405
      Y2              =   3180
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      X1              =   360
      X2              =   5880
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   360
      X2              =   5880
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   360
      X2              =   5880
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   5880
      X2              =   5880
      Y1              =   1260
      Y2              =   3180
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   360
      X2              =   5880
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   360
      X2              =   5880
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   360
      X2              =   360
      Y1              =   1260
      Y2              =   3180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Plating time(S)"
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
      Left            =   4080
      TabIndex        =   23
      Top             =   1800
      Width           =   2250
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "B Side"
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
      Index           =   1
      Left            =   4320
      TabIndex        =   22
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "A side"
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
      Index           =   0
      Left            =   2760
      TabIndex        =   21
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Panel area"
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
      Index           =   4
      Left            =   480
      TabIndex        =   20
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Low current"
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
      Left            =   480
      TabIndex        =   19
      Top             =   1800
      Width           =   1650
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Copper current density"
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
      Left            =   0
      TabIndex        =   18
      Top             =   2760
      Width           =   3780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tin current density"
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
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   2010
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enter PN mane"
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
      Left            =   480
      TabIndex        =   16
      Top             =   480
      Width           =   1950
   End
End
Attribute VB_Name = "LAO_eidt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command2_Click()
 Unload Me
End Sub
Private Sub Command3_Click()
 Me.Command1.Enabled = True
 Call EIDT_CHA
End Sub
Private Sub EIDT_CHA()
 Dim A_eidt As Integer
  Text2.Text = Chr(9)
  For i = 0 To 6
   Text1(i).Text = "0000"
  Next
  Me.Height = 7050
  Me.MSFlexGrid1.ColWidth(0) = 2100
  MSFlexGrid1.ColWidth(1) = 2500
  MSFlexGrid1.ColAlignment(0) = 1
  MSFlexGrid1.ColAlignment(1) = 1
  On Error GoTo DQ_ERROR          'erroe
 Dim sq1 As String
 Dim rs_eidt As New ADODB.Recordset
 If Trim(Text4.Text) = "" Then
   ' MsgBox "����������ѯ�ؼ���", 48, "��ʾ"
    sq1 = "select * from �Ϻ����������"
    rs_eidt.Open sq1, conn, adOpenKeyset, adLockPessimistic
    Me.MSFlexGrid1.Rows = rs_eidt.RecordCount
    Me.MSFlexGrid1.Cols = 2
    Me.MSFlexGrid1.TextMatrix(0, 0) = rs_eidt.Fields(16).Name
    Me.MSFlexGrid1.TextMatrix(0, 1) = rs_eidt.Fields(0).Name
    If rs_eidt.RecordCount < 0 Then MsgBox "û�д��Ϻ�", 48, "��ʾ": Exit Sub
    A_eidt = rs_eidt.RecordCount - 1
    rs_eidt.MoveFirst
    If (Not rs_eidt.EOF) Then
      For i = 1 To A_eidt
       MSFlexGrid1.TextMatrix(i, 0) = rs_eidt.Fields(16)
       MSFlexGrid1.TextMatrix(i, 1) = rs_eidt.Fields(0)
       rs_eidt.MoveNext
       Next
         Label4.Caption = "����" & rs_eidt.RecordCount - 1 & "���Ϻ�"
    End If
 Else
    sq1 = "select * from �Ϻ���������� where �Ϻ����� Like '" + Trim(Text4.Text) + "%'"
    rs_eidt.Open sq1, conn, adOpenKeyset, adLockPessimistic
    If rs_eidt.RecordCount = 0 Or rs_eidt.RecordCount < 0 Then
      ' Me.MSFlexGrid1.Visible = False
       Label4.Caption = "����0���Ϻ�":
       MsgBox "û�й���" & Text4.Text & "�Ϻ�", 48, "��ʾ":  Exit Sub
    Else
      Me.MSFlexGrid1.Rows = rs_eidt.RecordCount
      Me.MSFlexGrid1.Cols = 2
      Me.MSFlexGrid1.TextMatrix(0, 0) = rs_eidt.Fields(16).Name
     Me.MSFlexGrid1.TextMatrix(0, 1) = rs_eidt.Fields(0).Name
      A_eidt = rs_eidt.RecordCount - 1
      rs_eidt.MoveFirst
         If (Not rs_eidt.EOF) Then
            For i = 1 To A_eidt
               MSFlexGrid1.TextMatrix(i, 0) = rs_eidt.Fields(16)
               MSFlexGrid1.TextMatrix(i, 1) = rs_eidt.Fields(0)
               rs_eidt.MoveNext
            Next
               Label4.Caption = "����" & rs_eidt.RecordCount - 1 & "���Ϻ�"
        End If
   End If
End If
 'rs_eidt.Close
DQ_ERROR:
 'MsgBox "��ѯ����", 48, "��ʾ"
 Exit Sub

End Sub
Private Sub Command4_Click()
If Dir("C:\WINNT\system32\", vbDirectory) = "" Then
Shell ("C:\WINDOWS\system32\calc.exe")
Else
Shell ("C:\WINNT\system32\calc.exe")
End If
End Sub

Private Sub Command5_Click()
Me.Command1.Enabled = True
For i = 1 To min.MSFlexGrid2.Rows - 1
       min.MSFlexGrid2.Row = i
       min.MSFlexGrid2.Col = 0
If min.MSFlexGrid2.Text = Trim(Text2.Text) Then MsgBox "�Ϻ���ʹ��,ֻ�ܲ鿴���ܸ���!", 48, "��ʾ": Me.Command1.Enabled = False
Next
 Dim sq1 As String
 Dim rs_eidt As New ADODB.Recordset
 On Error GoTo XQ_ERROR
 sq1 = "select * from �Ϻ����������" 'where �����Ϻ�='" + Text2.Text + "'"
 rs_eidt.Open sq1, conn, adOpenKeyset, adLockPessimistic
 rs_eidt.MoveFirst
 While rs_eidt.EOF = False
   If rs_eidt.Fields(0) = Text2.Text Then
      For j = 1 To 7
       Text1(j - 1).Text = rs_eidt.Fields(j)
      Next j
      rs_eidt.Close
      Exit Sub
   Else
     rs_eidt.MoveNext
  End If
Wend
 rs_eidt.Close
XQ_ERROR: Exit Sub
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
For i = 1 To min.MSFlexGrid2.Rows - 1
       min.MSFlexGrid2.Row = i
       min.MSFlexGrid2.Col = 0
If min.MSFlexGrid2.Text = Trim(Text2.Text) Then MsgBox "�Ϻ���ʹ��,�����ڲ��ܶ�����и��ĺ�ɾ������!", 48, "��ʾ": Exit Sub
 Next
 Dim sq1 As String
 Dim rs_eidt As New ADODB.Recordset
 Dim rs_eidt_del As Integer
 Dim RS_EIDT_M As Variant
 Text2.Text = ""
  For i = 0 To 4
   Text1(i).Text = "0000"
  Next
  On Error GoTo EIDT_DEL_ERROR
 RS_EIDT_M = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.RowSel, 1)
 rs_eidt_del = MsgBox("���Ƿ�ɾ���Ϻ�" & RS_EIDT_M & "��?", vbYesNo, "��ʾ")
 If rs_eidt_del = 6 Then
  sq1 = "delete from �Ϻ���������� where �Ϻ�����='" & RS_EIDT_M & "'"
  rs_eidt.Open sq1, conn, adOpenKeyset, adLockPessimistic
  conn.Execute sq1
   MsgBox "���ѳɹ�ɾ���Ϻ�" & RS_EIDT_M, 48, "��ʾ"
 End If
 Call EIDT_CHA
'  rs_eidt.Close
EIDT_DEL_ERROR: Exit Sub
End Sub

Private Sub Form_Load()
Me.Command1.Enabled = True
Me.Height = 4350
Text4.Text = ""
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub


Private Sub MSFlexGrid1_Click()
 Text2.Text = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.RowSel, 1)
 For i = 0 To 6
   Text1(i).Text = "0000"
 Next
End Sub

Private Sub Text2_GotFocus()
 Text2.SelStart = 0
 Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_LostFocus()
 If Trim(Text2.Text) = "" Then
  MsgBox "�������Ϻţ�", 48, "��ʾ": Text2.SetFocus: Exit Sub
 ElseIf Len(Text2.Text) > 16 Then
  MsgBox "�Ϻ����볬����16���ַ���", 48, "��ʾ": Text2.SetFocus: Exit Sub
End If
End Sub
Private Sub Text1_GotFocus(Index As Integer)
  Text1(Index).Text = ""
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim a_a As Integer
Dim a_b As Integer
a_b = 0
    For i = 1 To 4
         If Len(Text1(Index).Text) < 4 Then Text1(Index).Text = "0" & Text1(Index).Text
    Next
     Select Case Index
     Case 0
       If Trim(Text1(0)) = "" Then MsgBox "�ܶ��������Ϊ��!", 48, "����": Text1(0).SetFocus: Exit Sub
       
       If Len(Text1(0).Text) > 4 Then MsgBox "������Χ!", 48, "����": Text1(0).SetFocus: Exit Sub
       For i = 1 To 4
        If Asc(Mid(Text1(0).Text, i, 1)) < 46 Or Asc(Mid(Text1(0).Text, i, 1)) > 57 Or Asc(Mid(Text1(0).Text, i, 1)) = 47 Then MsgBox "��������!", 48, "����": Text1(0).Text = "0000": Text1(0).SetFocus: Exit Sub
          a_a = Asc(Mid(Text1(0), i, 1))
        If a_a = 46 Then a_b = a_b + 1
        If a_b > 1 Then MsgBox "��������!", 48, "����": Text1(1).SetFocus: Exit Sub
        
        Next
     Case 1
       If Trim(Text1(1)) = "" Then MsgBox "�ܶ��������Ϊ��!", 48, "����": Text1(1).SetFocus: Exit Sub
        If Len(Text1(1).Text) > 4 Then MsgBox "������Χ!", 48, "����": Text1(1).SetFocus: Exit Sub
         For i = 1 To 4
          If Asc(Mid(Text1(1).Text, i, 1)) < 46 Or Asc(Mid(Text1(1).Text, i, 1)) > 57 Or Asc(Mid(Text1(1).Text, i, 1)) = 47 Then MsgBox "��������!", 48, "����": Text1(1).SetFocus: Text1(1).SetFocus: Exit Sub
          a_a = Asc(Mid(Text1(1), i, 1))
          If a_a = 46 Then a_b = a_b + 1
          If a_b > 1 Then MsgBox "��������!", 48, "����": Text1(1).SetFocus: Exit Sub
        Next
     Case 2
       If Trim(Text1(2)) = "" Then MsgBox "������͵���ϵ��!", 48, "����": Text1(2).SetFocus: Exit Sub
       If Val(Text1(2).Text) > 1 Then MsgBox "�͵�������ϵ��ֻ����0��1֮�����������[0,1]��!", 48, "����": Text1(2).SetFocus: Exit Sub   '���ܳ���1,ֻ����С��1��С��
        If Len(Text1(2).Text) > 4 Then MsgBox "������Χ!", 48, "����": Text1(2).SetFocus: Exit Sub
          For i = 1 To 4
          If Asc(Mid(Text1(2).Text, i, 1)) < 46 Or Asc(Mid(Text1(2).Text, i, 1)) > 57 Or Asc(Mid(Text1(2).Text, i, 1)) = 47 Then MsgBox "��������!", 48, "����": Text1(2).SetFocus: Text1(2).SetFocus: Exit Sub
          a_a = Asc(Mid(Text1(2), i, 1))
          If a_a = 46 Then a_b = a_b + 1
          If a_b > 1 Then MsgBox "��������!", 48, "����": Text1(2).SetFocus: Exit Sub
        Next
     Case 3
       If Trim(Text1(3)) = "" Then MsgBox "�����������ܶ�!", 48, "����": Text1(3).SetFocus: Exit Sub
        If Len(Text1(3).Text) > 4 Then MsgBox "������Χ!", 48, "����": Text1(3).SetFocus: Exit Sub
         For i = 1 To 4
          If Asc(Mid(Text1(3).Text, i, 1)) < 46 Or Asc(Mid(Text1(3).Text, i, 1)) > 57 Or Asc(Mid(Text1(3).Text, i, 1)) = 47 Then MsgBox "��������!", 48, "����": Text1(3).SetFocus: Text1(3).SetFocus: Exit Sub
          a_a = Asc(Mid(Text1(3), i, 1))
          If a_a = 46 Then a_b = a_b + 1
          If a_b > 1 Then MsgBox "��������!", 48, "����": Text1(3).SetFocus: Exit Sub
        Next
     Case 4
       If Trim(Text1(4)) = "" Then MsgBox "������ͭ���ܶ�!", 48, "����": Text1(4).SetFocus: Exit Sub
        If Len(Text1(4).Text) > 4 Then MsgBox "������Χ!", 48, "����": Text1(4).SetFocus: Exit Sub
          For i = 1 To 4
          If Asc(Mid(Text1(4).Text, i, 1)) < 46 Or Asc(Mid(Text1(4).Text, i, 1)) > 57 Or Asc(Mid(Text1(4).Text, i, 1)) = 47 Then MsgBox "��������!", 48, "����": Text1(4).SetFocus: Text1(4).SetFocus: Exit Sub
          a_a = Asc(Mid(Text1(4), i, 1))
          If a_a = 46 Then a_b = a_b + 1
          If a_b > 1 Then MsgBox "��������!", 48, "����": Text1(4).SetFocus: Exit Sub
        Next
     Case 5
       If Trim(Text1(5)) = "" Then MsgBox "���������ʱ��!", 48, "����": Text1(5).SetFocus: Exit Sub
        If Len(Text1(5).Text) > 4 Then MsgBox "������Χ!", 48, "����": Text1(5).SetFocus: Exit Sub
       For i = 1 To 4
        If Asc(Mid(Text1(5).Text, i, 1)) < 46 Or Asc(Mid(Text1(5).Text, i, 1)) > 57 Or Asc(Mid(Text1(5).Text, i, 1)) = 47 Then MsgBox "��������!", 48, "����": Text1(5).SetFocus: Text1(5).SetFocus: Exit Sub
          a_a = Asc(Mid(Text1(5), i, 1))
          If a_a = 46 Then a_b = a_b + 1
          If a_b > 1 Then MsgBox "��������!", 48, "����": Text1(5).SetFocus: Exit Sub
        Next
     Case 6
       If Trim(Text1(6)) = "" Then MsgBox "�������ͭʱ��!", 48, "����": Text1(6).SetFocus: Exit Sub
        If Len(Text1(6).Text) > 4 Then MsgBox "������Χ!", 48, "����": Text1(6).SetFocus: Exit Sub
       For i = 1 To 4
        If Asc(Mid(Text1(6).Text, i, 1)) < 46 Or Asc(Mid(Text1(6).Text, i, 1)) > 57 Or Asc(Mid(Text1(6).Text, i, 1)) = 47 Then MsgBox "��������!", 48, "����": Text1(6).SetFocus: Text1(6).SetFocus: Exit Sub
          a_a = Asc(Mid(Text1(6), i, 1))
          If a_a = 46 Then a_b = a_b + 1
          If a_b > 1 Then MsgBox "��������!", 48, "����": Text1(6).SetFocus: Exit Sub
        Next
     End Select
   End Sub
Private Sub Command1_Click()
Dim A_MSG As Integer
Dim sq1 As String
 Dim rs_eidt As New ADODB.Recordset
  Dim D_AXH As String
 Dim D_BXH As String
 Dim D_ATH As String
 Dim D_BTH As String
 Dim D_AXL As String
 Dim D_BXL As String
 Dim D_ATL As String
 Dim D_BTL As String
 D_AXH = Trim(Val(Text1(0).Text) * Val(Text1(3).Text))
  D_BXH = Trim(Val(Text1(1).Text) * Val(Text1(3).Text))
 D_AXL = Trim((Val(Text1(0).Text) * Val(Text1(3).Text)) * Val(Text1(2).Text))
 D_BXL = Trim((Val(Text1(1).Text) * Val(Text1(3).Text)) * Val(Text1(2).Text))
D_ATH = Trim(Val(Text1(0).Text) * Val(Text1(4).Text))
 D_BTH = Trim(Val(Text1(1).Text) * Val(Text1(4).Text))
 D_ATL = Trim((Val(Text1(0).Text) * Val(Text1(4).Text)) * Val(Text1(2).Text))
 D_BTL = Trim((Val(Text1(1).Text) * Val(Text1(4).Text)) * Val(Text1(2).Text))
 'D_xs = Trim(Text1(2).Text)
 'D_BTL = Trim((Val(text1(1).Text) * Val(text1(4).Text)) * Val(text1(2).Text))
 On Error GoTo eidt_error
 Text3.Text = Now
 sq1 = "select * from �Ϻ����������"
 rs_eidt.Open sq1, conn, adOpenKeyset, adLockPessimistic
 If rs_eidt.EOF = False Then
 While (rs_eidt.EOF = False)
       If Trim(rs_eidt.Fields(0)) = Trim(Text2.Text) Then
          A_MSG = MsgBox("���д��Ϻ����ϣ��Ƿ�Ҫ������ǰ���ϺŲ�������[��]�˳�����[��]����", vbYesNo, "��ʾ")
          If A_MSG = vbNo Then Text2.SetFocus: Exit Sub
          If A_MSG = vbYes Then GoTo UP_EIDT
        Else
          rs_eidt.MoveNext
        End If
 Wend
    ' MoveLast
     rs_eidt.AddNew
       rs_eidt.Fields(0) = Text2.Text
       For i = 0 To 6
       rs_eidt.Fields(i + 1) = Text1(i).Text
       Next
       rs_eidt.Fields(8) = Format(D_AXH, "000.0")
       rs_eidt.Fields(9) = Format(D_BXH, "000.0")
       rs_eidt.Fields(10) = Format(D_AXL, "000.0")
       rs_eidt.Fields(11) = Format(D_BXL, "000.0")
       rs_eidt.Fields(12) = Format(D_ATH, "000.0")
       rs_eidt.Fields(13) = Format(D_BTH, "000.0")
       rs_eidt.Fields(14) = Format(D_ATL, "000.0")
       rs_eidt.Fields(15) = Format(D_BTL, "000.0")
       rs_eidt.Fields(16) = Text3.Text
       rs_eidt.Update
       rs_eidt.Close
       Call L_add
       MsgBox "�Ϻű༭�ɹ�!", 48, "��ʾ"
       Text2.SetFocus
       Text2.Text = "0"
       For i = 0 To 6
        Text1(i) = "0000"
       Next
       End If
       Exit Sub
UP_EIDT:
        rs_eidt.Fields(0) = Text2.Text
        For i = 0 To 6
        rs_eidt.Fields(i + 1) = Text1(i).Text
         Next
       rs_eidt.Fields(8) = Format(D_AXH, "000.0")
       rs_eidt.Fields(9) = Format(D_BXH, "000.0")
       rs_eidt.Fields(10) = Format(D_AXL, "000.0")
       rs_eidt.Fields(11) = Format(D_BXL, "000.0")
       rs_eidt.Fields(12) = Format(D_ATH, "000.0")
       rs_eidt.Fields(13) = Format(D_BTH, "000.0")
       rs_eidt.Fields(14) = Format(D_ATL, "000.0")
       rs_eidt.Fields(15) = Format(D_BTL, "000.0")
       rs_eidt.Fields(16) = Text3.Text
       rs_eidt.Update
       rs_eidt.Close
       Call L_up
       MsgBox "�Ϻ��޸ĳɹ�!", 48, "��ʾ"
       Text2.SetFocus
       Text2.Text = "0"
       For i = 0 To 6
        Text1(i) = "0000"
       Next
       Exit Sub
eidt_error:
If Err > 0 Then MsgBox Str(Err) & Error
End Sub

Sub L_add()
 Open (App.Path & "\style\�Ϻű༭.ini") For Input As #2
                          Do While Not EOF(2)
                               Line Input #2, INTEXT
                               TSTR = TSTR + INTEXT + Chr(13) + Chr(10)
                          Loop
                      Close #2
                      TSTR = TSTR + "   " + Text2.Text + "               " + Format(Now, "yyyy-mm-dd hh:mm:ss") + "            " + "�Ϻ��½�" + Chr(13) + Chr(10)
                      If Len(TSTR) > 10000 Then TSTR = Right(TSTR, 9800)
                      Open (App.Path & "\style\�Ϻű༭.ini") For Output As #2
                            Print #2, TSTR
                      Close #2

End Sub
Sub L_up()
 Open (App.Path & "\style\�Ϻű༭.ini") For Input As #2
                          Do While Not EOF(2)
                               Line Input #2, INTEXT
                               TSTR = TSTR + INTEXT + Chr(13) + Chr(10)
                          Loop
                      Close #2
                      If Len(TSTR) > 10000 Then TSTR = Right(TSTR, 9800)
                      TSTR = TSTR + "   " + Text2.Text + "               " + Format(Now, "yyyy-mm-dd hh:mm:ss") + "            " + "�Ϻ��޸�" + Chr(13) + Chr(10)
                      Open (App.Path & "\style\�Ϻű༭.ini") For Output As #2
                            Print #2, TSTR
                      Close #2
End Sub

Private Sub Timer1_Timer()
 Me.Command1.Enabled = True
 Call EIDT_CHA
 Me.Timer1 = False
End Sub
