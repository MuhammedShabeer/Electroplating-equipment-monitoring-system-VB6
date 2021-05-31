VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form JI_LOCK 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form1"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   Icon            =   "JI_LOCK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   9000
   StartUpPosition =   2  '屏幕中心
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "JI_LOCK.frx":0442
      Height          =   7815
      Left            =   50
      TabIndex        =   3
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   13785
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Check3"
      Height          =   255
      Left            =   2640
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   9
      Top             =   8470
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Check1"
      Height          =   180
      Left            =   120
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   8
      Top             =   8470
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "记录整理"
      Enabled         =   0   'False
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
      Left            =   7560
      TabIndex        =   7
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
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
      Left            =   5400
      TabIndex        =   5
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
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
      Left            =   6600
      TabIndex        =   2
      Top             =   8400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Inquire"
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
      Left            =   4320
      TabIndex        =   1
      Top             =   8400
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   8445
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   21430273
      CurrentDate     =   39113
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DMS数据库.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DMS数据库.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from 故障表"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   8250
      Width           =   9000
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "JI_LOCK.frx":0457
         Left            =   3600
         List            =   "JI_LOCK.frx":0479
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Tank ："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   2880
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Date："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   630
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6240
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DMS数据库.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DMS数据库.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from 电镀记录"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   6240
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DMS数据库.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DMS数据库.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF80FF&
      Caption         =   "此生产记录仅作参考，只有在自动情况下才可能准确！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   5760
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "JI_LOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Check1_Click()
Me.Check1.CausesValidation = False
End Sub


Private Sub Combo1_Change()
'13297581286
End Sub

Private Sub Command1_Click()
 Me.Command5.Enabled = False
 Dim like_sql As String
On Error Resume Next
If Me.Caption = "故障记录" Then
     Me.Check1.Value = 1
     Me.Adodc1.RecordSource = "select * from 故障表 where 时间 Like '" + Trim(DTPicker1.Value) + "%'"
     Me.Adodc1.Refresh
     JI_LOCK.Label2.Caption = "共有" & JI_LOCK.Adodc1.Recordset.RecordCount & "条记录"
     'Call d_lock
Else
   If Me.Check1.Value = 0 And Me.Check3.Value = 0 Then MsgBox "请钩取一个查询条件！", , "提示": Exit Sub
   If Me.Check1.Value = 1 And Me.Check3.Value = 1 Then MsgBox "只能按一个条件查询！", , "提示": Exit Sub
     If Me.Check1.Value = 1 Then
       like_time = Year(DTPicker1.Value) & "年" & Val(Month(DTPicker1.Value)) & "月" & Day(DTPicker1.Value) & "日"
       'Me.Adodc1.RecordSource = "select * from 电镀记录 where 入缸时间 Like '" + Trim(DTPicker1.Value) + "%'"
       like_sql = "select * from" & Chr(9) & like_time
       Me.Adodc1.RecordSource = like_sql
       Me.Adodc1.Refresh
       JI_LOCK.Label2.Caption = "共有" & JI_LOCK.Adodc1.Recordset.RecordCount & "条记录"
       Call d_lock
       look_b = False
     Else
       If look_b = False Then
          MsgBox "您现在查看的是" & like_time & "的记录", , "提示"
          like_sql = "select * from" & Chr(9) & like_time & Chr(9) & "where 槽号 like '" + Trim(Me.Combo1.Text) + "%'"
          Me.Adodc1.RecordSource = like_sql
          ' a11111 = Trim(Me.Combo1.Text)
          Me.Adodc1.Refresh
          JI_LOCK.Label2.Caption = "共有" & JI_LOCK.Adodc1.Recordset.RecordCount & "条记录"
         Call d_lock
       Else
          like_sql = "select * from 电镀记录 where 槽号 like '" + Trim(Me.Combo1.Text) + "%'"
          Me.Adodc1.RecordSource = like_sql
          ' a11111 = Trim(Me.Combo1.Text)
          Me.Adodc1.Refresh
          JI_LOCK.Label2.Caption = "共有" & JI_LOCK.Adodc1.Recordset.RecordCount & "条记录"
         Call d_lock
       End If
    End If
End If
If Me.Check1.Value = 1 Then Me.Check1.Value = 0
If Me.Check3.Value = 1 Then Me.Check3.Value = 0
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If JI_LOCK.Caption = "故障记录" Then DataReport1.Show Else MsgBox "不完全", , ""
'DataReport1.Show
Unload Me
End Sub

Private Sub Command4_Click()
Dim a_input As Integer
On Error Resume Next
If userpow <> "系统管理" Then MsgBox "你没有权限删除记录", 48, "提示": Exit Sub
a_input = InputBox("请输入密码", "密码", "")
If a_input <> userpass Then MsgBox "密码错误", 16, "提示": Exit Sub
 If Me.Caption = "故障记录" Then
    Me.Adodc1.RecordSource = "select * from 故障表"
    Me.Refresh
    While Not Me.Adodc1.Recordset.EOF = True
    Me.Adodc1.Recordset.Delete
    Me.Adodc1.Recordset.MoveNext
    Wend
    Me.Refresh
    min.Adodc3.Refresh
    If JI_LOCK.Caption = "生产记录" Then Call d_lock
    JI_LOCK.Label2.Caption = "共有" & JI_LOCK.Adodc1.Recordset.RecordCount & "条记录"
 Else
    Me.Adodc1.RecordSource = "select * from 电镀记录"
    Me.Refresh
    While Not Me.Adodc1.Recordset.EOF = True
    Me.Adodc1.Recordset.Delete
    Me.Adodc1.Recordset.MoveNext
    Wend
    Me.Refresh
    If JI_LOCK.Caption = "生产记录" Then Call d_lock
    JI_LOCK.Label2.Caption = "共有" & JI_LOCK.Adodc1.Recordset.RecordCount & "条记录"
 End If
End Sub

Private Sub Command5_Click()
Me.Command1.Enabled = True
bbbb = False
bbbb_1 = True
On Error GoTo cat_err
Dim connectionstring As String
Dim cat_1 As New ADODB.Recordset
Dim sqlcreate As String
Call year_date_time
connectionstring = "provider=Microsoft.Jet.oledb.4.0;" & _
                   "data source=DMS数据库.mdb" '打开数据库.
Set cat = New ADODB.Connection
 cat.Open connectionstring
sqlcreate = "Create Table" & Chr(9) & Data_1 & "(ID char(5),槽号 char(3) not null ,料号 char(16) not null ,数" & _
            "量 char(3) not null ,A面电流 char(4) not null,B面电流 char(4) not null,电镀时间 char(20) not null,入" & _
            "缸时间 char(20) not null,出缸时间 char(20) not null,安培小时 char(4) not null)"
Set cat_1 = cat.Execute(sqlcreate, 1, adCmdText)
Set cat_1 = Nothing
Call cat_input
cat_err:
   MsgBox Err.Description & "此记录已整理完毕!"
End Sub

Private Sub Form_Load()
Dim i As Integer
Me.DTPicker1.Value = Year(Now) & "-" & Month(Now) & "-" & Day(Now)
Combo1.Clear
Combo1.Text = "9"
For i = 9 To 12
  Combo1.AddItem CStr(i)
Next
For i = 23 To 38
  Combo1.AddItem CStr(i)
Next
look_b = True
End Sub

Private Sub cat_input()
 Dim i As Integer
 Dim cat_input_1 As String
 Dim yestaday As String
  yestaday = Day(Data_1)
  cat_input_1 = "select * from" & Chr(9) & Data_1
  Me.Adodc3.RecordSource = cat_input_1
  Me.Adodc2.Refresh
  Me.Adodc3.Refresh
  Me.Adodc2.Recordset.MoveFirst
  While Not Me.Adodc2.Recordset.EOF = True
    If Day(Me.Adodc2.Recordset.Fields(7)) = yestaday Then
       Me.Adodc3.Recordset.AddNew
       For i = 1 To 9
           Me.Adodc3.Recordset.Fields(i) = Me.Adodc2.Recordset.Fields(i)
       Next
       Me.Adodc3.Recordset.Update
       Me.Adodc2.Recordset.Delete
       Me.Adodc2.Recordset.MoveNext
    Else
      Me.Adodc2.Recordset.MoveNext
    End If
  Wend
End Sub
