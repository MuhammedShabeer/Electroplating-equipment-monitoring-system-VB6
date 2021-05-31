VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form min 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Electroplating equipment monitoring system"
   ClientHeight    =   10815
   ClientLeft      =   165
   ClientTop       =   615
   ClientWidth     =   15240
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF80FF&
   Icon            =   "min.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "min.frx":0442
   ScaleHeight     =   10815
   ScaleMode       =   0  'User
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab2 
      Height          =   2235
      Left            =   8280
      TabIndex        =   57
      Top             =   4080
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   3942
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      WordWrap        =   0   'False
      BackColor       =   8421631
      TabCaption(0)   =   "Choose tank"
      TabPicture(0)   =   "min.frx":0D0C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(12)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(14)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text11(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text11(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text11(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Check1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Check1(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Check1(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Check1(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Check1(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Check1(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Check1(6)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Check1(7)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Check1(8)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Check1(9)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Check1(10)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Check1(11)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Check1(12)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Check1(13)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Vibrative"
      TabPicture(1)   =   "min.frx":0D28
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5(0)"
      Tab(1).Control(1)=   "Label5(1)"
      Tab(1).Control(2)=   "Label1(16)"
      Tab(1).Control(3)=   "Label1(17)"
      Tab(1).Control(4)=   "Check2(0)"
      Tab(1).Control(5)=   "Check2(1)"
      Tab(1).Control(6)=   "Check2(2)"
      Tab(1).Control(7)=   "Check2(3)"
      Tab(1).Control(8)=   "Check2(4)"
      Tab(1).Control(9)=   "Check2(5)"
      Tab(1).Control(10)=   "Check2(6)"
      Tab(1).Control(11)=   "Check2(7)"
      Tab(1).Control(12)=   "Check2(8)"
      Tab(1).Control(13)=   "Check2(9)"
      Tab(1).Control(14)=   "Check2(10)"
      Tab(1).Control(15)=   "Check2(11)"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Temperature"
      TabPicture(2)   =   "min.frx":0D44
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Check3(0)"
      Tab(2).Control(1)=   "Check3(1)"
      Tab(2).Control(2)=   "Check3(2)"
      Tab(2).Control(3)=   "Check3(3)"
      Tab(2).Control(4)=   "Check3(4)"
      Tab(2).Control(5)=   "Check3(5)"
      Tab(2).Control(6)=   "Check3(6)"
      Tab(2).Control(7)=   "Check3(7)"
      Tab(2).Control(8)=   "Check3(8)"
      Tab(2).Control(9)=   "Check3(9)"
      Tab(2).Control(10)=   "Check3(10)"
      Tab(2).Control(11)=   "Check3(11)"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Filter"
      TabPicture(3)   =   "min.frx":0D60
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Check4(0)"
      Tab(3).Control(1)=   "Check4(1)"
      Tab(3).Control(2)=   "Check4(2)"
      Tab(3).Control(3)=   "Check4(3)"
      Tab(3).Control(4)=   "Check4(4)"
      Tab(3).Control(5)=   "Check4(5)"
      Tab(3).Control(6)=   "Check4(6)"
      Tab(3).Control(7)=   "Check4(7)"
      Tab(3).Control(8)=   "Check4(8)"
      Tab(3).Control(9)=   "Check4(9)"
      Tab(3).Control(10)=   "Check4(10)"
      Tab(3).Control(11)=   "Check4(11)"
      Tab(3).Control(12)=   "Check4(12)"
      Tab(3).Control(13)=   "Check4(13)"
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Add 1"
      TabPicture(4)   =   "min.frx":0D7C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1(18)"
      Tab(4).Control(1)=   "Label1(19)"
      Tab(4).Control(2)=   "Label1(20)"
      Tab(4).Control(3)=   "Label1(21)"
      Tab(4).Control(4)=   "Label1(22)"
      Tab(4).Control(5)=   "Label1(23)"
      Tab(4).Control(6)=   "Label1(24)"
      Tab(4).Control(7)=   "Label1(25)"
      Tab(4).Control(8)=   "Label1(26)"
      Tab(4).Control(9)=   "Label1(27)"
      Tab(4).Control(10)=   "Label1(28)"
      Tab(4).Control(11)=   "Label1(29)"
      Tab(4).Control(12)=   "Label6(0)"
      Tab(4).Control(13)=   "Label6(1)"
      Tab(4).Control(14)=   "Label6(2)"
      Tab(4).Control(15)=   "Label6(3)"
      Tab(4).Control(16)=   "Label6(4)"
      Tab(4).Control(17)=   "Label6(5)"
      Tab(4).Control(18)=   "Label7(0)"
      Tab(4).Control(19)=   "Label7(1)"
      Tab(4).Control(20)=   "Label7(2)"
      Tab(4).Control(21)=   "Label7(3)"
      Tab(4).Control(22)=   "Label7(4)"
      Tab(4).Control(23)=   "Label7(5)"
      Tab(4).Control(24)=   "Label8(0)"
      Tab(4).Control(25)=   "Label8(1)"
      Tab(4).Control(26)=   "Label8(2)"
      Tab(4).Control(27)=   "Label8(3)"
      Tab(4).Control(28)=   "Label8(4)"
      Tab(4).Control(29)=   "Label8(5)"
      Tab(4).Control(30)=   "Label9(0)"
      Tab(4).Control(31)=   "Label9(1)"
      Tab(4).Control(32)=   "Label9(2)"
      Tab(4).Control(33)=   "Label9(3)"
      Tab(4).Control(34)=   "Label9(4)"
      Tab(4).Control(35)=   "Label9(5)"
      Tab(4).Control(36)=   "Line2(0)"
      Tab(4).Control(37)=   "Line3(0)"
      Tab(4).Control(38)=   "Line4(0)"
      Tab(4).Control(39)=   "Line5(0)"
      Tab(4).Control(40)=   "Line6(0)"
      Tab(4).Control(41)=   "Line7(0)"
      Tab(4).Control(42)=   "Line8(0)"
      Tab(4).Control(43)=   "Line9(0)"
      Tab(4).Control(44)=   "Line10(0)"
      Tab(4).Control(45)=   "Line11(0)"
      Tab(4).Control(46)=   "Line12(0)"
      Tab(4).Control(47)=   "Check5(0)"
      Tab(4).Control(48)=   "Check5(1)"
      Tab(4).Control(49)=   "Check5(2)"
      Tab(4).Control(50)=   "Check5(3)"
      Tab(4).Control(51)=   "Check5(4)"
      Tab(4).Control(52)=   "Check5(5)"
      Tab(4).ControlCount=   53
      TabCaption(5)   =   "Add 2"
      TabPicture(5)   =   "min.frx":0D98
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label1(30)"
      Tab(5).Control(1)=   "Label1(31)"
      Tab(5).Control(2)=   "Label1(32)"
      Tab(5).Control(3)=   "Label1(33)"
      Tab(5).Control(4)=   "Label1(34)"
      Tab(5).Control(5)=   "Label1(35)"
      Tab(5).Control(6)=   "Label1(36)"
      Tab(5).Control(7)=   "Label1(37)"
      Tab(5).Control(8)=   "Label1(38)"
      Tab(5).Control(9)=   "Label1(39)"
      Tab(5).Control(10)=   "Label1(40)"
      Tab(5).Control(11)=   "Label1(41)"
      Tab(5).Control(12)=   "Label6(6)"
      Tab(5).Control(13)=   "Label6(7)"
      Tab(5).Control(14)=   "Label6(8)"
      Tab(5).Control(15)=   "Label6(9)"
      Tab(5).Control(16)=   "Label6(10)"
      Tab(5).Control(17)=   "Label6(11)"
      Tab(5).Control(18)=   "Label7(6)"
      Tab(5).Control(19)=   "Label7(7)"
      Tab(5).Control(20)=   "Label7(8)"
      Tab(5).Control(21)=   "Label7(9)"
      Tab(5).Control(22)=   "Label7(10)"
      Tab(5).Control(23)=   "Label7(11)"
      Tab(5).Control(24)=   "Label8(6)"
      Tab(5).Control(25)=   "Label8(7)"
      Tab(5).Control(26)=   "Label8(8)"
      Tab(5).Control(27)=   "Label8(9)"
      Tab(5).Control(28)=   "Label8(10)"
      Tab(5).Control(29)=   "Label8(11)"
      Tab(5).Control(30)=   "Label9(6)"
      Tab(5).Control(31)=   "Label9(7)"
      Tab(5).Control(32)=   "Label9(8)"
      Tab(5).Control(33)=   "Label9(9)"
      Tab(5).Control(34)=   "Label9(10)"
      Tab(5).Control(35)=   "Label9(11)"
      Tab(5).Control(36)=   "Line2(1)"
      Tab(5).Control(37)=   "Line3(1)"
      Tab(5).Control(38)=   "Line4(1)"
      Tab(5).Control(39)=   "Line5(1)"
      Tab(5).Control(40)=   "Line6(1)"
      Tab(5).Control(41)=   "Line7(1)"
      Tab(5).Control(42)=   "Line8(1)"
      Tab(5).Control(43)=   "Line9(1)"
      Tab(5).Control(44)=   "Line10(1)"
      Tab(5).Control(45)=   "Line11(1)"
      Tab(5).Control(46)=   "Line12(1)"
      Tab(5).Control(47)=   "Check5(6)"
      Tab(5).Control(48)=   "Check5(7)"
      Tab(5).Control(49)=   "Check5(8)"
      Tab(5).Control(50)=   "Check5(9)"
      Tab(5).Control(51)=   "Check5(10)"
      Tab(5).Control(52)=   "Check5(11)"
      Tab(5).ControlCount=   53
      TabCaption(6)   =   "Alarm"
      TabPicture(6)   =   "min.frx":0DB4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Line27"
      Tab(6).Control(1)=   "Line26"
      Tab(6).Control(2)=   "Line25"
      Tab(6).Control(3)=   "Line24"
      Tab(6).Control(4)=   "Line23"
      Tab(6).Control(5)=   "Line22"
      Tab(6).Control(6)=   "Line21"
      Tab(6).Control(7)=   "Line20"
      Tab(6).Control(8)=   "Line19"
      Tab(6).Control(9)=   "Line18"
      Tab(6).Control(10)=   "Line17"
      Tab(6).Control(11)=   "Line16"
      Tab(6).Control(12)=   "Line15"
      Tab(6).Control(13)=   "Line14"
      Tab(6).Control(14)=   "Line13"
      Tab(6).Control(15)=   "Label13(13)"
      Tab(6).Control(16)=   "Label13(12)"
      Tab(6).Control(17)=   "Label11(3)"
      Tab(6).Control(18)=   "Label11(2)"
      Tab(6).Control(19)=   "Label13(11)"
      Tab(6).Control(20)=   "Label12(11)"
      Tab(6).Control(21)=   "Label13(10)"
      Tab(6).Control(22)=   "Label12(10)"
      Tab(6).Control(23)=   "Label13(9)"
      Tab(6).Control(24)=   "Label12(9)"
      Tab(6).Control(25)=   "Label13(8)"
      Tab(6).Control(26)=   "Label12(8)"
      Tab(6).Control(27)=   "Label13(7)"
      Tab(6).Control(28)=   "Label12(7)"
      Tab(6).Control(29)=   "Label13(6)"
      Tab(6).Control(30)=   "Label12(6)"
      Tab(6).Control(31)=   "Label13(5)"
      Tab(6).Control(32)=   "Label12(5)"
      Tab(6).Control(33)=   "Label13(4)"
      Tab(6).Control(34)=   "Label12(4)"
      Tab(6).Control(35)=   "Label13(3)"
      Tab(6).Control(36)=   "Label12(3)"
      Tab(6).Control(37)=   "Label13(2)"
      Tab(6).Control(38)=   "Label12(2)"
      Tab(6).Control(39)=   "Label13(1)"
      Tab(6).Control(40)=   "Label12(1)"
      Tab(6).Control(41)=   "Label13(0)"
      Tab(6).Control(42)=   "Label12(0)"
      Tab(6).Control(43)=   "Label11(1)"
      Tab(6).Control(44)=   "Label11(0)"
      Tab(6).Control(45)=   "Label10"
      Tab(6).ControlCount=   46
      Begin VB.CheckBox Check1 
         BackColor       =   &H00008080&
         Caption         =   "CU 10"
         Height          =   285
         Index           =   13
         Left            =   1740
         TabIndex        =   203
         Top             =   1350
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00008080&
         Caption         =   "CU 9"
         Height          =   285
         Index           =   12
         Left            =   105
         TabIndex        =   202
         Top             =   1350
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00008080&
         Caption         =   "CU 8"
         Height          =   285
         Index           =   11
         Left            =   5040
         TabIndex        =   201
         Top             =   1035
         Width           =   1620
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00008080&
         Caption         =   "CU 7"
         Height          =   285
         Index           =   10
         Left            =   3405
         TabIndex        =   200
         Top             =   1035
         Width           =   1605
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00008080&
         Caption         =   "CU 6"
         Height          =   285
         Index           =   9
         Left            =   1755
         TabIndex        =   199
         Top             =   1035
         Width           =   1620
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00008080&
         Caption         =   "CU 5"
         Height          =   285
         Index           =   8
         Left            =   105
         TabIndex        =   198
         Top             =   1035
         Width           =   1605
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00008080&
         Caption         =   "CU 4"
         Height          =   285
         Index           =   7
         Left            =   5040
         TabIndex        =   197
         Top             =   735
         Width           =   1620
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00008080&
         Caption         =   "CU 3"
         Height          =   285
         Index           =   6
         Left            =   3405
         TabIndex        =   196
         Top             =   735
         Width           =   1605
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00008080&
         Caption         =   "CU 2"
         Height          =   285
         Index           =   5
         Left            =   1755
         TabIndex        =   195
         Top             =   720
         Width           =   1620
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00008080&
         Caption         =   "CU 1"
         Height          =   285
         Index           =   4
         Left            =   105
         TabIndex        =   194
         Top             =   720
         Width           =   1605
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C000&
         Caption         =   "Tin-s"
         Height          =   285
         Index           =   3
         Left            =   5040
         TabIndex        =   193
         Top             =   405
         Width           =   1620
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000C000&
         Caption         =   "Tin-1"
         Height          =   285
         Index           =   2
         Left            =   3405
         TabIndex        =   192
         Top             =   405
         Width           =   1605
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H008080FF&
         Caption         =   "5 tank "
         Height          =   285
         Index           =   1
         Left            =   1755
         TabIndex        =   191
         Top             =   405
         Width           =   1620
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H008080FF&
         Caption         =   "4tank "
         Height          =   285
         Index           =   0
         Left            =   105
         TabIndex        =   190
         Top             =   405
         Width           =   1605
      End
      Begin VB.CheckBox Check5 
         Caption         =   "强加"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   -69090
         TabIndex        =   165
         Top             =   1605
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CheckBox Check5 
         Caption         =   "强加"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   -70035
         TabIndex        =   164
         Top             =   1605
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CheckBox Check5 
         Caption         =   "强加"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   -71010
         TabIndex        =   163
         Top             =   1605
         Width           =   690
      End
      Begin VB.CheckBox Check5 
         Caption         =   "强加"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   -72015
         TabIndex        =   162
         Top             =   1605
         Width           =   690
      End
      Begin VB.CheckBox Check5 
         Caption         =   "强加"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   -72975
         TabIndex        =   161
         Top             =   1605
         Width           =   690
      End
      Begin VB.CheckBox Check5 
         Caption         =   "强加"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   -73905
         TabIndex        =   160
         Top             =   1605
         Width           =   690
      End
      Begin VB.CheckBox Check5 
         Caption         =   "强加"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   -69090
         TabIndex        =   123
         Top             =   1605
         Width           =   690
      End
      Begin VB.CheckBox Check5 
         Caption         =   "强加"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   -70035
         TabIndex        =   122
         Top             =   1605
         Width           =   690
      End
      Begin VB.CheckBox Check5 
         Caption         =   "强加"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   -71010
         TabIndex        =   121
         Top             =   1605
         Width           =   690
      End
      Begin VB.CheckBox Check5 
         Caption         =   "强加"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   -72015
         TabIndex        =   120
         Top             =   1605
         Width           =   690
      End
      Begin VB.CheckBox Check5 
         Caption         =   "强加"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   -72975
         TabIndex        =   119
         Top             =   1605
         Width           =   690
      End
      Begin VB.CheckBox Check5 
         Caption         =   "强加"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   -73905
         TabIndex        =   118
         Top             =   1605
         Width           =   690
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808000&
         Caption         =   "Air agetation"
         Height          =   285
         Index           =   13
         Left            =   -73260
         TabIndex        =   105
         Top             =   1590
         Width           =   1620
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808000&
         Caption         =   "Replating pump"
         Height          =   285
         Index           =   12
         Left            =   -74895
         TabIndex        =   104
         Top             =   1575
         Width           =   1605
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H000040C0&
         Caption         =   "CU8"
         Height          =   285
         Index           =   11
         Left            =   -69930
         TabIndex        =   103
         Top             =   1215
         Width           =   1620
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H000040C0&
         Caption         =   "CU7"
         Height          =   285
         Index           =   10
         Left            =   -71580
         TabIndex        =   102
         Top             =   1215
         Width           =   1605
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H000040C0&
         Caption         =   "CU6"
         Height          =   285
         Index           =   9
         Left            =   -73245
         TabIndex        =   101
         Top             =   1230
         Width           =   1620
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H000040C0&
         Caption         =   "CU5"
         Height          =   285
         Index           =   8
         Left            =   -74895
         TabIndex        =   100
         Top             =   1215
         Width           =   1605
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808000&
         Caption         =   "CU4"
         Height          =   285
         Index           =   7
         Left            =   -69930
         TabIndex        =   99
         Top             =   870
         Width           =   1620
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808000&
         Caption         =   "CU3"
         Height          =   285
         Index           =   6
         Left            =   -71580
         TabIndex        =   98
         Top             =   870
         Width           =   1605
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808000&
         Caption         =   "CU2"
         Height          =   285
         Index           =   5
         Left            =   -73260
         TabIndex        =   97
         Top             =   870
         Width           =   1620
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808000&
         Caption         =   "CU1"
         Height          =   285
         Index           =   4
         Left            =   -74910
         TabIndex        =   96
         Top             =   855
         Width           =   1605
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H000040C0&
         Caption         =   "Etching"
         Height          =   285
         Index           =   3
         Left            =   -69930
         TabIndex        =   95
         Top             =   495
         Width           =   1620
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H000040C0&
         Caption         =   "unoil"
         Height          =   285
         Index           =   2
         Left            =   -71580
         TabIndex        =   94
         Top             =   495
         Width           =   1605
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H000040C0&
         Caption         =   "Tin 2"
         Height          =   285
         Index           =   1
         Left            =   -73245
         TabIndex        =   93
         Top             =   510
         Width           =   1620
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H000040C0&
         Caption         =   "Tin 1"
         Height          =   285
         Index           =   0
         Left            =   -74895
         TabIndex        =   92
         Top             =   495
         Width           =   1605
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C000C0&
         Caption         =   "CU8"
         Height          =   285
         Index           =   11
         Left            =   -69930
         TabIndex        =   87
         Top             =   1215
         Width           =   1620
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C000C0&
         Caption         =   "CU7"
         Height          =   285
         Index           =   10
         Left            =   -71580
         TabIndex        =   86
         Top             =   1215
         Width           =   1605
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C000C0&
         Caption         =   "CU6"
         Height          =   285
         Index           =   9
         Left            =   -73245
         TabIndex        =   85
         Top             =   1230
         Width           =   1620
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C000C0&
         Caption         =   "CU5"
         Height          =   285
         Index           =   8
         Left            =   -74895
         TabIndex        =   84
         Top             =   1215
         Width           =   1605
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0000C000&
         Caption         =   "CU4"
         Height          =   285
         Index           =   7
         Left            =   -69930
         TabIndex        =   83
         Top             =   870
         Width           =   1620
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0000C000&
         Caption         =   "CU3"
         Height          =   285
         Index           =   6
         Left            =   -71580
         TabIndex        =   82
         Top             =   870
         Width           =   1605
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0000C000&
         Caption         =   "CU 2"
         Height          =   285
         Index           =   5
         Left            =   -73260
         TabIndex        =   81
         Top             =   870
         Width           =   1620
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0000C000&
         Caption         =   "CU 1"
         Height          =   285
         Index           =   4
         Left            =   -74910
         TabIndex        =   80
         Top             =   855
         Width           =   1605
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C000C0&
         Caption         =   "Etching"
         Height          =   285
         Index           =   3
         Left            =   -69930
         TabIndex        =   79
         Top             =   495
         Width           =   1620
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C000C0&
         Caption         =   "unoil"
         Height          =   285
         Index           =   2
         Left            =   -71580
         TabIndex        =   78
         Top             =   495
         Width           =   1605
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C000C0&
         Caption         =   "Tin 2"
         Height          =   285
         Index           =   1
         Left            =   -73245
         TabIndex        =   77
         Top             =   510
         Width           =   1620
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C000C0&
         Caption         =   "Tin 1"
         Height          =   285
         Index           =   0
         Left            =   -74895
         TabIndex        =   76
         Top             =   495
         Width           =   1605
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0000C000&
         Caption         =   "CU 8"
         Height          =   285
         Index           =   11
         Left            =   -69960
         TabIndex        =   75
         Top             =   1035
         Width           =   1620
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0000C000&
         Caption         =   "CU 7"
         Height          =   285
         Index           =   10
         Left            =   -71595
         TabIndex        =   74
         Top             =   1035
         Width           =   1605
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0000C000&
         Caption         =   "Cu 6"
         Height          =   285
         Index           =   9
         Left            =   -73245
         TabIndex        =   73
         Top             =   1035
         Width           =   1620
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0000C000&
         Caption         =   "Cu 5"
         Height          =   285
         Index           =   8
         Left            =   -74895
         TabIndex        =   72
         Top             =   1035
         Width           =   1605
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FF8080&
         Caption         =   "Cu 4"
         Height          =   285
         Index           =   7
         Left            =   -69960
         TabIndex        =   71
         Top             =   735
         Width           =   1620
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FF8080&
         Caption         =   "Cu 3"
         Height          =   285
         Index           =   6
         Left            =   -71595
         TabIndex        =   70
         Top             =   735
         Width           =   1605
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FF8080&
         Caption         =   "Cu 2"
         Height          =   285
         Index           =   5
         Left            =   -73245
         TabIndex        =   69
         Top             =   720
         Width           =   1620
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FF8080&
         Caption         =   "cu 1"
         Height          =   285
         Index           =   4
         Left            =   -74895
         TabIndex        =   68
         Top             =   720
         Width           =   1605
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0000C000&
         Caption         =   "Etch"
         Height          =   285
         Index           =   3
         Left            =   -69960
         TabIndex        =   67
         Top             =   405
         Width           =   1620
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0000C000&
         Caption         =   "Unoil"
         Height          =   285
         Index           =   2
         Left            =   -71595
         TabIndex        =   66
         Top             =   405
         Width           =   1605
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0000C000&
         Caption         =   "Tin2"
         Height          =   285
         Index           =   1
         Left            =   -73245
         TabIndex        =   65
         Top             =   405
         Width           =   1620
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0000C000&
         Caption         =   "Tin 1"
         Height          =   285
         Index           =   0
         Left            =   -74895
         TabIndex        =   64
         Top             =   405
         Width           =   1605
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   1860
         Width           =   720
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   5850
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1860
         Width           =   735
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   1875
         Width           =   720
      End
      Begin VB.Line Line27 
         X1              =   -68820
         X2              =   -68820
         Y1              =   345
         Y2              =   1410
      End
      Begin VB.Line Line26 
         X1              =   -69360
         X2              =   -69360
         Y1              =   330
         Y2              =   1410
      End
      Begin VB.Line Line25 
         X1              =   -69870
         X2              =   -69870
         Y1              =   330
         Y2              =   1410
      End
      Begin VB.Line Line24 
         X1              =   -70410
         X2              =   -70410
         Y1              =   330
         Y2              =   1410
      End
      Begin VB.Line Line23 
         X1              =   -70935
         X2              =   -70935
         Y1              =   330
         Y2              =   1410
      End
      Begin VB.Line Line22 
         X1              =   -71445
         X2              =   -71445
         Y1              =   330
         Y2              =   1410
      End
      Begin VB.Line Line21 
         X1              =   -71970
         X2              =   -71970
         Y1              =   330
         Y2              =   1410
      End
      Begin VB.Line Line20 
         X1              =   -72465
         X2              =   -72465
         Y1              =   330
         Y2              =   1410
      End
      Begin VB.Line Line19 
         X1              =   -72960
         X2              =   -72960
         Y1              =   330
         Y2              =   1410
      End
      Begin VB.Line Line18 
         X1              =   -73440
         X2              =   -73440
         Y1              =   330
         Y2              =   1410
      End
      Begin VB.Line Line17 
         X1              =   -73965
         X2              =   -73965
         Y1              =   330
         Y2              =   1410
      End
      Begin VB.Line Line16 
         X1              =   -74445
         X2              =   -74445
         Y1              =   330
         Y2              =   1410
      End
      Begin VB.Line Line15 
         X1              =   -74955
         X2              =   -68265
         Y1              =   1410
         Y2              =   1410
      End
      Begin VB.Line Line14 
         X1              =   -74970
         X2              =   -68250
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Line Line13 
         X1              =   -74955
         X2              =   -68250
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   255
         Index           =   13
         Left            =   -70095
         TabIndex        =   244
         Top             =   1845
         Width           =   510
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   255
         Index           =   12
         Left            =   -72840
         TabIndex        =   243
         Top             =   1815
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Air Pump "
         Height          =   240
         Index           =   3
         Left            =   -71820
         TabIndex        =   242
         Top             =   1845
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Deplating pump"
         Height          =   240
         Index           =   2
         Left            =   -74820
         TabIndex        =   241
         Top             =   1830
         Width           =   1890
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   255
         Index           =   11
         Left            =   -68745
         TabIndex        =   240
         Top             =   1095
         Width           =   285
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   " "
         Height          =   225
         Index           =   11
         Left            =   -68730
         TabIndex        =   239
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   255
         Index           =   10
         Left            =   -69255
         TabIndex        =   238
         Top             =   1095
         Width           =   285
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   " "
         Height          =   225
         Index           =   10
         Left            =   -69240
         TabIndex        =   237
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   255
         Index           =   9
         Left            =   -69765
         TabIndex        =   236
         Top             =   1095
         Width           =   285
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   " "
         Height          =   225
         Index           =   9
         Left            =   -69750
         TabIndex        =   235
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   255
         Index           =   8
         Left            =   -70275
         TabIndex        =   234
         Top             =   1095
         Width           =   285
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   " "
         Height          =   225
         Index           =   8
         Left            =   -70260
         TabIndex        =   233
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   255
         Index           =   7
         Left            =   -70845
         TabIndex        =   232
         Top             =   1095
         Width           =   285
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   " "
         Height          =   225
         Index           =   7
         Left            =   -70830
         TabIndex        =   231
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   255
         Index           =   6
         Left            =   -71355
         TabIndex        =   230
         Top             =   1095
         Width           =   285
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   " "
         Height          =   225
         Index           =   6
         Left            =   -71340
         TabIndex        =   229
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   255
         Index           =   5
         Left            =   -71865
         TabIndex        =   228
         Top             =   1095
         Width           =   285
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   " "
         Height          =   225
         Index           =   5
         Left            =   -71850
         TabIndex        =   227
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   255
         Index           =   4
         Left            =   -72375
         TabIndex        =   226
         Top             =   1095
         Width           =   285
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   " "
         Height          =   225
         Index           =   4
         Left            =   -72360
         TabIndex        =   225
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   255
         Index           =   3
         Left            =   -72840
         TabIndex        =   224
         Top             =   1095
         Width           =   285
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   " "
         Height          =   225
         Index           =   3
         Left            =   -72825
         TabIndex        =   223
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   255
         Index           =   2
         Left            =   -73350
         TabIndex        =   222
         Top             =   1095
         Width           =   285
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   " "
         Height          =   225
         Index           =   2
         Left            =   -73335
         TabIndex        =   221
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   255
         Index           =   1
         Left            =   -73860
         TabIndex        =   220
         Top             =   1095
         Width           =   285
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   " "
         Height          =   225
         Index           =   1
         Left            =   -73845
         TabIndex        =   219
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   255
         Index           =   0
         Left            =   -74370
         TabIndex        =   218
         Top             =   1095
         Width           =   285
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   " "
         Height          =   225
         Index           =   0
         Left            =   -74355
         TabIndex        =   217
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label11 
         Caption         =   "Filter"
         Height          =   255
         Index           =   1
         Left            =   -74970
         TabIndex        =   216
         Top             =   1110
         Width           =   525
      End
      Begin VB.Label Label11 
         Caption         =   "water-level"
         Height          =   255
         Index           =   0
         Left            =   -74955
         TabIndex        =   215
         Top             =   735
         Width           =   525
      End
      Begin VB.Label Label10 
         Caption         =   "Tin1 Tin2 14  17 铜1 銅2 铜3 铜4 铜5 铜6 铜7 铜8"
         Height          =   225
         Left            =   -74400
         TabIndex        =   214
         Top             =   390
         Width           =   6135
      End
      Begin VB.Line Line12 
         Index           =   1
         Visible         =   0   'False
         X1              =   -69225
         X2              =   -69225
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Line11 
         Index           =   1
         X1              =   -70200
         X2              =   -70200
         Y1              =   345
         Y2              =   2175
      End
      Begin VB.Line Line10 
         Index           =   1
         X1              =   -71160
         X2              =   -71160
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Line9 
         Index           =   1
         X1              =   -72120
         X2              =   -72120
         Y1              =   360
         Y2              =   2200
      End
      Begin VB.Line Line8 
         Index           =   1
         X1              =   -73080
         X2              =   -73080
         Y1              =   360
         Y2              =   2145
      End
      Begin VB.Line Line7 
         Index           =   1
         X1              =   -74910
         X2              =   -68295
         Y1              =   1845
         Y2              =   1845
      End
      Begin VB.Line Line6 
         Index           =   1
         X1              =   -74910
         X2              =   -68280
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Line Line5 
         Index           =   1
         X1              =   -74925
         X2              =   -68280
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   -74925
         X2              =   -68265
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   -74940
         X2              =   -68250
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   -74100
         X2              =   -74100
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10000000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   11
         Left            =   -69210
         TabIndex        =   189
         Top             =   1275
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10000000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   10
         Left            =   -70185
         TabIndex        =   188
         Top             =   1275
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10000000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   9
         Left            =   -71145
         TabIndex        =   187
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10000000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   8
         Left            =   -72120
         TabIndex        =   186
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10000000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   7
         Left            =   -73065
         TabIndex        =   185
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10000000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   6
         Left            =   -74040
         TabIndex        =   184
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   11
         Left            =   -69090
         TabIndex        =   183
         Top             =   960
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   10
         Left            =   -70050
         TabIndex        =   182
         Top             =   960
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   9
         Left            =   -71040
         TabIndex        =   181
         Top             =   975
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   8
         Left            =   -71985
         TabIndex        =   180
         Top             =   975
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   7
         Left            =   -72930
         TabIndex        =   179
         Top             =   975
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   6
         Left            =   -73830
         TabIndex        =   178
         Top             =   975
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   11
         Left            =   -69090
         TabIndex        =   177
         Top             =   690
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   10
         Left            =   -70050
         TabIndex        =   176
         Top             =   690
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   9
         Left            =   -71055
         TabIndex        =   175
         Top             =   690
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   8
         Left            =   -71985
         TabIndex        =   174
         Top             =   690
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   7
         Left            =   -72930
         TabIndex        =   173
         Top             =   690
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   6
         Left            =   -73815
         TabIndex        =   172
         Top             =   690
         Width           =   660
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "12"
         Height          =   240
         Index           =   11
         Left            =   -69060
         TabIndex        =   171
         Top             =   1920
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "11"
         Height          =   240
         Index           =   10
         Left            =   -70005
         TabIndex        =   170
         Top             =   1920
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "10"
         Height          =   240
         Index           =   9
         Left            =   -71010
         TabIndex        =   169
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "9"
         Height          =   240
         Index           =   8
         Left            =   -71985
         TabIndex        =   168
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "8"
         Height          =   240
         Index           =   7
         Left            =   -72930
         TabIndex        =   167
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "7"
         Height          =   240
         Index           =   6
         Left            =   -73875
         TabIndex        =   166
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "镀铜10"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   41
         Left            =   -69060
         TabIndex        =   159
         Top             =   360
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "镀铜9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   40
         Left            =   -70080
         TabIndex        =   158
         Top             =   360
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "镀铜8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   39
         Left            =   -71040
         TabIndex        =   157
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "镀铜7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   38
         Left            =   -72000
         TabIndex        =   156
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "镀铜6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   37
         Left            =   -72900
         TabIndex        =   155
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "镀铜5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   36
         Left            =   -73785
         TabIndex        =   154
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "状态:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   35
         Left            =   -74730
         TabIndex        =   153
         Top             =   1905
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "强加:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   34
         Left            =   -74730
         TabIndex        =   152
         Top             =   1590
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "槽号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   33
         Left            =   -74730
         TabIndex        =   151
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "总累加:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   32
         Left            =   -74940
         TabIndex        =   150
         Top             =   1305
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "时长:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   31
         Left            =   -74730
         TabIndex        =   149
         Top             =   975
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "启动值:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   30
         Left            =   -74940
         TabIndex        =   148
         Top             =   675
         Width           =   795
      End
      Begin VB.Line Line12 
         Index           =   0
         X1              =   -69225
         X2              =   -69225
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Line11 
         Index           =   0
         X1              =   -70200
         X2              =   -70200
         Y1              =   345
         Y2              =   2175
      End
      Begin VB.Line Line10 
         Index           =   0
         X1              =   -71160
         X2              =   -71160
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Line9 
         Index           =   0
         X1              =   -72120
         X2              =   -72120
         Y1              =   360
         Y2              =   2200
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   -73080
         X2              =   -73080
         Y1              =   360
         Y2              =   2145
      End
      Begin VB.Line Line7 
         Index           =   0
         X1              =   -74910
         X2              =   -68295
         Y1              =   1845
         Y2              =   1845
      End
      Begin VB.Line Line6 
         Index           =   0
         X1              =   -74910
         X2              =   -68280
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Line Line5 
         Index           =   0
         X1              =   -74925
         X2              =   -68280
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   -74925
         X2              =   -68265
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   -74940
         X2              =   -68250
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   -74100
         X2              =   -74100
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10000000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   5
         Left            =   -69210
         TabIndex        =   147
         Top             =   1275
         Width           =   945
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10000000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   4
         Left            =   -70185
         TabIndex        =   146
         Top             =   1275
         Width           =   945
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10000000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   3
         Left            =   -71145
         TabIndex        =   145
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10000000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   2
         Left            =   -72120
         TabIndex        =   144
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10000000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   1
         Left            =   -73065
         TabIndex        =   143
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10000000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   0
         Left            =   -74040
         TabIndex        =   142
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   5
         Left            =   -69090
         TabIndex        =   141
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   4
         Left            =   -70050
         TabIndex        =   140
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   3
         Left            =   -71040
         TabIndex        =   139
         Top             =   975
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   2
         Left            =   -71985
         TabIndex        =   138
         Top             =   975
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   1
         Left            =   -72930
         TabIndex        =   137
         Top             =   975
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   0
         Left            =   -73830
         TabIndex        =   136
         Top             =   975
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   5
         Left            =   -69090
         TabIndex        =   135
         Top             =   690
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   4
         Left            =   -70050
         TabIndex        =   134
         Top             =   690
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   3
         Left            =   -71055
         TabIndex        =   133
         Top             =   690
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   2
         Left            =   -71985
         TabIndex        =   132
         Top             =   690
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   1
         Left            =   -72930
         TabIndex        =   131
         Top             =   690
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Index           =   0
         Left            =   -73815
         TabIndex        =   130
         Top             =   690
         Width           =   660
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "6"
         Height          =   240
         Index           =   5
         Left            =   -69060
         TabIndex        =   129
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "5"
         Height          =   240
         Index           =   4
         Left            =   -70005
         TabIndex        =   128
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "4"
         Height          =   240
         Index           =   3
         Left            =   -71010
         TabIndex        =   127
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "3"
         Height          =   240
         Index           =   2
         Left            =   -71985
         TabIndex        =   126
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "2"
         Height          =   240
         Index           =   1
         Left            =   -72930
         TabIndex        =   125
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "1"
         Height          =   240
         Index           =   0
         Left            =   -73875
         TabIndex        =   124
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "镀铜4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   29
         Left            =   -69060
         TabIndex        =   117
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "镀铜3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   28
         Left            =   -70080
         TabIndex        =   116
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "镀铜2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   27
         Left            =   -71040
         TabIndex        =   115
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "镀铜1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   26
         Left            =   -72000
         TabIndex        =   114
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "镀锡2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   25
         Left            =   -72900
         TabIndex        =   113
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "镀锡1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   24
         Left            =   -73665
         TabIndex        =   112
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "状态:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   23
         Left            =   -74730
         TabIndex        =   111
         Top             =   1905
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "强加:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   22
         Left            =   -74730
         TabIndex        =   110
         Top             =   1590
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "槽号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   21
         Left            =   -74730
         TabIndex        =   109
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "总累加:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   20
         Left            =   -74940
         TabIndex        =   108
         Top             =   1305
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "时长:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   19
         Left            =   -74730
         TabIndex        =   107
         Top             =   975
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "启动值:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   18
         Left            =   -74940
         TabIndex        =   106
         Top             =   675
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Stop time:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   17
         Left            =   -71715
         TabIndex        =   91
         Top             =   1770
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Run time:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   16
         Left            =   -74610
         TabIndex        =   90
         Top             =   1755
         Width           =   1485
      End
      Begin VB.Label Label5 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   330
         Index           =   1
         Left            =   -70275
         TabIndex        =   89
         Top             =   1755
         Width           =   930
      End
      Begin VB.Label Label5 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   330
         Index           =   0
         Left            =   -73155
         TabIndex        =   88
         Top             =   1740
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Current 4 -5"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   6
         Left            =   300
         TabIndex        =   63
         Top             =   1860
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Current-CU"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   14
         Left            =   4770
         TabIndex        =   62
         Top             =   1875
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Current Tin"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   12
         Left            =   2610
         TabIndex        =   61
         Top             =   1875
         Width           =   1485
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2205
      Left            =   8415
      TabIndex        =   30
      Top             =   15
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   3889
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   16761087
      TabCaption(0)   =   "A- Crame"
      TabPicture(0)   =   "min.frx":0DD0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "B-Crame"
      TabPicture(1)   =   "min.frx":0DEC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2(1)"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1890
         Index           =   1
         Left            =   -74985
         TabIndex        =   43
         Top             =   315
         Width           =   6795
         Begin VB.CommandButton Command11 
            Caption         =   "GO"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   6
            Left            =   4575
            TabIndex        =   52
            Top             =   1125
            Width           =   900
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Index           =   1
            Left            =   5865
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   1300
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Index           =   1
            Left            =   3645
            Locked          =   -1  'True
            TabIndex        =   50
            Tag             =   "7"
            Text            =   "8"
            Top             =   1300
            Width           =   615
         End
         Begin VB.CommandButton Command11 
            Caption         =   "B-Forward"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   4
            Left            =   2345
            TabIndex        =   49
            Top             =   180
            Width           =   900
         End
         Begin VB.CommandButton Command11 
            Caption         =   "B-back"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   5
            Left            =   2345
            TabIndex        =   48
            Top             =   1050
            Width           =   900
         End
         Begin VB.CommandButton Command11 
            Caption         =   "B2Decline"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   3
            Left            =   1225
            TabIndex        =   47
            Top             =   1050
            Width           =   900
         End
         Begin VB.CommandButton Command11 
            Caption         =   "B1Decline"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   2
            Left            =   105
            TabIndex        =   46
            Top             =   1050
            Width           =   900
         End
         Begin VB.CommandButton Command11 
            Caption         =   "B2-UP"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   1225
            TabIndex        =   45
            Top             =   180
            Width           =   900
         End
         Begin VB.CommandButton Command11 
            Caption         =   "B1-up"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   105
            TabIndex        =   44
            Top             =   180
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "P/N A:"
            Height          =   240
            Index           =   43
            Left            =   3420
            TabIndex        =   207
            Top             =   105
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "P/N B:"
            Height          =   240
            Index           =   42
            Left            =   3420
            TabIndex        =   206
            Top             =   540
            Width           =   810
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ABCDEFGABCDEFGAB"
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   4
            Left            =   4290
            TabIndex        =   205
            Top             =   60
            Width           =   2265
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   345
            Index           =   5
            Left            =   4290
            TabIndex        =   204
            Top             =   510
            Width           =   2250
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   3
            Left            =   2640
            Picture         =   "min.frx":0E08
            Top             =   240
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Image Image1 
            Height          =   750
            Index           =   4
            Left            =   1320
            Picture         =   "min.frx":12BD
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Image Image1 
            Height          =   750
            Index           =   5
            Left            =   0
            Picture         =   "min.frx":173F
            Top             =   0
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Target"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   5535
            TabIndex        =   54
            Top             =   945
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   3315
            TabIndex        =   53
            Top             =   945
            Width           =   1155
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   1980
         Index           =   0
         Left            =   15
         TabIndex        =   31
         Top             =   315
         Width           =   7290
         Begin VB.CommandButton Command1 
            Caption         =   "A1-up"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   105
            TabIndex        =   40
            Top             =   180
            Width           =   900
         End
         Begin VB.CommandButton Command1 
            Caption         =   "A2-up"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   1230
            TabIndex        =   39
            Top             =   180
            Width           =   900
         End
         Begin VB.CommandButton Command1 
            Caption         =   "A1-down"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   2
            Left            =   105
            TabIndex        =   38
            Top             =   1050
            Width           =   900
         End
         Begin VB.CommandButton Command1 
            Caption         =   "A2-down"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   3
            Left            =   1230
            TabIndex        =   37
            Top             =   1050
            Width           =   900
         End
         Begin VB.CommandButton Command1 
            Caption         =   "A- forward"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   4
            Left            =   2345
            TabIndex        =   36
            Top             =   180
            Width           =   900
         End
         Begin VB.CommandButton Command1 
            Caption         =   "A-back"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   5
            Left            =   2345
            TabIndex        =   35
            Top             =   1050
            Width           =   900
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Index           =   0
            Left            =   3645
            Locked          =   -1  'True
            TabIndex        =   34
            Tag             =   "7"
            Text            =   "8"
            Top             =   1300
            Width           =   615
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Index           =   0
            Left            =   5865
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1300
            Width           =   615
         End
         Begin VB.CommandButton Command1 
            Caption         =   "GO"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   6
            Left            =   4575
            TabIndex        =   32
            Top             =   1125
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Item A:"
            Height          =   240
            Index           =   45
            Left            =   3420
            TabIndex        =   211
            Top             =   105
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Item B:"
            Height          =   240
            Index           =   44
            Left            =   3420
            TabIndex        =   210
            Top             =   540
            Width           =   945
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ABCDEFGABCDEFGAB"
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   2
            Left            =   4290
            TabIndex        =   209
            Top             =   60
            Width           =   2265
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   345
            Index           =   3
            Left            =   4290
            TabIndex        =   208
            Top             =   510
            Width           =   2250
         End
         Begin VB.Image Image1 
            Height          =   750
            Index           =   0
            Left            =   0
            Picture         =   "min.frx":1C28
            Top             =   0
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Image Image1 
            Height          =   750
            Index           =   1
            Left            =   1320
            Picture         =   "min.frx":2111
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   2
            Left            =   2640
            Picture         =   "min.frx":2593
            Top             =   240
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "current position"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   0
            Left            =   3315
            TabIndex        =   42
            Top             =   945
            Width           =   1800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Goal position"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   5535
            TabIndex        =   41
            Top             =   945
            Width           =   2145
         End
      End
   End
   Begin VB.Timer Timer7 
      Interval        =   200
      Left            =   10920
      Top             =   0
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   220
      Left            =   11280
      Top             =   0
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   11640
      Top             =   0
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   12000
      Top             =   0
   End
   Begin VB.Timer Timer11 
      Interval        =   15000
      Left            =   12360
      Top             =   0
   End
   Begin VB.Timer Timer12 
      Interval        =   1000
      Left            =   12600
      Top             =   0
   End
   Begin VB.Timer Timer13 
      Interval        =   8000
      Left            =   12960
      Top             =   0
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   13320
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   2535
      Left            =   0
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "生产信息，当前所用的生产料号。"
      Top             =   7920
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4471
      _Version        =   393216
      BackColorSel    =   -2147483636
      BackColorBkg    =   14737632
      GridColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer6 
      Interval        =   430
      Left            =   10560
      Top             =   0
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Item"
      Height          =   4365
      Left            =   8385
      TabIndex        =   2
      ToolTipText     =   "料号投入口"
      Top             =   6075
      Width           =   6855
      Begin VB.ListBox List3 
         BackColor       =   &H00FFC0C0&
         Height          =   2460
         ItemData        =   "min.frx":2A48
         Left            =   90
         List            =   "min.frx":2ABB
         TabIndex        =   56
         Top             =   1800
         Width           =   6720
      End
      Begin VB.CommandButton Command2 
         Height          =   300
         Left            =   6375
         Picture         =   "min.frx":2B31
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   1550
         Width           =   330
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   360
         Index           =   1
         Left            =   5745
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   360
         Index           =   0
         Left            =   5745
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1845
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   -15
         X2              =   6825
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   345
         Index           =   1
         Left            =   1545
         TabIndex        =   25
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   0
         Left            =   1545
         TabIndex        =   24
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         Caption         =   "Fault display:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   11
         Left            =   75
         TabIndex        =   19
         Top             =   1530
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         Caption         =   "Run time:"
         Height          =   240
         Index           =   10
         Left            =   4545
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         Caption         =   "periodic time:"
         Height          =   240
         Index           =   9
         Left            =   4545
         TabIndex        =   15
         Top             =   360
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         Caption         =   "Upload B:"
         Height          =   240
         Index           =   8
         Left            =   105
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         Caption         =   "upload A:"
         Height          =   240
         Index           =   7
         Left            =   105
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "System control"
      Height          =   1605
      Left            =   8415
      TabIndex        =   1
      ToolTipText     =   "系统控制,请谨慎操作,注意安全!"
      Top             =   2235
      Width           =   6795
      Begin VB.CommandButton Command3 
         Caption         =   "AG ON"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   213
         Top             =   195
         Width           =   1545
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Once"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   212
         Top             =   195
         Width           =   1545
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   5040
         TabIndex        =   12
         Top             =   1140
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   645
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Maintain"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   1800
         Picture         =   "min.frx":2EE0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   465
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Manual"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   960
         Picture         =   "min.frx":37AA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   465
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Auto"
         DownPicture     =   "min.frx":4074
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   120
         Picture         =   "min.frx":493E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   465
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000001&
         Height          =   300
         Left            =   2880
         Top             =   1080
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C-state"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   15
         Left            =   2805
         TabIndex        =   27
         Top             =   900
         Width           =   840
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   3000
         TabIndex        =   26
         Top             =   1200
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Current state:"
         Height          =   240
         Index           =   5
         Left            =   3120
         TabIndex        =   10
         Top             =   1200
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Current positon:"
         Height          =   240
         Index           =   4
         Left            =   2880
         TabIndex        =   9
         Top             =   600
         Width           =   2160
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   600
      Top             =   8160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      RecordSource    =   "select * from 料号库"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   9480
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   9840
      Top             =   0
   End
   Begin VB.Timer Timer5 
      Interval        =   120
      Left            =   10200
      Top             =   0
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8520
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      ParitySetting   =   2
      DataBits        =   7
      StopBits        =   2
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   8760
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   9120
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   10440
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16184
            MinWidth        =   4410
            Picture         =   "min.frx":5208
            Text            =   " 中新电镀设备制造有限公司承制 ---- 版本 V2.3  "
            TextSave        =   " 中新电镀设备制造有限公司承制 ---- 版本 V2.3  "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
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
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      ItemData        =   "min.frx":62E9
      Left            =   120
      List            =   "min.frx":6311
      TabIndex        =   23
      Top             =   8160
      Width           =   285
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   13320
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "故障表"
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   8175
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   14420
         _Version        =   393216
         Rows            =   39
         Cols            =   11
         FixedCols       =   0
         RowHeightMin    =   14
         GridColor       =   16744576
         FocusRect       =   2
         FillStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5595
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "min.frx":636B
         Top             =   960
         Width           =   6615
      End
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   330
         Left            =   6600
         Top             =   120
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
         RecordSource    =   "select * from 电镀记录"
         Caption         =   "Adodc4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5520
         ItemData        =   "min.frx":6373
         Left            =   6480
         List            =   "min.frx":63F2
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   6600
         TabIndex        =   29
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      RecordSource    =   "select * from 监控画面"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "下次退镀:"
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
      Index           =   13
      Left            =   0
      TabIndex        =   21
      Top             =   40
      Width           =   945
   End
   Begin VB.Menu M_w01 
      Caption         =   "Filer(&F)"
      Begin VB.Menu M_W04 
         Caption         =   "operator record "
      End
      Begin VB.Menu M_w02 
         Caption         =   "Fault record"
      End
      Begin VB.Menu M_w03 
         Caption         =   "production record"
      End
      Begin VB.Menu M_x04 
         Caption         =   "Exit system"
      End
   End
   Begin VB.Menu M_p01 
      Caption         =   "Edit(&E)"
      Begin VB.Menu M_p02 
         Caption         =   "P/N Edit"
      End
      Begin VB.Menu M_p03 
         Caption         =   "P/N update"
      End
   End
   Begin VB.Menu M_c01 
      Caption         =   "Inquire(&C)"
      Visible         =   0   'False
      Begin VB.Menu M_cs01 
         Caption         =   "Production record query"
         Begin VB.Menu M_cs05 
            Caption         =   "Query by tank"
         End
         Begin VB.Menu M_cs04 
            Caption         =   "Query by area"
         End
         Begin VB.Menu M_cs03 
            Caption         =   "Query by P/N"
         End
         Begin VB.Menu M_cs02 
            Caption         =   "Query by time"
         End
      End
      Begin VB.Menu M_cl01 
         Caption         =   "P/N query"
         Begin VB.Menu M_cl02 
            Caption         =   "P/N query"
         End
         Begin VB.Menu M_cl03 
            Caption         =   "Query by Ampere density "
         End
      End
      Begin VB.Menu M_cq01 
         Caption         =   "Machine fault query"
      End
   End
   Begin VB.Menu M_l01 
      Caption         =   "Process information&L)"
   End
   Begin VB.Menu M_x01 
      Caption         =   "System management(&A)"
      Begin VB.Menu M_09 
         Caption         =   "change User"
      End
      Begin VB.Menu M_06 
         Caption         =   "User management"
      End
      Begin VB.Menu M_x02 
         Caption         =   "Add user"
      End
      Begin VB.Menu M_x03 
         Caption         =   "Modification urser"
      End
   End
   Begin VB.Menu M_UNLOAD 
      Caption         =   "Refresh(&U)"
      Visible         =   0   'False
   End
   Begin VB.Menu M_b01 
      Caption         =   "Help(&H)"
      Begin VB.Menu M_g01 
         Caption         =   "About zy(G)"
      End
   End
   Begin VB.Menu M_S01 
      Caption         =   "PN-Operaton(&G)"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu M_S02 
         Caption         =   "PN-operation(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu ly 
         Caption         =   "A/B Flybar change(&Q)"
         Shortcut        =   ^Q
      End
      Begin VB.Menu M_03 
         Caption         =   "PN Clean (&D)"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu m_h01 
      Caption         =   "h"
      Visible         =   0   'False
      Begin VB.Menu m_h03 
         Caption         =   "Feed Item"
         Shortcut        =   ^H
      End
      Begin VB.Menu M_H02 
         Caption         =   "Delete Item"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "取消(&U)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "剪切(&T)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "复制(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "粘贴(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "选择性粘贴(&S)..."
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDSelectAll 
         Caption         =   "全部选定(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditInvertSelection 
         Caption         =   "反向选定(&I)"
      End
   End
End
Attribute VB_Name = "min"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim eros As String

Private Sub mnuEditCopy_Click()
  MsgBox "将复制的代码放在此处！"
End Sub

Private Sub mnuEditCut_Click()
  MsgBox "将剪切的代码放在此处！"
End Sub

Private Sub mnuEditDSelectAll_Click()
  MsgBox "将全部选择的代码放在此处！"
End Sub

Private Sub mnuEditInvertSelection_Click()
  MsgBox "将反向选择的代码放在此处！"
End Sub

Private Sub mnuEditPaste_Click()
  MsgBox "将粘贴的代码放在此处！"
End Sub

Private Sub mnuEditPasteSpecial_Click()
  MsgBox "将粘贴的特殊代码放在此处！"
End Sub

Private Sub mnuEditUndo_Click()
  MsgBox "将 Undo代码放在此处！"
End Sub

Private Sub Check1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call PASS
 If A_P <> userpass Then Exit Sub
Dim a As String
    a = Trim(Str(600 + Index))
    If Check1(Index).Value = 0 Then
        o(3) = "@00KSHR  000" & a
    Else
        o(3) = "@00KRHR  000" & a
    End If
 o(4) = "@00KC"
End Sub

Private Sub Check2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As String
    a = Trim(Str(400 + Index))
    If Check2(Index).Value = 0 Then
        o(3) = "@00KSHR  000" & a
    Else
        o(3) = "@00KRHR  000" & a
    End If
    o(4) = "@00KC"
End Sub

Private Sub Check3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As String
    a = Trim(Str(8000 + Index))
    If Check3(Index).Value = 0 Then
        o(3) = "@00KSHR  00" & a
    Else
        o(3) = "@00KRHR  00" & a
    End If
    o(4) = "@00KC"
End Sub

Private Sub Check4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As String
    a = Trim(Str(20600 + Index))
    If Check4(Index).Value = 0 Then
        o(3) = "@00KSCIO 0" & a
    Else
        o(3) = "@00KRCIO 0" & a
    End If
    o(4) = "@00KC"
End Sub


Private Sub Check5_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call PASS
 If A_P <> userpass Then Exit Sub
Dim a As String
    a = Trim(Str(500 + Index))
    If Check5(Index).Value = 0 Then
        o(3) = "@00KSHR  000" & a
    Else
        o(3) = "@00KRHR  000" & a
    End If
    o(4) = "@00KC"
End Sub

Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Select Case Index
       Case 0
         o(3) = "@00KSCIO 000410"
       Case 1
         o(3) = "@00KSCIO 000514"
       Case 2
         o(3) = "@00KSCIO 000411"
       Case 3
         o(3) = "@00KSCIO 000515"
       Case 4
         o(3) = "@00KSCIO 000412"
       Case 5
         o(3) = "@00KSCIO 000413"
       Case 6
        o(3) = "@00KSCIO 002909"
        o(4) = "@00KRCIO 002909"
   End Select
End Sub
Private Sub Command1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If o(3) = "" Then o(3) = "@00KC" Else o(4) = "@00KC"
End Sub

Private Sub Command11_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Exit Sub
   Select Case Index
       Case 0
        o(3) = "@00KSCIO 000210"
       Case 1
        o(3) = "@00KSCIO 000312"
       Case 2
        o(3) = "@00KSCIO 000211"
       Case 3
        o(3) = "@00KSCIO 000313"
       Case 4
        o(3) = "@00KSCIO 000212"
       Case 5
        o(3) = "@00KSCIO 000213"
       Case 6
        o(3) = "@00KSCIO 002914"
        o(4) = "@00KRCIO 002914"
    End Select
End Sub

Private Sub Command11_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If o(3) = "" Then o(3) = "@00KC" Else o(4) = "@00KC"
End Sub

Private Sub Command2_Click()
eros = ""
min.Text10.Text = ""
List3.Clear
End Sub

Private Sub Command3_Click(Index As Integer)
Select Case Index
Case 0
  o(3) = "@00KSCIO 000000"
  o(4) = "@00KRCIO 000000"
  o(5) = "@00KC"
Case 1
  o(3) = "@00KRCIO 000001"
  o(4) = "@00KSCIO 000001"
  o(5) = "@00KRHR  000514"
  o(6) = "@00KC"
Case 2
  o(3) = "@00KSHR  000514"
  o(4) = "@00KC"
Case 3
  If Command3(Index).Caption <> "Cont" Then o(3) = "@00KSHR  000414" Else o(3) = "@00KRHR  000414"
  o(4) = "@00KC"
Case 4
  If Command3(Index).Caption <> "AG Open" Then o(3) = "@00KSHR  000512" Else o(3) = "@00KRHR  000512"
  o(4) = "@00KC"
End Select
End Sub

Private Sub Form_DblClick()
 Text5.Visible = True
End Sub

Private Sub Form_Load()
Dim rs1 As New Recordset
 Call REGEDIT
 StatusBar1.Panels.Item(2) = Chr(9) & Chr(9) & "Current operator:" & Chr(9) & userID
  Call T_in
 For i = 0 To 10
 If i <> 4 Then MSFlexGrid1.ColAlignment(i) = 4 Else MSFlexGrid1.ColAlignment(i) = 1
 Next
 List1.Clear
 sq1 = "select * from 监控画面"
 rs1.Open sq1, conn, adOpenKeyset, adLockPessimistic
  vv = rs1.RecordCount
  Do While Not rs1.EOF
      List1.AddItem rs1.Fields(1) ' & rs1.Fields(1)
      rs1.MoveNext
      'vv = vv + 1
  Loop '
  rs1.Close
 MSFlexGrid1.RowHeight(0) = 400
 Me.MSFlexGrid2.RowHeight(0) = 350
 List3.Clear
   For i = 1 To vv
    min.MSFlexGrid1.Col = 1
    min.MSFlexGrid1.Row = i
    min.MSFlexGrid1.CellBackColor = &H4488FF
    min.MSFlexGrid1.Col = 2
    min.MSFlexGrid1.Row = i
    min.MSFlexGrid1.CellBackColor = &H44FF00
    min.MSFlexGrid1.Col = 5
    min.MSFlexGrid1.Row = i
    min.MSFlexGrid1.CellBackColor = &HFFFF00
    min.MSFlexGrid1.Col = 6
    min.MSFlexGrid1.Row = i
    min.MSFlexGrid1.CellBackColor = &HFF80FF
    min.MSFlexGrid1.Col = 7
    min.MSFlexGrid1.Row = i
    min.MSFlexGrid1.CellBackColor = &H44FFFF
    min.MSFlexGrid1.Col = 8
    min.MSFlexGrid1.Row = i
    min.MSFlexGrid1.CellBackColor = &H44FF88
   Next

 For i = 1 To vv
  Me.MSFlexGrid1.RowHeight(i) = 300
 Next
 'Flash1.Movie = App.Path & "\mrkj.swf"
 MSFlexGrid1.ColWidth(0) = GetSetting("zxdms", "settings", "colw0", "1")
 MSFlexGrid1.ColWidth(1) = GetSetting("zxdms", "settings", "colw1", "1200")
 MSFlexGrid1.ColWidth(2) = GetSetting("zxdms", "settings", "colw2", "2000")
 MSFlexGrid1.ColWidth(3) = GetSetting("zxdms", "settings", "colw3", "1000")
 MSFlexGrid1.ColWidth(4) = GetSetting("zxdms", "settings", "colw4", "2320")
 MSFlexGrid1.ColWidth(5) = GetSetting("zxdms", "settings", "colw5", "1000")
 MSFlexGrid1.ColWidth(6) = GetSetting("zxdms", "settings", "colw6", "830")
 MSFlexGrid1.ColWidth(7) = GetSetting("zxdms", "settings", "colw7", "1010")
 MSFlexGrid1.ColWidth(8) = GetSetting("zxdms", "settings", "colw8", "1010")
 MSFlexGrid1.ColWidth(9) = GetSetting("zxdms", "settings", "colw9", "1")
 MSFlexGrid1.ColWidth(10) = GetSetting("zxdms", "settings", "colw10", "1")
 MSFlexGrid1.TextMatrix(0, 1) = "Tank /N"
 MSFlexGrid1.TextMatrix(0, 2) = "process"
 MSFlexGrid1.TextMatrix(0, 3) = "crane "
 'MSFlexGrid1.TextMatrix(0, 4) = "生产料号--(A飞靶)"
 MSFlexGrid1.TextMatrix(0, 5) = "QTY"
 MSFlexGrid1.TextMatrix(0, 6) = "Time"
 MSFlexGrid1.TextMatrix(0, 7) = "A -A"
 MSFlexGrid1.TextMatrix(0, 8) = "B- A"
 For i = 1 To vv
 MSFlexGrid1.Col = 3
 MSFlexGrid1.Row = i
  min.MSFlexGrid1.CellFontBold = False
  min.MSFlexGrid1.CellFontSize = 16
  min.MSFlexGrid1.CellForeColor = &HFF&
 Next

 For i = 1 To vv
  MSFlexGrid1.TextMatrix(i, 1) = i
 Next i
 For i = 0 To vv - 1
  MSFlexGrid1.TextMatrix(i + 1, 2) = List1.List(i)
 Next

min.MSComm1.OutBufferSize = 1024
 min.MSComm1.InBufferSize = 1024
If Not min.MSComm1.PortOpen Then
    min.MSComm1.CommPort = intport
    min.MSComm1.Settings = setting
    min.MSComm1.PortOpen = True
End If
min.MSComm1.InputLen = 0
min.MSComm1.InputMode = comInputModeText
'min.MSComm1.RThreshold = 1
min.MSComm1.Output = "@00XZ42*" & o(0)
min.Text5 = "load complete"
min.Timer3.Enabled = True
min.Timer4.Enabled = True
min.Timer5.Enabled = True
min.Timer6.Enabled = True
Call lao_xin
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Select Case Index
 Case 0
    Add_liao.Label2.Caption = "Feed Item ："
    Add_liao.Label3.Caption = "A Flybar"
    pindex = 4050
    Label2(0).BackColor = &H80C0FF
    MA = 0
 Case 1
    Add_liao.Label2.Caption = "Feed Item："
    Add_liao.Label3.Caption = "B Flybar"
    pindex = 4100
    Label2(1).BackColor = &H80C0FF
    MA = 0
 Case 2
    Add_liao.Label2.Caption = "Feed Ttem："
    Add_liao.Label3.Caption = "A_A"
    pindex = 4109
    Label2(2).BackColor = &H80C0FF
    MA = 0
 Case 3
    Add_liao.Label2.Caption = "Feed Item："
    Add_liao.Label3.Caption = "A_B"
    pindex = 4110
    Label2(3).BackColor = &H80C0FF
    MA = 0
 Case 4
    Exit Sub
    Add_liao.Label2.Caption = "Feed Item："
    Add_liao.Label3.Caption = "B_A"
    pindex = 4111
    Label2(4).BackColor = &H80C0FF
    MA = 0
 Case 5
    Exit Sub
    Add_liao.Label2.Caption = "Feed Item ："
    Add_liao.Label3.Caption = "B_B"
    pindex = 4112
    Label2(5).BackColor = &H80C0FF
    MA = 0
 End Select
If Button = 2 Then
  PopupMenu m_h01
End If
Label2(Index).BackColor = &HFFFFFF
End Sub

Private Sub Label5_Click(Index As Integer)
Dim a As String
a = InputBox("VBT time:", "SET", "15")
If a = "" Or Val(a) > 40 Then Exit Sub
If Len(a) = 1 Then a = "0" + a
b = InputBox("STOP Time:", "SET", "20")
If b = "" Or Val(b) > 60 Then Exit Sub
If Len(b) = 1 Then b = "0" + b
o(3) = "@00WD412800" + a + "00" + b
End Sub

Private Sub Label7_Click(Index As Integer)
Dim a As String
    a = InputBox("Start time(A.h):", "AUT dosing pump")
    If Len(a) > 4 Or a = "" Then MsgBox "error!", , "err": Exit Sub
    a = Right("000" & a, 4)
   If Index < 12 Then o(3) = "@00WD" & Trim(Str(4180 + Index)) & a
End Sub

Private Sub Label8_Click(Index As Integer)
  Dim a As String
 a = InputBox("Dosing pump time(s):", "Dosing pump")
    If Len(a) > 4 Or a = "" Then MsgBox "error!", , "err": Exit Sub
    a = Right("000" & a, 4)
    If Index < 12 Then o(3) = "@00WD" & Trim(Str(4190 + Index)) & a
End Sub

Private Sub M_03_Click()
o(3) = "@00WD" & Trim(pindex) & "0000"
Call Add_liao.D_LIAO_IN1
End Sub
Private Sub M_06_Click()
If userpow <> "SYS Arrange" Then
      MsgBox "PLS login Adim modle!", 48, "Marked"
      Exit Sub
Else
 DEL_user.Show
End If
End Sub

Private Sub M_09_Click()
ChUser.Show 1
End Sub

Private Sub M_l01_Click()
 LU_lock.Show
End Sub

Private Sub M_lock_Click()
 If TIN_01 = True Then TIN_01 = False Else TIN_01 = True
 Call AfroB
End Sub
Private Sub M_p02_Click()
LAO_eidt.Show
End Sub

Private Sub M_p03_Click()
 LAO_eidt.Show
 LAO_eidt.Caption = "Item Change"
 LAO_eidt.Height = 6510
 LAO_eidt.Timer1 = True
End Sub
Private Sub m_h02_click()
  o(3) = "@00WD" & Trim(pindex) & "0000"
   Call Add_liao.D_LIAO_IN1
   min.Label2(0).BackColor = &HFFFFFF
   min.Label2(1).BackColor = &HFFFFFF
End Sub

Private Sub m_h03_click()
 Add_liao.Show
 Add_liao.Top = 6800
 Add_liao.Left = 5250
End Sub

Private Sub M_UNLOAD_Click()
If Me.MSComm1.PortOpen = True Then Me.MSComm1.PortOpen = False
  qsr0 = False
  Me.MSComm1.InBufferCount = 0
  Me.MSComm1.OutBufferCount = 0
  Form_Load
End Sub

Private Sub M_w02_Click()
 JI_LOCK.Label3.Caption = "您可以输入相应的时间来查询故障显示"
 JI_LOCK.Caption = "故障记录"
 JI_LOCK.Label1(1).Enabled = False
 JI_LOCK.Combo1.Enabled = False
 JI_LOCK.Check3.Enabled = False
 JI_LOCK.Adodc1.RecordSource = "select * from 故障表"
 JI_LOCK.Adodc1.Refresh
 JI_LOCK.Label2.Caption = "共有" & JI_LOCK.Adodc1.Recordset.RecordCount & "条记录"
 JI_LOCK.Show
 'JI_LOCK.Command3.Visible = False
End Sub

Private Sub M_w03_Click()
 JI_LOCK.Width = 10500
 JI_LOCK.DataGrid1.Width = 10360
 JI_LOCK.Caption = "生产记录"
 JI_LOCK.Adodc1.RecordSource = "select * from 电镀记录"
 JI_LOCK.Adodc1.Refresh
 JI_LOCK.Label2.Caption = "共有" & JI_LOCK.Adodc1.Recordset.RecordCount & "条记录"
 JI_LOCK.Show
 Call d_lock
' JI_LOCK.Command3.Visible = True
  'JI_LOCK.Adodc1.Recordset.Close
  If bbbb = True Then JI_LOCK.Command5.Enabled = True: JI_LOCK.Command1.Enabled = False
End Sub

Private Sub M_W04_Click()
 T_lock.Show
End Sub

Private Sub MSFlexGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Add_liao.Visible = True Then
 Exit Sub
 Else
 m_t = Y
 m_l = MSFlexGrid1.ColWidth(0) + MSFlexGrid1.ColWidth(1) + MSFlexGrid1.ColWidth(2) + MSFlexGrid1.ColWidth(3) + _
  MSFlexGrid1.ColWidth(4)
 
 Add_liao.Top = m_l
 Add_liao.Left = 5250 'm_t
 End If
End Sub
Private Sub MSFlexGrid2_DblClick()
Dim a_unmer As Variant
Dim a_input As String
 a_instr = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.RowSel, 0)
 a_unmer = Me.MSFlexGrid2.TextMatrix(Me.MSFlexGrid2.RowSel, 11)
Me.Timer6 = False
a_vbYesNo = MsgBox("是否要删除生产信息" & a_instr, vbYesNo, "提示")
If a_vbYesNo = 6 Then
    For m = 0 To vv
        If Val(a_unmer) = Val(Right(Ncode(m), 2)) Then MsgBox "料号正在使用，不能删除！", 46, "提示": Me.Timer6 = True: Exit Sub
    Next
    For i = 0 To 22
    If Val(a_unmer) = Val(Right(Ncode(m + 50), 2)) Then MsgBox "料号正在使用，不能删除！", 46, "提示": Me.Timer6 = True: Exit Sub
     Next i
   ' For i = 12 To 21
   ' If Val(a_unmer) = Val(Right(Ncode(m + 50), 2)) Then MsgBox "料号正在使用，不能删除！", 46, "提示": Me.Timer6 = True: Exit Sub
   '  Next i
 a_input = InputBox("请输入密码!", "密码")
 If a_input = userpass Then
    Me.Adodc2.Refresh
    If Me.Adodc2.Recordset.RecordCount > 1 Then
       While Me.Adodc2.Recordset.EOF = False
            If Me.Adodc2.Recordset.Fields(0) = a_instr Then
               Me.Adodc2.Recordset.Delete
               Me.Adodc2.Recordset.Update
               Call lao_xin
                Call D_ING
              'Me.Adodc2.Recordset.Close
            Else
              Me.Adodc2.Recordset.MoveNext
            End If
       Wend
   Else
        MsgBox "生产料号列表不能为空！", 48, "提示"
    End If
 Else
  MsgBox "密码错误", 16, "密码"
 End If
 Else
 Exit Sub
 End If
 Me.Timer6 = True
 'Me.Adodc2.Recordset.Close
End Sub
Private Sub MSFlexGrid2_LostFocus()
Me.Timer6 = True
End Sub

Private Sub SSTab1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If SSTab1.Tab <> 0 Then SSTab1.Tab = 0
End Sub

Private Sub Text1_DblClick(Index As Integer)
 MsgBox "crane location is obsolute value！" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "If the value has wrong pls checkproximity sweetch ！", , "REMINDR"
 Exit Sub
End Sub

Private Sub Text11_DblClick(Index As Integer)
 Call PASS
 If A_P <> userpass Then Exit Sub
If Index = 1 Then
    Dim a As String
        a = InputBox("Input Tin tank NO (1～2):", "Parity tank oprator")
        If a < 3 And a > 0 Then
            o(3) = "@00WD01100" & CStr(110 + Val(a))
        Else
            MsgBox "Input range is errow!", , "parity tank oprator"
        End If
End If
If Index = 2 Then
        a = InputBox("Input coppper tank NO (1～8):", "parity tank oprator")
        If a < 9 And a > 0 Then
            o(3) = "@00WD01200" & CStr(120 + Val(a))
        Else
            MsgBox "Input range is errow!", , "parity tank oprator"
        End If
End If
End Sub

Private Sub Text2_DblClick(Index As Integer)
 Dim DATA1 As String
 Select Case Index
Case 0
    DATA1 = CStr(Val(InputBox("A Crane goal value (0~ B):", "change ")))
    If Val(DATA1) > 0 And Val(DATA1) < Text1(1).Text - 2 Then
        If Len(DATA1) = 1 Then DATA1 = "0" + DATA1
        If o(3) = "" Then o(3) = "@00WD030600" + CStr(DATA1) Else o(4) = "@00WD030600" + CStr(DATA1)
    Else
        MsgBox "Input range is wrong!", , "error"
    End If
Case 1
    Exit Sub
    DATA1 = CStr(Val(InputBox("B Crane goal value(A ~ 41):", "change")))
    If Val(DATA1) > Text1(0).Text + 2 And Val(DATA1) <= 41 Then
        If Len(DATA1) = 1 Then DATA1 = "0" + DATA1
        If o(3) = "" Then o(3) = "@00WD032600" + CStr(DATA1) Else o(4) = "@00WD032600" + CStr(DATA1)
    Else
        MsgBox "Input range is wrong!", , "errow"
    End If
End Select
End Sub

Private Sub Text10_Change()
Dim lis3(40) As String
Dim lis3c As Integer
If Text10.Text <> "Failure-free！" Then
    If List3.ListCount > 40 Then lis3c = 40 Else lis3c = List3.ListCount
    For i = 0 To lis3c
    lis3(i) = List3.List(i)
    Next
    List3.Clear
    List3.AddItem Trim(Now) & Chr(9) & Trim(Text10.Text) & Chr(9) & userID  '& Chr(9) & Trim(Text7.Text)
    For i = 1 To lis3c + 1
    List3.AddItem lis3(i - 1)
    Next
    On Error GoTo TEXT10_ERROR
    min.Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
       min.Adodc3.Recordset.MoveLast
          If Trim(min.Adodc3.Recordset.Fields(0)) = Trim(Now) Then
          Exit Sub
          Else
          Call TEXT10_SB
          End If
     Else
        Call TEXT10_SB
    End If
 End If
 min.Adodc3.Refresh
TEXT10_ERROR:
End Sub

Sub TEXT10_SB()
On Error GoTo TEXT10_ERROR
If errec <> Trim(min.Text10.Text) Then
 errec = Trim(min.Text10.Text)
 min.Adodc3.Refresh
 min.Adodc3.Recordset.AddNew
 min.Adodc3.Recordset.Fields(0) = Trim(Now)
 min.Adodc3.Recordset.Fields(1) = Trim(min.Text10.Text)
 min.Adodc3.Recordset.Fields(2) = estr
 min.Adodc3.Recordset.Fields(3) = userID
 min.Adodc3.Recordset.Update
 min.Adodc3.Refresh
 End If
TEXT10_ERROR:
  Exit Sub
End Sub
Private Sub Text6_DblClick()
 Dim DATA1 As String
 Dim AKEY As Integer
    If Text6.Enabled = True Then
        DATA1 = InputBox("Please enter the process number to use(1~5):", "Process")
        If DATA1 = "" Then Exit Sub
        If DATA1 = "0" Or DATA1 > 5 Then GoTo eee
        AKEY = MsgBox("You do want to change the process to:" + DATA1 + "? ", vbOKCancel, "Process")
        If AKEY = vbCancel Then Exit Sub
        o(3) = "@00WD0068000" + Trim(Str(Val(DATA1)))
        MsgBox "At the same time, press the 'Manual/Stop' button and the 'Reset' button on the panel of the power distribution cabinet to enable the new process!", , "attention"
    End If
Exit Sub
eee:
MsgBox "Input value is wrong!", , "wrong"
End Sub

Private Sub Timer7_Timer()
If min.Label3.Visible = False Then min.Timer9.Enabled = True Else min.Timer9.Enabled = False
End Sub

Private Sub M_S02_Click()
 Add_liao.Show
  Add_liao.Top = Val(m_t)
  Add_liao.Left = 5250 'Val(m_l)
End Sub

Private Sub M_x02_Click()
 Add_user.Show
End Sub

Private Sub M_x03_Click()
 MOD_PAWWD.Show
End Sub

Private Sub M_x04_Click()
 End
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cc As Integer
cc = 0
For i = 0 To 9
  cc = cc + MSFlexGrid1.ColWidth(i)
  If X <= cc Then MSFlexGrid1.Col = i: Exit For
Next
MA = MSFlexGrid1.TopRow + ((Y - min.MSFlexGrid1.RowHeight(0)) \ min.MSFlexGrid1.RowHeight(1))
Call input_index
If MSFlexGrid1.Col = 4 And Button = 1 Then
  'Call input_index
  Add_liao.Timer1.Enabled = True
End If
If Button = 2 Then
  If MSFlexGrid1.Col = 4 Then PopupMenu M_S01 Else Exit Sub  'M_S01料号操作菜单
  'MA = MSFlexGrid1.TopRow + ((Y - min.MSFlexGrid1.RowHeight(0)) \ min.MSFlexGrid1.RowHeight(1)) 'MB = MSHFlexGrid1.RowSel
  'Call input_index
  AfroB
  Add_liao.Timer1.Enabled = True
End If
End Sub
Private Sub Text5_DblClick()
 Text5.Visible = False
End Sub

Private Sub Timer1_Timer()
 StatusBar1.Panels.Item(3) = Chr(9) & Chr(9) & Now
End Sub

Private Sub MSComm1_OnComm()
Dim et As Integer
Dim st As Integer
Dim sts As Integer
Dim lt As Integer
min.Label3.Visible = True
min.Timer8.Enabled = True
Select Case min.MSComm1.CommEvent
    Case 2
        If Not min.MSComm1.PortOpen Then
            min.MSComm1.CommPort = intport
            min.MSComm1.Settings = setting
            min.MSComm1.PortOpen = True
        End If
        min.MSComm1.InputMode = comInputModeText
        ssss = min.MSComm1.Input
        ss2 = ss2 + ssss
        lt = Len(ss2)
        sts = InStr(ss2, "@")
        et = InStr(ss2, "*")
        If et <> 0 Then st = InStrRev(ss2, "@", et)
        If et <> 0 Then
            If st = 0 Or et - st + 2 < 1 Then
                ss1 = ""
            Else
                ss1 = Mid(ss2, st, et - st + 2)
            End If
                ss2 = Right(ss2, lt - et)
            z = 0
        End If
        If sts > 0 Then Call O_Con
        If ss1 <> "" Then Call Ms_Do
        If InStr(ssss, "*") >= 1 Then
            If qsr0 Then qsr0 = False
            If qsr1 Then qsr1 = False
            'If qlisc Then qlisc = False
            min.Timer3.Enabled = True
            min.Timer5.Enabled = False
        ElseIf Len(ss1) > 3 Then
            Call O_Con
        End If
End Select
End Sub
Private Sub O_Con()
If Not MSComm1.PortOpen Then
    MSComm1.CommPort = intport
    MSComm1.Settings = setting
    MSComm1.PortOpen = True
End If
MSComm1.Output = Chr(13)
End Sub
 Private Sub Timer3_Timer()
If bchq Then
    MsgBox "check is wrong!", , "error"
   bchq = False
End If
If msgb <> "" Then
    MsgBox msgb, , til
    msgb = ""
End If
    min.Timer5.Enabled = True
    Call Msg_S
End Sub

Private Function Msg_S()
Static cc As Boolean
Static VN As Integer
If Len(o(2)) < 5 Then o(2) = ""
If o(2) = "" And o(3) <> "" Then o(2) = o(3): o(3) = o(4): o(4) = o(5): o(5) = o(6): o(6) = "" 'And Not o(1)q
If o(2) <> "" And Not qsr0 And Not qsr1 Then
    If Right(o(2), 2) = "*" & Chr(13) Then o(2) = Left(o(2), Len(o(2)) - 4)
    Call S_Check
    If Not min.MSComm1.PortOpen Then
        min.MSComm1.CommPort = intport
        min.MSComm1.Settings = setting
        min.MSComm1.PortOpen = True
    End If
     MSComm1.OutBufferCount = 0
    min.MSComm1.RThreshold = 1
     min.MSComm1.Output = o(2)
     qsr1 = True
ElseIf o(2) <> "" And qsr0 Then
    If Not min.MSComm1.PortOpen Then
        min.MSComm1.CommPort = intport
        min.MSComm1.Settings = setting
        min.MSComm1.PortOpen = True
    End If
     MSComm1.OutBufferCount = 0
     min.MSComm1.Output = "@00XZ42*" + o(0)
     qsr0 = False
ElseIf o(2) = "" And Not qsr0 Then
    MSComm1.OutBufferCount = 0
    o(1) = "@00RD39900230"
    Call S_Check
    If Not min.MSComm1.PortOpen Then
        min.MSComm1.CommPort = intport
        min.MSComm1.Settings = setting
        min.MSComm1.PortOpen = True
    End If
        MSComm1.OutBufferCount = 0
        min.MSComm1.RThreshold = 1
        min.MSComm1.Output = o(1)
        qsr0 = True
Else
    Call O_Con
End If
If min.Timer3.Enabled = False Then min.Timer3.Enabled = True
End Function
Private Sub Timer5_Timer()
z = z + 1
If z = 8 Then
    qsr0 = False
    qsr1 = False
    min.MSComm1.Output = "@00KC48*" + Chr(13)
    min.Timer3.Enabled = True
ElseIf z > 12 Then
    min.Timer5.Enabled = False
    min.Timer3.Enabled = False
    qsr0 = False
    qsr1 = False
    z = 0
    min.Timer3.Enabled = True
End If
End Sub
Private Sub r0_display()
 On Error Resume Next
Dim j, k, l As Integer
Dim i, q, n, m As Integer
Dim Ii(15, 9) As Boolean
Dim AH(11) As String
Dim stime(11) As String
Dim AHand(11) As String
If RID(13, 4) = "" Or Val(RID(13, 4)) = 0 Then Exit Sub
n = 0
For i = 1 To 5
    For q = 0 To 9
        If n < 39 Then Ntim(n) = RID(i, q)
        If n < 39 Then Ncode(n) = RID(i + 5, q) Else Ncode(n) = ""
        If n < 23 Then Ncode(n + 50) = RID(i + 10, q)
        If n < 10 Then
          AH(n) = RID(i + 18, q)
          stime(n) = RID(i + 19, q)
          AHand(i) = RID(Int(q / 5) + 20 + i, (q * 2) + 1) + RID(Int(q / 5) + 20 + i, (q * 2))
        End If
        n = n + 1
    Next q
    q = 0
Next i
  TIN_02 = True
For m = 0 To vv - 1
  MSFlexGrid1.TextMatrix(m + 1, 6) = Val(Ntim(m + 1))
  If Val(Ntim(m + 1)) <> 0 Then Ntim_1(m + 2) = Val(Ntim(m + 1))
  If TIN_01 = False And TIN_02 = True Then
     If Ncode(m + 1) = "0000" Then
        MSFlexGrid1.TextMatrix(m + 1, 4) = "": MSFlexGrid1.TextMatrix(m + 1, 5) = ""
      Else
          For k = 1 To min.MSFlexGrid2.Rows - 1
             If Val(Right(Ncode(m + 1), 2)) = Val(min.MSFlexGrid2.TextMatrix(k, 11)) Then
                min.MSFlexGrid1.TextMatrix(m + 1, 4) = min.MSFlexGrid2.TextMatrix(k, 0)
                min.MSFlexGrid1.TextMatrix(m + 1, 5) = Val(Left(Ncode(m + 1), 2))
                Exit For
             End If
          Next k
     End If
   Else
     If TIN_01 = True Then
        If m < 8 Then
           If Ncode(m + 51) = "0000" Then
                MSFlexGrid1.TextMatrix(m + 1, 4) = "": MSFlexGrid1.TextMatrix(m + 1, 5) = ""
            Else
                   For k = 1 To min.MSFlexGrid2.Rows - 1
                       If Val(Right(Ncode(m + 51), 2)) = Val(min.MSFlexGrid2.TextMatrix(k, 11)) Then
                          min.MSFlexGrid1.TextMatrix(m + 1, 4) = min.MSFlexGrid2.TextMatrix(k, 0)
                          min.MSFlexGrid1.TextMatrix(m + 1, 5) = Val(Left(Ncode(m + 51), 2))
                          Exit For
                      End If
                  Next k
           End If
         ElseIf m > 11 And m < 22 Then
               If Ncode(m + 51) = "0000" Then
                  MSFlexGrid1.TextMatrix(m + 1, 4) = "": MSFlexGrid1.TextMatrix(m + 1, 5) = ""
               Else
                  For k = 1 To min.MSFlexGrid2.Rows - 1
                      If Val(Right(Ncode(m + 51), 2)) = Val(min.MSFlexGrid2.TextMatrix(k, 11)) Then
                         min.MSFlexGrid1.TextMatrix(m + 1, 4) = min.MSFlexGrid2.TextMatrix(k, 0)
                         min.MSFlexGrid1.TextMatrix(m + 1, 5) = Val(Left(Ncode(m + 51), 2))
                         Exit For
                      End If
                  Next k
              End If
         Else
           If Ncode(m + 1) = "0000" Then
        MSFlexGrid1.TextMatrix(m + 1, 4) = "": MSFlexGrid1.TextMatrix(m + 1, 5) = ""
     Else
        For k = 1 To min.MSFlexGrid2.Rows - 1
          If Val(Right(Ncode(m + 1), 2)) = Val(min.MSFlexGrid2.TextMatrix(k, 11)) Then
            min.MSFlexGrid1.TextMatrix(m + 1, 4) = min.MSFlexGrid2.TextMatrix(k, 0)
            min.MSFlexGrid1.TextMatrix(m + 1, 5) = Val(Left(Ncode(m + 1), 2))
            Exit For
          End If
        Next k
     End If
        End If
   End If
End If
Next m
    For i = 0 To 3
        min.MSFlexGrid1.TextMatrix(i + 9, 7) = Val(RID(Int((2 * i) / 10) + 14, 2 * i - Int((2 * i) / 10) * 10))
        min.MSFlexGrid1.TextMatrix(i + 9, 8) = Val(RID(Int((2 * i + 1) / 10) + 14, 2 * i - Int((2 * i) / 10) * 10 + 1))
    Next i
    For i = 4 To 23
        min.MSFlexGrid1.TextMatrix(i + 19, 7) = Val(RID(Int((2 * i) / 10) + 14, 2 * i - Int((2 * i) / 10) * 10))
        min.MSFlexGrid1.TextMatrix(i + 19, 8) = Val(RID(Int((2 * i + 1) / 10) + 14, 2 * i - Int((2 * i) / 10) * 10 + 1))
    Next i
    dd = Int((2 * i + 1) / 10) + 14
    ddd = 2 * i - Int((2 * i) / 10) * 10 + 1
 If RID(6, 0) = "0000" Then
    min.Label2(0).Caption = "Without item"
   Else
      For k = 1 To min.MSFlexGrid2.Rows - 1
         If Val(Right(RID(6, 0), 2)) = min.MSFlexGrid2.TextMatrix(k, 11) Then
            min.Label2(0).Caption = "(" & Val(Left(RID(6, 0), 2)) & ")" & min.MSFlexGrid2.TextMatrix(k, 0)
         End If
      Next
End If
If RID(11, 0) = "0000" Then
   min.Label2(1).Caption = "without item "
Else
  For k = 1 To min.MSFlexGrid2.Rows - 1
    If Val(Right(RID(11, 0), 2)) = min.MSFlexGrid2.TextMatrix(k, 11) Then
       min.Label2(1).Caption = "(" & Val(Left(RID(11, 0), 2)) & ")" & min.MSFlexGrid2.TextMatrix(k, 0)
    End If
  Next
End If
 If RID(11, 9) = "0000" Then
    min.Label2(2).Caption = "without item"
   Else
      For k = 1 To min.MSFlexGrid2.Rows - 1
         If Val(Right(RID(11, 9), 2)) = min.MSFlexGrid2.TextMatrix(k, 11) Then
            min.Label2(2).Caption = "(" & Val(Left(RID(11, 9), 2)) & ")" & min.MSFlexGrid2.TextMatrix(k, 0)
         End If
      Next
End If
If RID(12, 0) = "0000" Then
   min.Label2(3).Caption = "without item"
Else
  For k = 1 To min.MSFlexGrid2.Rows - 1
    If Val(Right(RID(12, 0), 2)) = min.MSFlexGrid2.TextMatrix(k, 11) Then
       min.Label2(3).Caption = "(" & Val(Left(RID(12, 0), 2)) & ")" & min.MSFlexGrid2.TextMatrix(k, 0)
    End If
  Next
End If
 If RID(12, 1) = "0000" Then
    min.Label2(4).Caption = "without item"
   Else
      For k = 1 To min.MSFlexGrid2.Rows - 1
         If Val(Right(RID(12, 1), 2)) = min.MSFlexGrid2.TextMatrix(k, 11) Then
            min.Label2(4).Caption = "(" & Val(Left(RID(12, 1), 2)) & ")" & min.MSFlexGrid2.TextMatrix(k, 0)
         End If
      Next
End If
If RID(12, 2) = "0000" Then
   min.Label2(5).Caption = "without item"
Else
  For k = 1 To min.MSFlexGrid2.Rows - 1
    If Val(Right(RID(12, 2), 2)) = min.MSFlexGrid2.TextMatrix(k, 11) Then
       min.Label2(5).Caption = "(" & Val(Left(RID(12, 2), 2)) & ")" & min.MSFlexGrid2.TextMatrix(k, 0)
    End If
  Next
End If
min.Text1(1).Text = Val(Right(RID(0, 0), 2))
min.Text1(0).Text = Val(Left(RID(0, 0), 2))
min.Text2(0).Text = Val(Left((RID(0, 1)), 2))
min.Text2(1).Text = Val(Right((RID(0, 1)), 2))
min.Text9(0).Text = Str(Val(RID(0, 2))) & "s"
min.Text9(1).Text = Str(Val(RID(0, 3))) & "s"
min.Text11(0).Text = Str(Val(RID(5, 7)))
min.Text11(1).Text = Str(Val(RID(5, 8)))
min.Text11(2).Text = Str(Val(RID(5, 9)))
min.Label5(0).Caption = CStr(Val(RID(13, 8)))
min.Label5(1).Caption = CStr(Val(RID(13, 9)))
min.Text6.Text = "process--- " & Left(RID(0, 5), 1)

If Val(Mid(RID(0, 5), 2, 1)) = 0 Then
    min.Text7.Text = "service"
    min.Text7.Locked = False
    min.Command3(0).Enabled = False
    min.Command3(1).Enabled = True
    min.Command3(2).Enabled = False
    For i = 0 To 6
        min.Command1(i).Enabled = True
        min.Command11(i).Enabled = True
    Next i
ElseIf Val(Mid(RID(0, 5), 2, 1)) = 1 Then
    min.Text7.Text = "Manu"
    min.Text7.Locked = True
    min.Command3(0).Enabled = True
    min.Command3(2).Enabled = True
    min.Command3(1).Enabled = False
    For i = 0 To 3
        min.Command1(i).Enabled = True
        min.Command11(i).Enabled = True
    Next i
    For i = 4 To 5
        min.Command1(i).Enabled = False
        min.Command11(i).Enabled = False
    Next i
    min.Command1(6).Enabled = True
    min.Command11(6).Enabled = True
ElseIf Val(Mid(RID(0, 5), 2, 1)) = 2 Then
    min.Text7.Text = "Auto"
    min.Text7.Locked = True
    min.Command3(2).Enabled = False
    min.Command3(0).Enabled = False
    min.Command3(1).Enabled = True
    For i = 0 To 6
        min.Command1(i).Enabled = False
        min.Command11(i).Enabled = False
    Next i
    
End If
If RID(0, 4) = "0000" Then
  min.Text10.Text = "Failure-free！"
Else
  Dim rs1 As New Recordset
  Dim sq1 As String
  If eros <> RID(0, 4) Then
   sq1 = "select * from 故障代码  where 代码 ='" & RID(0, 4) & "'"
   rs1.Open sq1, conn, adOpenKeyset, adLockPessimistic
    Do While Not rs1.EOF
      min.Text10.Text = rs1.Fields(1)       ' & rs1.Fields(1)
      rs1.MoveNext
    Loop
    rs1.Close
    eros = RID(0, 4)
  End If
End If
For i = 0 To 9
    Me.Label7(i).Caption = Val(AH(i))
    Me.Label8(i).Caption = Val(stime(i))
    Me.Label9(i).Caption = Val(AHand(i))
Next
For i = 1 To 4
    hn = Mid(RID(5, 2), i, 1)
    Call hex_bin
    Ii(15 + 4 - 4 * i, 2) = bn(0)
    Ii(15 + 3 - 4 * i, 2) = bn(1)
    Ii(15 + 2 - 4 * i, 2) = bn(2)
    Ii(15 + 1 - 4 * i, 2) = bn(3)
    hn = Mid(RID(5, 3), i, 1)
    Call hex_bin
    Ii(15 + 4 - 4 * i, 3) = bn(0)
    Ii(15 + 3 - 4 * i, 3) = bn(1)
    Ii(15 + 2 - 4 * i, 3) = bn(2)
    Ii(15 + 1 - 4 * i, 3) = bn(3)
    hn = Mid(RID(5, 4), i, 1)
    Call hex_bin
    Ii(15 + 4 - 4 * i, 4) = bn(0)
    Ii(15 + 3 - 4 * i, 4) = bn(1)
    Ii(15 + 2 - 4 * i, 4) = bn(2)
    Ii(15 + 1 - 4 * i, 4) = bn(3)
    hn = Mid(RID(5, 6), i, 1)
    Call hex_bin
    Ii(15 + 4 - 4 * i, 5) = bn(0)
    Ii(15 + 3 - 4 * i, 5) = bn(1)
    Ii(15 + 2 - 4 * i, 5) = bn(2)
    Ii(15 + 1 - 4 * i, 5) = bn(3)
    hn = Mid(RID(5, 5), i, 1)
    Call hex_bin
    Ii(15 + 4 - 4 * i, 6) = bn(0)
    Ii(15 + 3 - 4 * i, 6) = bn(1)
    Ii(15 + 2 - 4 * i, 6) = bn(2)
    Ii(15 + 1 - 4 * i, 6) = bn(3)
    hn = Mid(RID(13, 5), i, 1)
    Call hex_bin
    Ii(15 + 4 - 4 * i, 7) = bn(0)
    Ii(15 + 3 - 4 * i, 7) = bn(1)
    Ii(15 + 2 - 4 * i, 7) = bn(2)
    Ii(15 + 1 - 4 * i, 7) = bn(3)
    hn = Mid(RID(13, 6), i, 1)
    Call hex_bin
    Ii(15 + 4 - 4 * i, 8) = bn(0)
    Ii(15 + 3 - 4 * i, 8) = bn(1)
    Ii(15 + 2 - 4 * i, 8) = bn(2)
    Ii(15 + 1 - 4 * i, 8) = bn(3)
    hn = Mid(RID(13, 7), i, 1)
    Call hex_bin
    Ii(15 + 4 - 4 * i, 9) = bn(0)
    Ii(15 + 3 - 4 * i, 9) = bn(1)
    Ii(15 + 2 - 4 * i, 9) = bn(2)
    Ii(15 + 1 - 4 * i, 9) = bn(3)
Next i
For i = 0 To 11
    If Ii(i, 6) Then min.Check1(i).Value = 1 Else min.Check1(i).Value = 0
    If Ii(i, 4) Then min.Check2(i).Value = 1 Else min.Check2(i).Value = 0
    If Ii(i, 2) Then min.Check3(i).Value = 1 Else min.Check3(i).Value = 0
    If Ii(i, 8) Then Me.Label12(i).BackColor = &HFF& Else Me.Label12(i).BackColor = &HFF00&
Next i
For i = 0 To 9
    If Ii(i, 5) Then min.Check5(i).Value = 1 Else min.Check5(i).Value = 0
    If Ii(i, 7) Then Me.Label6(i).BackColor = &HFF& Else Me.Label6(i).BackColor = &HFF00&
Next i
For i = 0 To 13
    If Ii(i, 3) Then min.Check4(i).Value = 1 Else min.Check4(i).Value = 0
    If Ii(i, 9) Then Me.Label13(i).BackColor = &HFF& Else Me.Label13(i).BackColor = &HFF00&
Next i
If Ii(12, 5) Then Command3(4).Caption = "AGT-ON" Else Command3(4).Caption = "AGT-OFF"
If Ii(14, 4) Then Command3(3).Caption = "Cont" Else Command3(3).Caption = "once"
  If Text7.Text <> "service" Then
    If Ii(15, 4) Then
       estr = "Auto"
    Else
       estr = "Manul"
    End If
  Else
    estr = "service"
  End If
'If Ii(10, 6) Or Ii(11, 6) Then MsgBox "请把触摸屏电流控制上板料号开关关闭！", , "警告"  'If Ii(11, 6) Then MsgBox "请把触摸屏电流控制上板料号开关关闭！", , "警告"
For i = 1 To vv
   min.MSFlexGrid1.TextStyle = 0
   min.MSFlexGrid1.CellPictureAlignment = 4
   min.MSFlexGrid1.FillStyle = flexFillSingle
If Val(Left(RID(0, 1), 2)) = i Then
   min.MSFlexGrid1.TextMatrix(i, 3) = "A"
 ElseIf Val(Right(RID(0, 1), 2)) = i Then
   min.MSFlexGrid1.TextMatrix(i, 3) = "B"
ElseIf Val(Left(RID(0, 0), 2)) = i Then
   If Val(Left(RID(0, 0), 2)) > Val(Left(RID(0, 1), 2)) Then
      min.MSFlexGrid1.TextMatrix(i, 3) = "▲"  '△▲▼▽
    Else
      min.MSFlexGrid1.TextMatrix(i, 3) = "▼"
    End If
 ElseIf Val(Right(RID(0, 0), 2)) = i Then
   If Val(Right(RID(0, 0), 2)) > Val(Right(RID(0, 1), 2)) Then
      min.MSFlexGrid1.TextMatrix(i, 3) = "△"
    Else
      min.MSFlexGrid1.TextMatrix(i, 3) = "▽"
    End If
 Else
   If i <> Val(RID(0, 0)) Then
     min.MSFlexGrid1.TextMatrix(i, 3) = ""
   End If
 End If
 Next i
 Timer10 = True
End Sub
Private Sub Timer6_Timer()
'If Not Add_liao.Visible Then
 If Text5.Text <> "load complete" Then Call r0_display: r0q = False
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
qsr0 = False
 qsr1 = False
 qsr2 = False
 ss1 = ""
 ss2 = ""
 ss3 = ""
 qbch = False
 r(0) = "": r(1) = ""
 ssss = ""
 For i = 0 To 3
  bn(i) = False
 Next
 For i = 0 To 7
  r(i) = ""
Next
For i = 0 To 9
   o(9) = ""
Next
r0q = False
xx = ""
 strr1 = ""
 i = q = 0
cher0 = ""
 cher1 = ""
If Me.MSComm1.PortOpen = True Then Me.MSComm1.PortOpen = False
Call REGEDIT
Call T_exit
MsgBox "你关掉后生产记录将会记录不准确！", 48, "提示"
End
End Sub
Sub T_exit()
 On Error Resume Next
         Open (App.Path & "\STYLE\系统日志.ini") For Input As #1
          Do While Not EOF(1)
               Line Input #1, INTEXT
               TSTR = TSTR + INTEXT + Chr(13) + Chr(10)
          Loop
      Close #1
      TSTR = TSTR + "   " + userID + "               " + Format(Now, "yyyy-mm-dd hh:mm:ss") + "            " + "退出系统" + Chr(13) + Chr(10)
        If Len(TSTR) > 10000 Then TSTR = Right(TSTR, 9800)
      Open (App.Path & "\STYLE\系统日志.ini") For Output As #1
      Print #1, TSTR
      Close #1
      End
End Sub
Sub T_in()
Dim s As Long
s = FileLen(App.Path & "\style\系统日志.ini")
 Open (App.Path & "\style\系统日志.ini") For Input As #1
    Do While Not EOF(1)
         Line Input #1, INTEXT
         TSTR = TSTR + INTEXT + Chr(13) + Chr(10)
    Loop
Close #1
If s > 10000 Then TSTR = Right(TSTR, 9971)
TSTR = TSTR + "   " + userID + "               " + Format(Now, "yyyy-mm-dd hh:mm:ss") + "            " + "系统登录" + Chr(13) + Chr(10)
        If Len(TSTR) > 10000 Then TSTR = Right(TSTR, 9800)
Open (App.Path & "\style\系统日志.ini") For Output As #1
      Print #1, TSTR
Close #1
End Sub
Sub D_ING()
 Open (App.Path & "\style\使用料号.ini") For Input As #4
                          Do While Not EOF(4)
                               Line Input #4, INTEXT
                               TSTR = TSTR + INTEXT + Chr(13) + Chr(10)
                          Loop
                      Close #4
                      TSTR = TSTR + "   " + a_instr + "               " + Format(Now, "yyyy-mm-dd hh:mm:ss") + "            " + "删除操作" + Chr(13) + Chr(10)
        If Len(TSTR) > 10000 Then TSTR = Right(TSTR, 9800)
                      Open (App.Path & "\style\使用料号.ini") For Output As #4
                            Print #4, TSTR
                      Close #4
End Sub

Private Sub Timer8_Timer()
 min.Label3.Visible = False
 min.Timer8.Enabled = False
End Sub

Private Sub Timer9_Timer()
If Me.MSComm1.PortOpen = True Then Me.MSComm1.PortOpen = False
qsr0 = False
Me.MSComm1.InBufferCount = 0
Me.MSComm1.OutBufferCount = 0
Form_Load
Me.Timer9.Enabled = True
End Sub
Private Sub Timer10_Timer()
 For m = 0 To vv
   If 8 < m And m < 13 Or m > 22 Then
     If Ncode_1(m) <> Ncode(m) Then
      If Ncode(m + 0) <> "0000" Then '
        add_n(1) = min.MSFlexGrid1.TextMatrix(m + 0, 1)
        add_n(2) = min.MSFlexGrid1.TextMatrix(m + 0, 4)
        add_n(3) = min.MSFlexGrid1.TextMatrix(m + 0, 5)
        If 8 < m And m < 13 Then
           For k = 1 To min.MSFlexGrid2.Rows - 1
                If add_n(2) = Trim(min.MSFlexGrid2.TextMatrix(k, 0)) Then
                   add_n(4) = Int(Val(min.MSFlexGrid2.TextMatrix(k, 1) * Val(Left(Ncode(m + 0), 2))))
                   add_n(5) = Int(Val(min.MSFlexGrid2.TextMatrix(k, 2) * Val(Left(Ncode(m + 0), 2))))
                   Exit For
                End If
            Next k
         Else
            For k = 1 To min.MSFlexGrid2.Rows - 1
                If add_n(2) = Trim(min.MSFlexGrid2.TextMatrix(k, 0)) Then
                   add_n(4) = Int(Val(min.MSFlexGrid2.TextMatrix(k, 6) * Val(Left(Ncode(m + 0), 2))))
                   add_n(5) = Int(Val(min.MSFlexGrid2.TextMatrix(k, 7) * Val(Left(Ncode(m + 0), 2))))
                   Exit For
                End If
            Next k
         End If
       ' add_n(4) = min.MSFlexGrid1.TextMatrix(m + 0, 7)
        'add_n(5) = min.MSFlexGrid1.TextMatrix(m + 0, 8)
        If add_n(2) = "" Then add_n(2) = "Item is not clear"
        If add_n(3) = "" Then add_n(3) = "quanlity is not clear"
        If add_n(4) = "" Then add_n(4) = "Apere is not clear"
        If add_n(5) = "" Then add_n(5) = "Apere is not clear"
        add_n(6) = "processing "
        add_n(7) = Trim(Now)
        add_n(8) = "processing"
        add_n(9) = "processing"
        min.Adodc5.Refresh
       Do While (min.Adodc5.Recordset.EOF = False)
             If add_n(1) = min.Adodc5.Recordset.Fields(1) And add_n(2) = min.Adodc5.Recordset.Fields(2) And add_n(6) = min.Adodc5.Recordset.Fields(6) Then
                 ad_b = True
                 Exit Do
              Else
                 ad_b = False
                 min.Adodc5.Recordset.MoveNext
              End If
       Loop
        If ad_b = False Then
           min.Adodc5.Recordset.AddNew
               For k = 1 To 9
                    min.Adodc5.Recordset.Fields(k) = add_n(k)
               Next k
             min.Adodc5.Recordset.Update
             On Error GoTo add_n
        End If
    Else
      If Ncode(m + 0) = "0000" Then  'Ncode_1(m) <> Ncode(m)
       min.Adodc5.Refresh
       Do While (min.Adodc5.Recordset.EOF = False)
          If min.Adodc5.Recordset.Fields(1) = min.MSFlexGrid1.TextMatrix(m + 0, 1) And min.Adodc5.Recordset.Fields(6) = "待出板" Then
              add_n(1) = min.Adodc5.Recordset.Fields(1)
              add_n(2) = min.Adodc5.Recordset.Fields(2)
              add_n(3) = min.Adodc5.Recordset.Fields(3)
              add_n(4) = min.Adodc5.Recordset.Fields(4)
              add_n(5) = min.Adodc5.Recordset.Fields(5)
              add_n(7) = min.Adodc5.Recordset.Fields(7)
              add_n(8) = Trim(Now)
              add_n(6) = Ntim_1(m + 0)
              If Trim(add_n(6)) = "" Then add_n(6) = 0
              add_n(9) = Int(((Val(add_n(4)) + Val(add_n(5))) * Val(add_n(6))) / 3600)
              For k = 1 To 9
                    min.Adodc5.Recordset.Fields(k) = add_n(k)
               Next k
               On Error GoTo add_n
               min.Adodc5.Recordset.Update
              Exit Do
          Else
              min.Adodc5.Recordset.MoveNext
          End If
      Loop
       End If
      End If
     Ncode_1(m) = Ncode(m)
      End If
      End If
   Next m
'   Me.Adodc5.Recordset.Close
add_n:
    Exit Sub 'MsgBox "生产记录出错", , ""
End Sub

Private Sub Timer11_Timer()
 If Val(Hour(Time)) > 9 And bbbb_1 = False Then MsgBox "You must record the production record, open the production record direct point (record sorting)", 48, "Prompt action": bbbb = True
ggg = Val(Hour(Time))
End Sub
Private Sub Timer12_Timer()
Me.Label4.Caption = Trim(RID(12, 5))
End Sub
Private Sub Label4_Change()
 Me.Timer13.Enabled = False
 Me.Timer13.Enabled = True
End Sub
Private Sub Timer13_Timer()
Static otim As String
If otim <> RID(13, 4) Then otim = RID(13, 4): Exit Sub
If Me.MSComm1.PortOpen = True Then Me.MSComm1.PortOpen = False
qsr0 = False
 qsr1 = False
 qsr2 = False
 ss1 = ""
 ss2 = ""
 ss3 = ""
 qbch = False
 r(0) = "": r(1) = ""
 ssss = ""
 For i = 0 To 3
  bn(i) = False
 Next
 For i = 0 To 7
  r(i) = ""
Next
r0q = False
xx = ""
 strr1 = ""
 i = q = 0
cher0 = ""
 cher1 = ""
 Form_Load
End Sub

Private Sub Text5_Change()
 Me.Timer13.Enabled = False
 Me.Timer13.Enabled = True
End Sub

