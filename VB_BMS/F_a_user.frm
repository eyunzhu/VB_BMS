VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form F_a_user 
   Caption         =   "读者用户管理"
   ClientHeight    =   8850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   Picture         =   "F_a_user.frx":0000
   ScaleHeight     =   8850
   ScaleWidth      =   12090
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "注销用户"
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "注销用户"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加用户"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   7680
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   9120
      Top             =   7440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=BMS;Data Source=."
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=BMS;Data Source=."
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from ruser"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "F_a_user.frx":11CA7
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7011
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Books Management System"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Left            =   6960
      TabIndex        =   4
      Top             =   1440
      Width           =   4860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "读者用户："
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   2280
      Width           =   900
   End
End
Attribute VB_Name = "F_a_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
F_a_user_add.Show
Me.Hide
End Sub

Private Sub Command2_Click()
F_a_user_add.Show
Me.Hide
End Sub

Private Sub Form_Load()
Adodc1.Visible = False
DataGrid1.Columns(0).Width = 1200
DataGrid1.Columns(1).Width = 1000
DataGrid1.Columns(2).Width = 800
DataGrid1.Columns(3).Width = 250
DataGrid1.Columns(4).Width = 1200
DataGrid1.Columns(5).Width = 400
DataGrid1.Columns(6).Width = 1400
DataGrid1.Columns(7).Width = 200
DataGrid1.Columns(8).Width = 1100
DataGrid1.Columns(9).Width = 1100
DataGrid1.Columns(10).Width = 900
DataGrid1.Columns(11).Width = 1600
End Sub
