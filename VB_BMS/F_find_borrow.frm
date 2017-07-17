VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form F_find_borrow 
   Caption         =   "图书查询借阅"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   Picture         =   "F_find_borrow.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   11955
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "返回首页"
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "借阅"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   7080
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   9480
      Top             =   7920
      Width           =   3255
      _ExtentX        =   5741
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "F_find_borrow.frx":11CA7
      Height          =   3255
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   5741
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
         Size            =   9
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查找"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "F_find_borrow.frx":11CBC
      Left            =   1800
      List            =   "F_find_borrow.frx":11CCC
      TabIndex        =   1
      Text            =   "查找方式"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label4 
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
      TabIndex        =   10
      Top             =   1440
      Width           =   4860
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "所借书号："
      Height          =   180
      Left            =   2040
      TabIndex        =   6
      Top             =   7200
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "图书借阅："
      Height          =   180
      Left            =   1080
      TabIndex        =   5
      Top             =   6840
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "图书查找："
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   2280
      Width           =   900
   End
End
Attribute VB_Name = "F_find_borrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim find$, text$
find = Combo1.text
text = Text1.text

Adodc1.ConnectionString = "provider=SQLOLEDB.1;datasource=(local);persist security info=false;integrated security=sspi;database=bms"

If find = "书号" Then

    SQL = "select * from book where bno='" & text & "'"
    Adodc1.RecordSource = SQL
    Adodc1.Refresh
End If

If find = "书名" Then

    SQL = "select * from book where bname  like '%" & text & "%'"
    Adodc1.RecordSource = SQL
    Adodc1.Refresh
    
End If
If find = "作者" Then

    SQL = "select * from book where author like '%" & text & "%'"
    Adodc1.RecordSource = SQL
    Adodc1.Refresh
End If
If find = "摘要" Then

    SQL = "select * from book where abstract like '%" & text & "%'"
    Adodc1.RecordSource = SQL
    Adodc1.Refresh
     
End If

End Sub

Private Sub Command2_Click()
Dim time
time = Now

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim stru$, strp$, strsql$
   
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset

conn.Open "provider=SQLOLEDB.1;datasource=(local);persist security info=false;integrated security=sspi;database=bms"

rs.Open "select * from bor  ", conn

conn.Execute "insert into bor(ruser , bno, borrowtime) values('" & username & "','" & Text2.text & "','" & time & "')"

conn.Execute "UPDATE book SET exist = '否' WHERE bno = '" & Text2.text & "' "

MsgBox "图书借阅成功！"

End Sub

Private Sub Command3_Click()
F_index.Show
Me.Hide
End Sub

Private Sub Form_Load()

Adodc1.Visible = False

End Sub
