VERSION 5.00
Begin VB.Form F_a_user_add 
   Caption         =   "添加/删除用户"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "F_a_user_add.frx":0000
   ScaleHeight     =   8685
   ScaleWidth      =   12000
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "返回首页"
      Height          =   495
      Left            =   7440
      TabIndex        =   23
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox Text10 
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
      Left            =   7320
      TabIndex        =   22
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox Text9 
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
      Left            =   3120
      TabIndex        =   21
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox Text8 
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
      Left            =   7320
      TabIndex        =   20
      Top             =   5400
      Width           =   1935
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
      Left            =   3105
      TabIndex        =   8
      Top             =   3555
      Width           =   1935
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
      Left            =   3105
      TabIndex        =   7
      Top             =   4155
      Width           =   1935
   End
   Begin VB.TextBox Text3 
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
      Left            =   3105
      TabIndex        =   6
      Top             =   4755
      Width           =   1935
   End
   Begin VB.TextBox Text4 
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
      Left            =   3105
      TabIndex        =   5
      Top             =   5355
      Width           =   1935
   End
   Begin VB.TextBox Text5 
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
      Left            =   7305
      TabIndex        =   4
      Top             =   3555
      Width           =   1935
   End
   Begin VB.TextBox Text6 
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
      Left            =   7305
      TabIndex        =   3
      Top             =   4155
      Width           =   1935
   End
   Begin VB.TextBox Text7 
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
      Left            =   7305
      TabIndex        =   2
      Top             =   4755
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label Label12 
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
      Left            =   6840
      TabIndex        =   24
      Top             =   1440
      Width           =   4860
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "地址："
      Height          =   195
      Left            =   6360
      TabIndex        =   19
      Top             =   5400
      Width           =   540
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail："
      Height          =   195
      Left            =   6360
      TabIndex        =   18
      Top             =   6120
      Width           =   600
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tel："
      Height          =   195
      Left            =   2280
      TabIndex        =   17
      Top             =   6120
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "添加/删除用户"
      Height          =   195
      Left            =   1080
      TabIndex        =   16
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名："
      Height          =   195
      Left            =   2265
      TabIndex        =   15
      Top             =   3675
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码："
      Height          =   195
      Left            =   2265
      TabIndex        =   14
      Top             =   4275
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名："
      Height          =   195
      Left            =   2265
      TabIndex        =   13
      Top             =   4875
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性别："
      Height          =   195
      Left            =   2265
      TabIndex        =   12
      Top             =   5475
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单位："
      Height          =   195
      Left            =   6345
      TabIndex        =   11
      Top             =   3675
      Width           =   540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "类型："
      Height          =   195
      Left            =   6345
      TabIndex        =   10
      Top             =   4275
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "身份证号："
      Height          =   195
      Left            =   6345
      TabIndex        =   9
      Top             =   4800
      Width           =   900
   End
End
Attribute VB_Name = "F_a_user_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim ruser$, rpsw$, rname$, rsex$, company$, rtype$, idnumber$, address$, tel$, email$
ruser1 = Text1.text
rpsw1 = Text2.text
rname1 = Text3.text
rsex1 = Text4.text
company1 = Text5.text
rtype1 = Text6.text
idnumber1 = Text7.text
address1 = Text8.text
tel1 = Text9.text
email1 = Text10.text

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim stru$, strp$, strsql$
   
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset

conn.Open "provider=SQLOLEDB.1;datasource=(local);persist security info=false;integrated security=sspi;database=bms"
rs.Open "select * from ruser  ", conn
conn.Execute "insert into ruser(ruser , rpsw, rname, rsex, company, rtype, idnumber, address, tel, email) values('" & ruser1 & "','" & rpsw1 & "','" & rname1 & "','" & rsex1 & "','" & company1 & "','" & rtype1 & "','" & idnumber1 & "','" & address1 & "','" & tel1 & "','" & email1 & "')"
MsgBox "用户添加成功！"
End Sub

Private Sub Command2_Click()
Dim ruser$, rpsw$, rname$, rsex$, company$, rtype$, idnumber$, address$, tel$, email$
ruser2 = Text1.text
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim stru$, strp$, strsql$

Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "provider=SQLOLEDB.1;datasource=(local);persist security info=false;integrated security=sspi;database=bms"
rs.Open "select * from ruser  ", conn
conn.Execute "delete from ruser where ruser= '" & ruser2 & "'  "

MsgBox "用户删除成功！"
End Sub

Private Sub Command3_Click()
F_index.Show
Me.Hide
End Sub
