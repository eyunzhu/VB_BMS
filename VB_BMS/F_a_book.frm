VERSION 5.00
Begin VB.Form F_a_book 
   Caption         =   "图书管理"
   ClientHeight    =   8910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "F_a_book.frx":0000
   ScaleHeight     =   8910
   ScaleWidth      =   12000
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "返回首页"
      Height          =   495
      Left            =   7080
      Picture         =   "F_a_book.frx":11CA7
      TabIndex        =   17
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除"
      Height          =   495
      Left            =   5520
      TabIndex        =   16
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加"
      Height          =   495
      Left            =   3960
      TabIndex        =   15
      Top             =   6240
      Width           =   975
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
      Left            =   7553
      TabIndex        =   14
      Top             =   4568
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
      Left            =   7553
      TabIndex        =   13
      Top             =   3968
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
      Left            =   7553
      TabIndex        =   12
      Top             =   3368
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
      Left            =   3353
      TabIndex        =   11
      Top             =   5168
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
      Left            =   3353
      TabIndex        =   10
      Top             =   4568
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
      Left            =   3353
      TabIndex        =   9
      Top             =   3968
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
      Left            =   3353
      TabIndex        =   8
      Top             =   3368
      Width           =   1935
   End
   Begin VB.Label Label9 
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
      TabIndex        =   18
      Top             =   1440
      Width           =   4860
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "摘要："
      Height          =   195
      Left            =   6593
      TabIndex        =   7
      Top             =   4688
      Width           =   540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "价格："
      Height          =   195
      Left            =   6593
      TabIndex        =   6
      Top             =   4088
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "出版社："
      Height          =   195
      Left            =   6593
      TabIndex        =   5
      Top             =   3488
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分类号："
      Height          =   195
      Left            =   2513
      TabIndex        =   4
      Top             =   5288
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "作者："
      Height          =   195
      Left            =   2513
      TabIndex        =   3
      Top             =   4688
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "书名："
      Height          =   195
      Left            =   2513
      TabIndex        =   2
      Top             =   4088
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "书号："
      Height          =   195
      Left            =   2513
      TabIndex        =   1
      Top             =   3488
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "添加/删除图书："
      Height          =   180
      Left            =   1200
      TabIndex        =   0
      Top             =   2640
      Width           =   1350
   End
End
Attribute VB_Name = "F_a_book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim bno$, bname$, author$, classno$, press$, price$, abstract$

bno1 = Text1.text
bname1 = Text2.text
author1 = Text3.text
classno1 = Text4.text
press1 = Text5.text
price1 = Text6.text
abstract1 = Text7.text

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim stru$, strp$, strsql$
   
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset

conn.Open "provider=SQLOLEDB.1;datasource=(local);persist security info=false;integrated security=sspi;database=bms"


rs.Open "select * from book  ", conn

'Set MSHFlexGrid.DataSource = rst

conn.Execute "insert into book(bno,bname,author,classno,press,price,abstract) values('" & bno1 & "','" & bname1 & "','" & author1 & "','" & classno1 & "','" & press1 & "','" & price1 & "','" & abstract1 & "')"

 'conn.Execute "insert into book(bno) values('vdevdsv')"

MsgBox "图书添加成功！"
End Sub

Private Sub Command2_Click()

Dim bno$

bno2 = Text1.text

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim stru$, strp$, strsql$
   
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset

conn.Open "provider=SQLOLEDB.1;datasource=(local);persist security info=false;integrated security=sspi;database=bms"


rs.Open "select * from book  ", conn


conn.Execute "delete from book where bno= '" & bno2 & "'  "

MsgBox "图书删除成功！"
End Sub


Private Sub Command3_Click()
F_index.Show
Me.Hide
End Sub

Private Sub Form_Load()

End Sub
