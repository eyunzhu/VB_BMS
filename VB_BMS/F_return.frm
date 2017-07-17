VERSION 5.00
Begin VB.Form F_return 
   Caption         =   "图书归还"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   Picture         =   "F_return.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   11970
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "返回首页"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "还书"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   5280
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
      TabIndex        =   3
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入归还书号："
      Height          =   180
      Left            =   1800
      TabIndex        =   2
      Top             =   3480
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "图书归还："
      Height          =   180
      Left            =   960
      TabIndex        =   1
      Top             =   2520
      Width           =   900
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
      TabIndex        =   0
      Top             =   1440
      Width           =   4860
   End
End
Attribute VB_Name = "F_return"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim time
time = Now

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim stru$, strp$, strsql$
   
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset

conn.Open "provider=SQLOLEDB.1;datasource=(local);persist security info=false;integrated security=sspi;database=bms"
rs.Open "select * from bor  ", conn
conn.Execute "UPDATE book SET exist = '是' WHERE bno = '" & Text1.text & "' "

conn.Execute "UPDATE bor SET returntime = '" & time & "' ,isreturn='是' WHERE bno = '" & Text1.text & "' "
MsgBox "还书成功！"

End Sub
Private Sub Command2_Click()
F_index.Show
Me.Hide
End Sub
