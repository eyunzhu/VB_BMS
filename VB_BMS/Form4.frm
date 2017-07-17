VERSION 5.00
Begin VB.Form F_self 
   Caption         =   "个人信息"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   11970
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "返回首页"
      Height          =   375
      Left            =   9480
      TabIndex        =   23
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   2400
      TabIndex        =   22
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   2400
      TabIndex        =   21
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   2400
      TabIndex        =   20
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   2400
      TabIndex        =   19
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   2400
      TabIndex        =   18
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   7560
      TabIndex        =   17
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   7560
      TabIndex        =   16
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   7560
      TabIndex        =   15
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   7560
      TabIndex        =   14
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7560
      TabIndex        =   13
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "身份证号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   11
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "罚款金额："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   10
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "办理日期："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   9
      Top             =   3000
      Width           =   1275
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "账号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   8
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "类型："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   7
      Top             =   4200
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "住址："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   6
      Top             =   3600
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "电话："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   5
      Top             =   4200
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "邮箱："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   4
      Top             =   4800
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   3
      Top             =   2400
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性别："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   2
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单位："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   1
      Top             =   3600
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "个人信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1020
   End
End
Attribute VB_Name = "F_self"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
F_index.Show
Me.Hide
End Sub

Private Sub Form_Load()

Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str1 As String

db.Open "provider=SQLOLEDB.1;datasource=(local);persist security info=false;integrated security=sspi;database=bms"
str1 = "select * from ruser where ruser= '" & username & " ' "
rs.Open str1, db, 1
If rs.RecordCount <= 0 Then
MsgBox "没有记录!"
rs.Close
Exit Sub
End If
Text1.text = rs.Fields("ruser").Value
Text11.text = rs.Fields("rname").Value
Text10.text = rs.Fields("rsex").Value
Text9.text = rs.Fields("company").Value
Text8.text = rs.Fields("rtype").Value
Text7.text = rs.Fields("email").Value
Text6.text = rs.Fields("idnumber").Value
Text5.text = rs.Fields("fine").Value
Text4.text = rs.Fields("date").Value
Text3.text = rs.Fields("address").Value
Text2.text = rs.Fields("tel").Value
rs.Close
db.Close
End Sub

