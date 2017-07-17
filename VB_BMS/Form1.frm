VERSION 5.00
Begin VB.Form F_login 
   Caption         =   "登陆"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   11940
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "重新输入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "管理员登陆"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "读者登陆"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6000
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   4320
      Width           =   1740
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   6000
      TabIndex        =   3
      Top             =   3465
      Width           =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密  码："
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4920
      TabIndex        =   2
      Top             =   4440
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名："
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4920
      TabIndex        =   1
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "图书管理系统"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   840
      Left            =   3840
      TabIndex        =   0
      Top             =   1560
      Width           =   5040
   End
End
Attribute VB_Name = "F_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim stru$, strp$, strsql$
Static count As Byte
stru = Text1.text
strp = Text2.text
username = Text1.text

If stru = "" Then
MsgBox "用户名不能为空，请输入用户名！", , "登陆错误"
Text1.SetFocus
Exit Sub
ElseIf strp = "" Then
MsgBox "密码不能为空，请输入密码！", , "登陆错误"
Text2.SetFocus
Exit Sub
End If
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset

conn.Open "provider=SQLOLEDB.1;datasource=(local);persist security info=false;integrated security=sspi;database=bms"


rs.Open "select * from ruser where ruser='" & stru & " 'and rpsw='" & strp & "'; ", conn

If rs.EOF Then
count = count + 1
MsgBox "用户名不存在或者密码错误！", , "登录失败"
Text1.text = ""
Text2.text = ""
Text1.SetFocus
Else
logins = True
username = rs("ruser").Value
pass = rs("rpsw").Value
F_index.Show
Me.Hide
End If
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
If count >= 3 Then
MsgBox "超过登录次数，无权登录本系统！", , "登录失败"
End
End If
End Sub


Private Sub Command2_Click()

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim stru$, strp$, strsql$
Static count As Byte
stru = Text1.text
strp = Text2.text
If stru = "" Then
MsgBox "用户名不能为空，请输入用户名！", , "登陆错误"
Text1.SetFocus
Exit Sub
ElseIf strp = "" Then
MsgBox "密码不能为空，请输入密码！", , "登陆错误"
Text2.SetFocus
Exit Sub
End If
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset

conn.Open "provider=SQLOLEDB.1;datasource=(local);persist security info=false;integrated security=sspi;database=bms"


rs.Open "select * from auser where auser='" & stru & " 'and apsw='" & strp & "'; ", conn

If rs.EOF Then
count = count + 1
MsgBox "用户名不存在或者密码错误！", , "登录失败"
Text1.text = ""
Text2.text = ""
Text1.SetFocus
Else
logins = True
username = rs("auser").Value
pass = rs("apsw").Value
F_a_index.Show
Me.Hide
End If
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
If count >= 3 Then
MsgBox "超过登录次数，无权登录本系统！", , "登录失败"
End
End If
End Sub

Private Sub Command3_Click()
Text1.text = ""
Text2.text = ""
Text1.SetFocus
End Sub

