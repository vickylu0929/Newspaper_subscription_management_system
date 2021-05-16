VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   5820
   ScaleWidth      =   12075
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Administrator"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   5520
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5520
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Login name"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

Set conn = New ADODB.Connection
conn.ConnectionString = "DSN=ttt;UID=yh;PWD=123"
conn.Open

If Option2.Value = True Then

Set rs = conn.Execute("select * from User_information")

Do While Not rs.EOF

If Trim(rs.Fields(0)) = Trim(Text1.Text) And Trim(rs.Fields(1)) = Trim(Text2.Text) Then

frmLogin.Hide
frmUser.Show
Text1.Text = ""
Text2.Text = ""
Exit Do
End If
rs.MoveNext
Loop

Else

If Option1.Value = True Then

Set rs = conn.Execute("select * from Administrator_information")

Do While Not rs.EOF

If Trim(rs.Fields(0)) = Trim(Text1.Text) And Trim(rs.Fields(1)) = Trim(Text2.Text) Then

frmLogin.Hide
frmAdministratorMenu.Show
Exit Do
End If
rs.MoveNext
Loop
End If
End If
If rs.EOF Then MsgBox "The login name or password you entered is incorrect", vbOKCancel, "System prompt"

End Sub

Private Sub Command2_Click()
frmLogin.Hide
frmMain.Show
End Sub

Private Sub Form_Load()
Text1.TabIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Command2_Click
End Sub
