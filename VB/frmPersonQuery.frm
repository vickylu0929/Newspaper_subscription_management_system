VERSION 5.00
Begin VB.Form frmPersonQuery 
   Caption         =   "ersonal information query and modification"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12465
   LinkTopic       =   "Form2"
   Picture         =   "frmPersonQuery.frx":0000
   ScaleHeight     =   5835
   ScaleWidth      =   12465
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "Display"
      Height          =   495
      Left            =   10680
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   10680
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modification"
      Height          =   495
      Left            =   10680
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal information query and modification"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   15
      Top             =   240
      Width           =   8415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User name"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   14
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   13
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   12
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
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
      Index           =   3
      Left            =   6360
      TabIndex        =   11
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tele number"
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
      Index           =   4
      Left            =   5520
      TabIndex        =   10
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5880
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
   End
End
Attribute VB_Name = "frmPersonQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim username As String
Dim password As String
username = Text1.Text
password = Text2.Text
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection
conn.ConnectionString = "DSN=ttt;UID=yh;PWD=123"
conn.Open
Set rs = conn.Execute("select * from User_information")
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
frmPersonQuery.Hide
frmUser.Show
End Sub

Private Sub Command4_Click()
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Set conn = New ADODB.Connection
conn.ConnectionString = "DSN=ttt;UID=yh;PWD=123"
conn.CursorLocation = adUseServer
conn.Open

conn.BeginTrans
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = conn
rs.CursorType = adOpenDynamic
rs.LockType = adLockPessimistic
rs.Open "select * from User_information"

rs(0) = Text1.Text
rs(1) = Text2.Text
rs(2) = Text3.Text
rs(3) = Text4.Text
rs(4) = Text5.Text
rs(5) = Text6.Text
rs.Update

If vbYes = MsgBox("Are you sure you want to modify the information?", vbYesNo + vbQuestion, "System prompt") Then
conn.CommitTrans
MsgBox "Modified successfully!", vbOKCancel, "System prompt"
Else
conn.RollbackTrans
End If
rs.Close
conn.Close
End Sub

Private Sub Form_Load()
Text1.TabIndex = 0
End Sub

