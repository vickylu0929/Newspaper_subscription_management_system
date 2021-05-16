VERSION 5.00
Begin VB.Form frmUserQuery 
   Caption         =   "User information query and modification"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   Picture         =   "frmUserQuery.frx":0000
   ScaleHeight     =   5865
   ScaleWidth      =   12375
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command4 
      Caption         =   "Modification"
      Height          =   495
      Left            =   10440
      TabIndex        =   13
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   7080
      TabIndex        =   8
      Top             =   4680
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   7080
      TabIndex        =   7
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   7080
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   10440
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continue"
      Height          =   495
      Left            =   10440
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Query"
      Height          =   495
      Left            =   10440
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
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
      Left            =   6480
      TabIndex        =   16
      Top             =   2520
      Width           =   375
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
      Left            =   5640
      TabIndex        =   15
      Top             =   3720
      Width           =   1335
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
      Left            =   6000
      TabIndex        =   14
      Top             =   4800
      Width           =   1095
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
      Left            =   1560
      TabIndex        =   12
      Top             =   3720
      Width           =   735
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
      Left            =   1200
      TabIndex        =   11
      Top             =   2520
      Width           =   975
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
      Left            =   3960
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User information query and modification"
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
      TabIndex        =   9
      Top             =   360
      Width           =   7455
   End
End
Attribute VB_Name = "frmUserQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

Set conn = New ADODB.Connection
conn.ConnectionString = "DSN=ttt;UID=yh;PWD=123"
conn.CursorLocation = adUseServer
conn.Open
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = conn
rs.CursorType = adOpenDynamic
rs.LockType = adLockPessimistic
rs.Open "select * from User_information"
rs.MoveFirst
cond = "User_name='" & Text1 & "'"
rs.Find (cond)
If rs.EOF Then
MsgBox "There is no record", vbOKOnly, "System prompt"
For i = 1 To 5
Texti = ""
Next i
Else
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Text5.Text = rs(4)
Text6.Text = rs(5)
MsgBox "Query was successful!", vbOKCancel, "System prompt"
End If
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
frmUserQuery.Hide
frmAdministratorMenu.Show
End Sub

Private Sub Command4_Click()
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

Set conn = New ADODB.Connection
conn.ConnectionString = "DSN=ttt;UID=yh;PWD=123"
conn.CursorLocation = adUseServer
conn.Open
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = conn
rs.CursorType = adOpenDynamic
rs.LockType = adLockPessimistic
rs.Open "select * from User_information"
rs.MoveFirst
cond = "User_name='" & Text1 & "'"
rs.Find (cond)
If rs.EOF Then
MsgBox "There is no record", vbOKOnly, "System prompt"
For i = 1 To 5
Texti = ""
Next i
Else
rs(1) = Text2.Text
rs(2) = Text3.Text
rs(3) = Text4.Text
rs(4) = Text5.Text
rs(5) = Text6.Text
rs.Update
MsgBox "Modified successfully!", vbOKCancel, "System prompt"
End If
End Sub

Private Sub Form_Load()
Text1.TabIndex = 0
End Sub

