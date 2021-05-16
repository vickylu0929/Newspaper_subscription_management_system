VERSION 5.00
Begin VB.Form frmSubscribedNewspaper 
   Caption         =   "Subscribed newspaper"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   Picture         =   "frmSubscribedNewspaper.frx":0000
   ScaleHeight     =   5880
   ScaleWidth      =   12465
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "Display"
      Height          =   495
      Left            =   3360
      TabIndex        =   19
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   8520
      TabIndex        =   18
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   495
      Left            =   5520
      TabIndex        =   14
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   8520
      TabIndex        =   6
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   8520
      TabIndex        =   5
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   8520
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2760
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Return"
      Height          =   495
      Left            =   7680
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of subscriptions"
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
      Index           =   7
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total price"
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
      Index           =   6
      Left            =   7080
      TabIndex        =   15
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly unit price"
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
      Left            =   6240
      TabIndex        =   13
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Publication cycle"
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
      Left            =   6360
      TabIndex        =   12
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Publishing house"
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
      TabIndex        =   11
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Newspaper name"
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
      Left            =   960
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Classification number"
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
      Left            =   360
      TabIndex        =   9
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Newspaper code"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Subscribed newspaper"
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
      Left            =   3840
      TabIndex        =   7
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmSubscribedNewspaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'Dim username As String
'Dim password As String
'username = Text1.Text
'password = Text2.Text
'Dim conn As ADODB.Connection
'Dim rs As ADODB.Recordset

'Set rs = New ADODB.Recordset
'Set conn = New ADODB.Connection
'conn.ConnectionString = "DSN=ttt;UID=yh;PWD=123"
'conn.Open
'Set rs = conn.Execute("select * from Newspaper_information")

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text1.SetFocus
End Sub

Private Sub Command2_Click()

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection
conn.ConnectionString = "DSN=ttt;UID=yh;PWD=123"
conn.Open
Set rs = conn.Execute("select * from Newspaper_information")

Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)

Set rs = conn.Execute("select * from Order_information")

Text7.Text = rs.Fields(2)
Text8.Text = rs.Fields(4)
End Sub

Private Sub Command3_Click()
frmSubscribedNewspaper.Hide
frmUser.Show
End Sub

Private Sub Form_Load()
Text1.TabIndex = 0
End Sub


