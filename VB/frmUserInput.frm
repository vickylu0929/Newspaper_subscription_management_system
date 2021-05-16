VERSION 5.00
Begin VB.Form frmUserInput 
   Caption         =   "User information input"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   Picture         =   "frmUserInput.frx":0000
   ScaleHeight     =   5835
   ScaleWidth      =   12420
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "Input"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   3600
      Width           =   2415
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
      Left            =   840
      TabIndex        =   15
      Top             =   1440
      Width           =   975
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
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   14
      Top             =   2640
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
      Left            =   1080
      TabIndex        =   13
      Top             =   3840
      Width           =   735
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
      Left            =   6120
      TabIndex        =   12
      Top             =   1440
      Width           =   735
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
      Left            =   5400
      TabIndex        =   11
      Top             =   2640
      Width           =   1575
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
      Height          =   375
      Index           =   5
      Left            =   5760
      TabIndex        =   10
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User information input"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmUserInput"
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

conn.BeginTrans
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = conn
rs.CursorType = adOpenDynamic
rs.LockType = adLockPessimistic
rs.Open "select * from User_information"

rs.AddNew
rs(0) = Text1.Text
rs(1) = Text2.Text
rs(2) = Text3.Text
rs(3) = Text4.Text
rs(4) = Text5.Text
rs(5) = Text6.Text
rs.Update

If vbYes = MsgBox("Are you sure you want to enter information?", vbYesNo + vbQuestion, "System prompt") Then
conn.CommitTrans
MsgBox "Entered successfully!", vbOKCancel, "System prompt"
Else
conn.RollbackTrans
End If
rs.Close
conn.Close
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
frmUserInput.Hide
frmAdministratorMenu.Show
End Sub

Private Sub Form_Load()
Text1.TabIndex = 0
End Sub


