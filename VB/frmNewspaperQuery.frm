VERSION 5.00
Begin VB.Form frmNewspaperQuery 
   Caption         =   "报刊查询及修改"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   Picture         =   "frmNewspaperQuery.frx":0000
   ScaleHeight     =   5880
   ScaleWidth      =   12960
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   7680
      TabIndex        =   17
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modification"
      Height          =   495
      Left            =   11160
      TabIndex        =   16
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Query"
      Height          =   495
      Left            =   11160
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continue"
      Height          =   495
      Left            =   11160
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   11160
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2400
      TabIndex        =   4
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   7680
      TabIndex        =   0
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Classification name"
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
      Left            =   5400
      TabIndex        =   18
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Newspaper query and modification"
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
      Left            =   3480
      TabIndex        =   15
      Top             =   240
      Width           =   6615
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
      Left            =   3720
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
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
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   2535
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
      Left            =   480
      TabIndex        =   12
      Top             =   3480
      Width           =   1935
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
      Left            =   240
      TabIndex        =   11
      Top             =   4440
      Width           =   2175
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
      Left            =   5640
      TabIndex        =   10
      Top             =   2400
      Width           =   2295
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
      Left            =   5520
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
   End
End
Attribute VB_Name = "frmNewspaperQuery"
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

rs.Open "select * from Newspaper_information"
rs.MoveFirst
cond = "Newspaper_code='" & Text1 & "'"
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
rs.Close

rs.Open "select * from Newspaper_category_information"
rs.MoveFirst
cond = "Classification_number='" & Text2 & "'"
rs.Find (cond)
If rs.EOF Then
MsgBox "There is no record", vbOKOnly, "System prompt"
For i = 1 To 5
Texti = ""
Next i
Else
Text7.Text = rs(1)

MsgBox "Query was successful!", vbOKCancel, "System prompt"
End If
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
frmNewspaperQuery.Hide
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
rs.Open "select * from Newspaper_information"
rs.MoveFirst
cond = "Newspaper_code='" & Text1 & "'"
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
rs.Close

rs.Open "select * from Newspaper_category_information"
rs.MoveFirst
cond = "Classification_number='" & Text2 & "'"
rs.Find (cond)
If rs.EOF Then
MsgBox "There is no record", vbOKOnly, "System prompt"
For i = 1 To 5
Texti = ""
Next i
Else
rs(1) = Text7.Text
rs.Update

MsgBox "Modified successfully!", vbOKCancel, "System prompt"
End If
End If

End Sub

Private Sub Form_Load()
Text1.TabIndex = 0
End Sub

