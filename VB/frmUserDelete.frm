VERSION 5.00
Begin VB.Form frmUserDelete 
   Caption         =   "User information deletion"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   Picture         =   "frmUserDelete.frx":0000
   ScaleHeight     =   5880
   ScaleWidth      =   12405
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continue"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7560
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User information deletion"
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
      Left            =   3600
      TabIndex        =   5
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the user name to delete:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1680
      Width           =   3975
   End
End
Attribute VB_Name = "frmUserDelete"
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
rs(0) = ""
rs(1) = ""
rs(2) = ""
rs(3) = ""
rs(4) = ""
rs(5) = ""
rs.Update
MsgBox "Successfully deleted!", vbOKCancel, "System prompt"
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
frmUserDelete.Hide
frmAdministratorMenu.Show
End Sub

Private Sub Form_Load()
Text1.TabIndex = 0
End Sub
