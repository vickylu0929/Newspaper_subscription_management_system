VERSION 5.00
Begin VB.Form frmAdministratorDelete 
   Caption         =   "Administrator information deletion"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   Picture         =   "frmAdministratorDelete.frx":0000
   ScaleHeight     =   5865
   ScaleWidth      =   12435
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continue"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the name of the administrator to delete："
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
      Left            =   3120
      TabIndex        =   1
      Top             =   1680
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator information deletion"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
End
Attribute VB_Name = "frmAdministratorDelete"
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
rs.Open "select * from Administrator_information"
rs.MoveFirst
cond = "Administrator_name='" & Text1 & "'"
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
frmAdministratorDelete.Hide
frmAdministratorMenu.Show
End Sub

Private Sub Form_Load()
Text1.TabIndex = 0
End Sub
