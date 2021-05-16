VERSION 5.00
Begin VB.Form frmSubscriptionDelete 
   Caption         =   "Subscription information deletion"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   Picture         =   "frmSubscriptionDelete.frx":0000
   ScaleHeight     =   5835
   ScaleWidth      =   12390
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continue"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7560
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Subscription information deletion"
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
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the order number to delete："
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
      TabIndex        =   4
      Top             =   1680
      Width           =   4455
   End
End
Attribute VB_Name = "frmSubscriptionDelete"
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
rs.Open "select * from Order_information"

If rs(0) = Text1.Text Then
rs(0) = ""
rs(1) = ""
rs(2) = Val("")
rs(3) = ""
rs(4) = Val("")
rs.Update

If vbYes = MsgBox("Are you sure you want to delete the information?", vbYesNo + vbQuestion, "System prompt") Then
conn.CommitTrans
MsgBox "Successfully deleted!", vbOKCancel, "System prompt"
Else
conn.RollbackTrans
End If
rs.Close
conn.Close

Set conn = New ADODB.Connection
conn.ConnectionString = "DSN=ttt;UID=yh;PWD=123"
conn.CursorLocation = adUseServer
conn.Open

conn.BeginTrans
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = conn
rs.CursorType = adOpenDynamic
rs.LockType = adLockPessimistic
rs.Open "select * from Order_information"

If rs(1) = Text1.Text Then
rs(0) = ""
rs(1) = ""
rs.Update
conn.CommitTrans
rs.Close
conn.Close
End If
End If

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
frmSubscriptionDelete.Hide
frmAdministratorMenu.Show
End Sub

Private Sub Form_Load()
Text1.TabIndex = 0
End Sub
