VERSION 5.00
Begin VB.Form frmNewspaperDelete 
   Caption         =   "Newspaper information deletion"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   Picture         =   "frmNewspaperDelete.frx":0000
   ScaleHeight     =   5850
   ScaleWidth      =   12420
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continue"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the category number to delete："
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
      Left            =   840
      TabIndex        =   7
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the newspaper code to delete："
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
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Newspaper information deletion"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   4
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "frmNewspaperDelete"
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
rs.Open "select * from Newspaper_information"

If rs(0) = Text1.Text Then
rs(0) = ""
rs(1) = ""
rs(2) = ""
rs(3) = ""
rs(4) = ""
rs(5) = Val("")
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
rs.Open "select * from Newspaper_category_information"

If rs(0) = Text2.Text Then
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
Text2.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
frmNewspaperDelete.Hide
frmAdministratorMenu.Show
End Sub

Private Sub Form_Load()
Text1.TabIndex = 0
End Sub

