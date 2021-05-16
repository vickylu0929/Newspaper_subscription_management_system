VERSION 5.00
Begin VB.Form frmNewspaperSubscription 
   Caption         =   "Newspaper subscription"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   Picture         =   "frmNewspaperSubscription.frx":0000
   ScaleHeight     =   5865
   ScaleWidth      =   12390
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command4 
      Caption         =   "Continue"
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Details"
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3480
      ItemData        =   "frmNewspaperSubscription.frx":2575B
      Left            =   7440
      List            =   "frmNewspaperSubscription.frx":2575D
      TabIndex        =   9
      Top             =   1200
      Width           =   3855
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      ItemData        =   "frmNewspaperSubscription.frx":2575F
      Left            =   2760
      List            =   "frmNewspaperSubscription.frx":25778
      TabIndex        =   8
      Top             =   4080
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frmNewspaperSubscription.frx":257B9
      Left            =   2760
      List            =   "frmNewspaperSubscription.frx":257DB
      TabIndex        =   7
      Top             =   2640
      Width           =   2055
   End
   Begin VB.ComboBox combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frmNewspaperSubscription.frx":257FE
      Left            =   2760
      List            =   "frmNewspaperSubscription.frx":2580B
      TabIndex        =   6
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Subscribe"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Time of subscription"
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
      Left            =   360
      TabIndex        =   5
      Top             =   4080
      Width           =   2295
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
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2640
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
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Newspaper subscription"
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
      TabIndex        =   2
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmNewspaperSubscription"
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

rs.AddNew
rs(0) = "2003"
rs(1) = "1511"
rs(2) = Combo2.Text
rs(3) = Combo3.Text
rs(4) = "40"
'rs(4) = Combo2.Text * Combo3.Text

rs.Update

If vbYes = MsgBox("Are you sure you want to subscribe to this newspaper?", vbYesNo + vbQuestion, "System prompt") Then
conn.CommitTrans
MsgBox "Subscription succeeded!", vbOKCancel, "System prompt"
Else
conn.RollbackTrans
End If
rs.Close
conn.Close
End Sub

Private Sub Command2_Click()
List1.Clear
If "People's Daily" = combo1.Text Then
List1.AddItem "Newspaper name：People's Daily"
List1.AddItem "Classification number：1001"
List1.AddItem "Classification name：political newspaper"
List1.AddItem "Newspaper code：1-1"
List1.AddItem "Publishing house：People's Publishing House"
List1.AddItem "Publication cycle：every day"
List1.AddItem "Monthly unit price：1"
Else
List1.Clear
If "IT Time magazine" = combo1.Text Then
List1.AddItem "Newspaper name：IT Time magazine"
List1.AddItem "Classification number：1002"
List1.AddItem "Classification name：business finance"
List1.AddItem "Newspaper code：1-2"
List1.AddItem "Publishing house：Science and Technology Press"
List1.AddItem "Publication cycle：half month"
List1.AddItem "Monthly unit price：10"
Else
List1.Clear
If "Vogue" = combo1.Text Then
List1.AddItem "Newspaper name：Vogue"
List1.AddItem "Classification number：1003"
List1.AddItem "Classification name：fashion magazine"
List1.AddItem "Newspaper code：1-3"
List1.AddItem "Publishing house：Literature and Art Publishing House"
List1.AddItem "Publication cycle：every month"
List1.AddItem "Monthly unit price：15"
End If
End If
End If

End Sub

Private Sub Command3_Click()
frmNewspaperSubscription.Hide
frmUser.Show
End Sub

Private Sub Command4_Click()
combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
List1.Clear

combo1.SetFocus
End Sub

Private Sub Form_Load()
combo1.TabIndex = 0
End Sub

