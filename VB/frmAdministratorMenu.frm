VERSION 5.00
Begin VB.Form frmAdministratorMenu 
   Caption         =   "Administrator menu"
   ClientHeight    =   5580
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   12165
   LinkTopic       =   "Form2"
   Picture         =   "frmAdministratorMenu.frx":0000
   ScaleHeight     =   5580
   ScaleWidth      =   12165
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator menu"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      TabIndex        =   0
      Top             =   1920
      Width           =   7695
   End
   Begin VB.Menu input 
      Caption         =   "Data input"
      Begin VB.Menu gly 
         Caption         =   "Administrator information"
         Index           =   0
      End
      Begin VB.Menu yh 
         Caption         =   "User information"
         Index           =   0
      End
      Begin VB.Menu bk 
         Caption         =   "Newspaper"
         Index           =   0
      End
   End
   Begin VB.Menu cx 
      Caption         =   "Data query and modification"
      Begin VB.Menu gly2 
         Caption         =   "Administrator information"
         Index           =   0
      End
      Begin VB.Menu yh2 
         Caption         =   "User information"
         Index           =   0
      End
      Begin VB.Menu bk2 
         Caption         =   "Newspaper"
         Index           =   0
      End
   End
   Begin VB.Menu sc 
      Caption         =   "Data deletion"
      Begin VB.Menu gly3 
         Caption         =   "Administrator information"
         Index           =   0
      End
      Begin VB.Menu yh3 
         Caption         =   "User information"
         Index           =   0
      End
      Begin VB.Menu bk3 
         Caption         =   "Newspaper"
         Index           =   0
      End
      Begin VB.Menu dx 
         Caption         =   "Subscription information"
      End
   End
   Begin VB.Menu tc 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmAdministratorMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bf_Click()
'For Each tdf In backup_db.TableDefs
'If (tdf.Attributes And dbSystemObject) = 0 Then
'End If
'Next
MsgBox "Database backup succeeded!", vbOKCancel, "System prompt"
End Sub

Private Sub bk_Click(Index As Integer)
frmAdministratorMenu.Hide
frmNewspaperInput.Show
End Sub

Private Sub bk2_Click(Index As Integer)
frmAdministratorMenu.Hide
frmNewspaperQuery.Show
End Sub

Private Sub bk3_Click(Index As Integer)
frmAdministratorMenu.Hide
frmNewspaperDelete.Show
End Sub


Private Sub dx_Click()
frmAdministratorMenu.Hide
frmSubscriptionDelete.Show
End Sub

Private Sub dy_Click()
frmAdministratorMenu.Hide
frmSubscriptionStatistics.Show
End Sub

Private Sub gly_Click(Index As Integer)
frmAdministratorMenu.Hide
frmAdministratorInput.Show
End Sub



Private Sub gly2_Click(Index As Integer)
frmAdministratorMenu.Hide
frmAdministratorQuery.Show
End Sub

Private Sub gly3_Click(Index As Integer)
frmAdministratorMenu.Hide
frmAdministratorDelete.Show
End Sub

Private Sub tc_Click()
frmAdministratorMenu.Hide
frmMain.Show
End Sub

Private Sub xw_Click()
'For Each tdf In backup_db.TableDefs
'If (tdf.Attributes And dbSystemObject) = 0 Then
'End If
'Next
MsgBox "Successful information maintenance!", vbOKCancel, "System prompt"
End Sub

Private Sub yh_Click(Index As Integer)
frmAdministratorMenu.Hide
frmUserInput.Show
End Sub

Private Sub yh2_Click(Index As Integer)
frmAdministratorMenu.Hide
frmUserQuery.Show
End Sub

Private Sub yh3_Click(Index As Integer)
frmAdministratorMenu.Hide
frmUserDelete.Show
End Sub


