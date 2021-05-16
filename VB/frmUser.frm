VERSION 5.00
Begin VB.Form frmUser 
   Caption         =   "User"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   Picture         =   "frmUser.frx":0000
   ScaleHeight     =   5595
   ScaleWidth      =   12180
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User menu"
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
      Left            =   4320
      TabIndex        =   0
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Menu dy 
      Caption         =   "Data subscription"
      Begin VB.Menu bk 
         Caption         =   "Newspaper subscription"
      End
   End
   Begin VB.Menu cx 
      Caption         =   "Data query and modification"
      Begin VB.Menu yd 
         Caption         =   "Subscribed newspaper"
      End
      Begin VB.Menu gx 
         Caption         =   "Personal information"
      End
   End
   Begin VB.Menu tc 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bk_Click()
frmUser.Hide
frmNewspaperSubscription.Show
End Sub

Private Sub gx_Click()
frmUser.Hide
frmPersonQuery.Show
End Sub

Private Sub tc_Click()
frmUser.Hide
frmLogin.Show
End Sub

Private Sub yd_Click()
frmUser.Hide
frmSubscribedNewspaper.Show
End Sub
