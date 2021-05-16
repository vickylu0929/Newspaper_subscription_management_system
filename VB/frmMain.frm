VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Newspaper subscription management system"
   ClientHeight    =   5865
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   5865
   ScaleWidth      =   12120
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   11520
      Top             =   5280
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Newspaper subscription management system"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   720
      TabIndex        =   0
      Top             =   2160
      Width           =   11055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
frmMain.Hide
frmLogin.Show
End Sub
