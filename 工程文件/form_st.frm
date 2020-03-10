VERSION 5.00
Begin VB.Form form_st 
   Caption         =   "幽灵工厂—教你学编程"
   ClientHeight    =   7620
   ClientLeft      =   3990
   ClientTop       =   2790
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   13530
   Begin VB.CommandButton cmd_ent 
      Caption         =   "enter"
      Height          =   975
      Left            =   5520
      TabIndex        =   1
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label lblst 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   $"form_st.frx":0000
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13455
   End
End
Attribute VB_Name = "form_st"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ent_Click()
form_st.Hide
form_begin.Show
End Sub
