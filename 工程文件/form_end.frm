VERSION 5.00
Begin VB.Form form_end 
   Caption         =   "幽灵工厂—教你学编程"
   ClientHeight    =   6285
   ClientLeft      =   3840
   ClientTop       =   3375
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   Picture         =   "form_end.frx":0000
   ScaleHeight     =   6285
   ScaleWidth      =   13245
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1680
      Top             =   3960
   End
   Begin VB.CommandButton cmd_end 
      Caption         =   "完"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2040
      TabIndex        =   0
      Top             =   4440
      Width           =   8175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Now,after your help Our company has returned to its previous order and sincerely thank you!"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1695
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   8655
   End
End
Attribute VB_Name = "form_end"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub cmd_end_Click()
End
End Sub

Private Sub Form_Load()
i = 0
Timer1.Interval = 1000
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
cmd_end.Caption = cmd_end.Caption + "。"
i = i + 1
If i = 4 Then cmd_end.Caption = "完": i = 0
End Sub
