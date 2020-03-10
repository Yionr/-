VERSION 5.00
Begin VB.Form form_begin 
   Caption         =   "幽灵工厂—教你学编程"
   ClientHeight    =   8205
   ClientLeft      =   3840
   ClientTop       =   2820
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   Picture         =   "form_begin.frx":0000
   ScaleHeight     =   8205
   ScaleWidth      =   14565
   Begin VB.PictureBox pic_robot_main 
      AutoSize        =   -1  'True
      Height          =   2130
      Left            =   5520
      Picture         =   "form_begin.frx":52E78
      ScaleHeight     =   2070
      ScaleWidth      =   1560
      TabIndex        =   8
      Top             =   720
      Width           =   1620
   End
   Begin VB.PictureBox pic_robot2 
      AutoSize        =   -1  'True
      Height          =   2130
      Left            =   7800
      Picture         =   "form_begin.frx":56970
      ScaleHeight     =   2070
      ScaleWidth      =   1560
      TabIndex        =   7
      Top             =   6000
      Width           =   1620
   End
   Begin VB.PictureBox pic_robot3 
      AutoSize        =   -1  'True
      Height          =   2130
      Left            =   8520
      Picture         =   "form_begin.frx":5A468
      ScaleHeight     =   2070
      ScaleWidth      =   1560
      TabIndex        =   6
      Top             =   4200
      Width           =   1620
   End
   Begin VB.PictureBox pic_robot1 
      AutoSize        =   -1  'True
      Height          =   2130
      Left            =   5520
      Picture         =   "form_begin.frx":5DF60
      ScaleHeight     =   2070
      ScaleWidth      =   1560
      TabIndex        =   5
      Top             =   5160
      Width           =   1620
   End
   Begin VB.Timer timer_a2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6720
      Top             =   3840
   End
   Begin VB.Timer timer_mo3 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4320
      Top             =   4320
   End
   Begin VB.Timer timer_mo2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6600
      Top             =   4800
   End
   Begin VB.Timer timer_a 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5760
      Top             =   3960
   End
   Begin VB.Timer timer_mo1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   7320
      Top             =   4440
   End
   Begin VB.CommandButton cmd_give 
      Caption         =   "给票"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer timer_move 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6720
      Top             =   3840
   End
   Begin VB.CommandButton cmd_start 
      Caption         =   "开始游戏"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2760
      TabIndex        =   1
      Top             =   4440
      Width           =   6135
   End
   Begin VB.CommandButton cmd_work 
      Caption         =   "工作台"
      Height          =   735
      Left            =   4560
      TabIndex        =   0
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3795
      Left            =   9240
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.Label lbl_speak 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   7200
      Visible         =   0   'False
      Width           =   8655
   End
End
Attribute VB_Name = "form_begin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_give_Click()
Dim a As String
    a = InputBox("给票")
    If pic_robot1.Left <= 600 And pic_robot1.Left >= 200 Then
        If a = "A" Or a = "a" Then
            timer_mo1.Enabled = True
        Else
            MsgBox "输入错误！"
        End If
    ElseIf pic_robot2.Left <= 600 And pic_robot2.Left >= 200 Then
        Print a
        If a = "B" Or a = "b" Then
        timer_mo2.Enabled = True
        Else
            MsgBox "输入错误！"
        End If
    ElseIf pic_robot3.Left <= 600 Then
        If a = "C" Or a = "c" Then
        timer_mo3.Enabled = True
        Else
            MsgBox "输入错误！"
        End If
    End If
    a = ""
End Sub

Private Sub cmd_start_Click()
cmd_start.Visible = False
timer_move.Enabled = True
Label1.Visible = True
Label1.Caption = "(点击给票)规则：身高在1.8米以上的给a票（即在给票内输入a或A）；身高在1.8米以下的，如果体重超过70KG给B票；否则给C票。"
End Sub

Private Sub Form_Load()
    form_begin.Height = 9 / 16 * form_begin.Width
    form_begin.Scale (0, 0)-(1440, 900)
pic_robot1.Left = 1100: pic_robot1.Top = 500
pic_robot2.Left = 1250: pic_robot2.Top = 500
pic_robot3.Left = 1400: pic_robot3.Top = 500
pic_robot_main.Left = 510: pic_robot_main.Top = 100
End Sub

Private Sub timer_a_Timer()
pic_robot1.Left = pic_robot1.Left - 10
If pic_robot1.Left < 50 Then pic_robot1.Visible = False: timer_a.Enabled = False
End Sub

Private Sub timer_a2_Timer()
    pic_robot2.Left = pic_robot2.Left - 10
        If pic_robot2.Left < 50 Then pic_robot2.Visible = False: timer_a2.Enabled = False
End Sub

Private Sub timer_mo1_Timer()
    pic_robot1.Left = pic_robot1.Left - 5
    pic_robot2.Left = pic_robot2.Left - 5
    pic_robot3.Left = pic_robot3.Left - 5
    If pic_robot2.Left <= 600 Then
        timer_mo1.Enabled = False
        timer_a.Enabled = True
        lbl_speak.Caption = "我身高1米5，体重80KG"
    End If
End Sub

Private Sub timer_mo2_Timer()
    pic_robot2.Left = pic_robot2.Left - 5
    pic_robot3.Left = pic_robot3.Left - 5
    If pic_robot3.Left <= 600 Then
        lbl_speak.Caption = "我身高1米5，体重50KG"
        timer_mo2.Enabled = False
        timer_a2.Enabled = True
    End If
End Sub

Private Sub timer_mo3_Timer()
        pic_robot3.Left = pic_robot3.Left - 10
    If pic_robot3.Left <= 50 Then
        pic_robot3.Visible = False
        MsgBox "全员入内成功！"
        form_begin.Hide
        form_main.Show
        timer_mo3.Enabled = False
    End If
End Sub

Private Sub timer_move_Timer()
If pic_robot1.Left >= 600 Then
    pic_robot1.Left = pic_robot1.Left - 5
    pic_robot2.Left = pic_robot2.Left - 5
    pic_robot3.Left = pic_robot3.Left - 5
ElseIf pic_robot1.Left <= 600 Then
    cmd_give.Visible = True
    lbl_speak.Visible = True
    lbl_speak.Caption = "我身高1米8，体重100KG"
    timer_move.Enabled = False
End If
End Sub

Private Sub Timer1_Timer()

End Sub

