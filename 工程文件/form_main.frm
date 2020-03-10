VERSION 5.00
Begin VB.Form form_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "幽灵工厂—教你学编程"
   ClientHeight    =   8985
   ClientLeft      =   1905
   ClientTop       =   2205
   ClientWidth     =   16920
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "form_main.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   16920
   Begin VB.PictureBox pic_ghost 
      AutoSize        =   -1  'True
      Height          =   2130
      Left            =   4200
      Picture         =   "form_main.frx":6CD2D
      ScaleHeight     =   2070
      ScaleWidth      =   1560
      TabIndex        =   10
      Top             =   2760
      Width           =   1620
   End
   Begin VB.PictureBox pic_box3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   960
      Picture         =   "form_main.frx":70825
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   9
      Top             =   6720
      Width           =   1500
   End
   Begin VB.PictureBox pic_box2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   1080
      Picture         =   "form_main.frx":726A4
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   8
      Top             =   3360
      Width           =   1500
   End
   Begin VB.PictureBox pic_box1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   1080
      Picture         =   "form_main.frx":7450E
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   7
      Top             =   720
      Width           =   1500
   End
   Begin VB.CommandButton cmd_quit 
      Caption         =   "退出"
      Height          =   375
      Left            =   13560
      TabIndex        =   6
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmd_help 
      Caption         =   "帮助"
      Height          =   375
      Left            =   12240
      TabIndex        =   4
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Timer timer_pass 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   9840
      Top             =   4080
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6360
      Top             =   2400
   End
   Begin VB.Timer timer_go 
      Left            =   6480
      Top             =   4560
   End
   Begin VB.Timer timer_go2 
      Left            =   2880
      Top             =   4800
   End
   Begin VB.Timer timer_go3 
      Left            =   3240
      Top             =   7680
   End
   Begin VB.CommandButton cmd_recovery 
      Caption         =   "复位"
      Height          =   375
      Left            =   14880
      TabIndex        =   3
      Top             =   8520
      Width           =   2055
   End
   Begin VB.TextBox txt_line 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "form_main.frx":76154
      Top             =   360
      Width           =   375
   End
   Begin VB.Timer timer_get3 
      Left            =   1800
      Top             =   8040
   End
   Begin VB.Timer timer_get2 
      Left            =   1920
      Top             =   4920
   End
   Begin VB.Timer timer_put 
      Left            =   6840
      Top             =   4080
   End
   Begin VB.Timer timer_re 
      Left            =   6360
      Top             =   4080
   End
   Begin VB.Timer timer_go1 
      Left            =   3240
      Top             =   1920
   End
   Begin VB.Timer timer_get1 
      Left            =   1800
      Top             =   1560
   End
   Begin VB.TextBox txt_code 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   12720
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton cmd_down 
      BackColor       =   &H00FFFF00&
      Caption         =   "▽"
      Height          =   375
      Left            =   12360
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.PictureBox pic_tra 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   11895
      Left            =   6360
      Picture         =   "form_main.frx":76159
      ScaleHeight     =   11895
      ScaleWidth      =   4500
      TabIndex        =   11
      Top             =   -480
      Width           =   4500
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "return;"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13920
      TabIndex        =   16
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "put;"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15360
      TabIndex        =   17
      Top             =   1680
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "go;"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12720
      TabIndex        =   15
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "get(3);"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15240
      TabIndex        =   14
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "get(2);"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13800
      TabIndex        =   13
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "get(1);"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12360
      TabIndex        =   12
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label lbl_help 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   12240
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "form_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim box(1 To 3) As Integer
Dim a As Integer
Dim x As Integer

Private Sub cmd_help_Click()
     
    If lbl_help.Visible = False Then
        lbl_help.Visible = True
    Else
        lbl_help.Visible = False
    End If
End Sub

Private Sub cmd_quit_Click()
End
End Sub

Private Sub cmd_recovery_Click()
    pic_ghost.Left = 400: pic_ghost.Top = 350
    pic_tra.Left = 750: pic_tra.Top = 80
    pic_box1.Left = 100: pic_box1.Top = 100
    pic_box2.Left = 100: pic_box2.Top = 450
    pic_box3.Left = 100: pic_box3.Top = 800
    pic_box1.Visible = True
    pic_box2.Visible = True
    pic_box3.Visible = True
    box(1) = 1
    box(2) = 1
    box(3) = 1
    timer_go1.Enabled = False
    timer_go2.Enabled = False
    timer_go3.Enabled = False
    timer_re.Enabled = False
    timer_put.Enabled = False
    timer_get1.Enabled = False
    timer_get2.Enabled = False
    timer_get3.Enabled = False
End Sub

Private Sub cmd_submit_Click()

End Sub

Private Sub Form_Load()
    '定义初始量
    timer_go1.Enabled = False
    timer_go2.Enabled = False
    timer_go3.Enabled = False
    timer_re.Enabled = False
    timer_put.Enabled = False
    timer_get1.Enabled = False
    timer_get2.Enabled = False
    timer_get3.Enabled = False
    '定义移动频率
    timer_put.Interval = 20
    timer_get1.Interval = 20
    timer_get2.Interval = 20
    timer_get3.Interval = 20
    timer_re.Interval = 20
    timer_go1.Interval = 20
    timer_go2.Interval = 20
    timer_go3.Interval = 20
    
    form_main.Height = 9 / 16 * form_main.Width
    form_main.Scale (0, 0)-(1920, 1080)
    '定义各物品大小
    pic_tra.Height = 900: pic_tra.Width = 400
    '定义各物品位置
    pic_ghost.Left = 400: pic_ghost.Top = 350
    pic_tra.Left = 750: pic_tra.Top = 80
    pic_box1.Left = 100: pic_box1.Top = 100
    pic_box2.Left = 100: pic_box2.Top = 450
    pic_box3.Left = 100: pic_box3.Top = 800
    
    box(1) = 1: box(2) = 1: box(3) = 1
    lbl_help.Caption = "get(x);将幽灵传送到对应的x箱子前；go;将幽灵移动到传送带附近；put;将幽灵手上的货物抛下；return;命令幽灵回到原位置"
    
End Sub


Private Sub Label1_Click()
txt_code.Text = "get(1);"
End Sub

Private Sub Label2_Click()
txt_code.Text = "get(2);"
End Sub

Private Sub Label3_Click()
txt_code.Text = "get(3);"
End Sub

Private Sub Label4_Click()
txt_code.Text = "go;"
End Sub

Private Sub Label5_Click()
txt_code.Text = "return;"
End Sub

Private Sub Label6_Click()
txt_code.Text = "put;"
End Sub

Private Sub timer_pass_Timer()
If pic_box1.Left > 500 And pic_box1.Visible = True Then
        pic_box1.Top = pic_box1.Top + 3
            If pic_box1.Top >= 700 Then
                box(1) = 0
                pic_box1.Visible = False
                MsgBox "真棒，那么接下来该回去了，试试return;命令吧！"
                timer_pass.Enabled = False
            End If
ElseIf pic_box2.Left > 500 And pic_box2.Visible = True Then
        pic_box2.Top = pic_box2.Top + 3
            If pic_box2.Top >= 700 Then
                box(2) = 0
                pic_box2.Visible = False
                timer_pass.Enabled = False
            End If
ElseIf pic_box3.Left > 500 And pic_box3.Visible = True Then
        pic_box3.Top = pic_box3.Top + 3
            If pic_box3.Top >= 700 Then
                box(3) = 0
                pic_box3.Visible = False
                timer_pass.Enabled = False
            End If
End If
            If pic_box1.Visible = False And pic_box2.Visible = False And pic_box3.Visible = False Then
                    MsgBox "perfect!"
                    form_main.Hide
                    form_end.Show
            End If
End Sub

Private Sub Timer1_Timer()
MsgBox "第一步，先试着控制幽灵往第一个盒子的方向运动，你可以输入“get(1);”(必须是英文符号)命令来控制机器人走到盒子1的位置，快来试试吧！"
Timer1.Enabled = False
End Sub
Sub txt_code_change()
If txt_code.Text = "get(1);" Then txt_code.Text = "": timer_get1.Enabled = True: txt_code.Enabled = False
If txt_code.Text = "get(2);" Then txt_code.Text = "": timer_get2.Enabled = True: txt_code.Enabled = False
If txt_code.Text = "get(3);" Then txt_code.Text = "": timer_get3.Enabled = True: txt_code.Enabled = False
If txt_code.Text = "go;" And pic_ghost.Top <= 200 Then txt_code.Text = "": timer_go1.Enabled = True: txt_code.Enabled = False
If txt_code.Text = "go;" And pic_ghost.Top <= 500 Then txt_code.Text = "": timer_go2.Enabled = True: txt_code.Enabled = False
If txt_code.Text = "go;" And pic_ghost.Top <= 1000 Then txt_code.Text = "": timer_go3.Enabled = True: txt_code.Enabled = False
If txt_code.Text = "go;" And pic_ghost.Top = 350 Then txt_code.Text = "": timer_go.Enabled = True: txt_code.Enabled = False
If txt_code.Text = "put;" Then txt_code.Text = "": timer_put.Enabled = True: txt_code.Enabled = False
If txt_code.Text = "return;" Then txt_code.Text = "": timer_re.Enabled = True: txt_code.Enabled = False
End Sub
Private Sub cmd_down_Click()
        If txt_code.Visible = True Then
            txt_code.Visible = False
            txt_line.Visible = False
            cmd_down.Caption = "△"
        Else
            txt_code.Visible = True
            txt_line.Visible = True
            cmd_down.Caption = "▽"
        End If
End Sub
Private Sub timer_get1_Timer()
        pic_ghost.Left = pic_ghost.Left - 4: pic_ghost.Top = pic_ghost.Top - 5
    If pic_ghost.Left <= 250 And pic_ghost.Top <= 300 Then
    If box(1) = 0 Then MsgBox "咦，箱子呢？": timer_re.Enabled = True: timer_get1.Enabled = False
        timer_get1.Enabled = False
        MsgBox "真聪明，接下来你可以用go;命令来控制幽灵将盒子运动到传送带上！"
        box(1) = 0
    ElseIf pic_ghost.Left >= 600 Then
        txt_code.Enabled = True
        MsgBox "请先return到原位置（命令return;）!"
        timer_get1.Enabled = False
    End If
End Sub

Private Sub timer_get2_Timer()
        pic_ghost.Left = pic_ghost.Left - 3.2
    If pic_ghost.Left <= 250 Then
    If box(2) = 0 Then MsgBox "咦，箱子呢？": timer_re.Enabled = True: timer_get2.Enabled = False
    txt_code.Enabled = True
    timer_get2.Enabled = False
    End If
End Sub

Private Sub timer_get3_Timer()
    pic_ghost.Left = pic_ghost.Left - 3: pic_ghost.Top = pic_ghost.Top + 9
        If pic_ghost.Left <= 250 Then
        If box(3) = 0 Then MsgBox "咦，箱子呢？": timer_re.Enabled = True: timer_get3.Enabled = False
            txt_code.Enabled = True
            timer_get3.Enabled = False
        ElseIf pic_ghost.Left >= 600 Then
            MsgBox "请先return到原位置（命令return;）!"
            txt_code.Enabled = True
            timer_get3.Enabled = False
        End If
End Sub
Private Sub timer_put_Timer()
If pic_box1.Left >= 300 And pic_box1.Visible = True Then
    pic_box1.Left = pic_box1.Left + 10
        If pic_box1.Left >= 800 Then
            timer_pass.Enabled = True
            txt_code.Enabled = True
            timer_put.Enabled = False
        End If
ElseIf pic_box2.Left >= 300 And pic_box2.Visible = True Then
    pic_box2.Left = pic_box2.Left + 10
        If pic_box2.Left >= 800 Then
            timer_pass.Enabled = True
            txt_code.Enabled = True
            timer_put.Enabled = False
        End If
ElseIf pic_box3.Left >= 300 And pic_box3.Visible = True Then
    pic_box3.Left = pic_box3.Left + 10
        If pic_box3.Left >= 800 Then
            timer_pass.Enabled = True
            txt_code.Enabled = True
            timer_put.Enabled = False
        End If
End If
End Sub

Private Sub timer_go1_Timer()
    If pic_ghost.Top <= 400 Then
        pic_box1.Left = pic_box1.Left + 8: pic_box1.Top = pic_box1.Top + 4.2
        pic_ghost.Left = pic_ghost.Left + 8: pic_ghost.Top = pic_ghost.Top + 4.2
            If pic_ghost.Left >= 550 Then
            MsgBox "接下来你可以用put;命令来控制幽灵抛货！"
                txt_code.Enabled = True
                timer_go1.Enabled = False
            End If
    End If

End Sub
Private Sub timer_go2_Timer()
    If pic_ghost.Top <= 400 Then
        pic_ghost.Left = pic_ghost.Left + 7
        pic_box2.Left = pic_box2.Left + 7
            If pic_ghost.Left >= 650 Then txt_code.Enabled = True: timer_go2.Enabled = False
    End If
End Sub
Private Sub timer_go3_Timer()
    If pic_ghost.Top <= 850 Then
        pic_ghost.Left = pic_ghost.Left + 32 / 7: pic_ghost.Top = pic_ghost.Top - 4.2
        pic_box3.Left = pic_box3.Left + 32 / 7: pic_box3.Top = pic_box3.Top - 4.2
            If pic_ghost.Left >= 650 Then txt_code.Enabled = True: timer_go3.Enabled = False
    End If
End Sub
Private Sub timer_go_Timer()
    If pic_ghost.Top = 350 Then
        pic_ghost.Left = pic_ghost.Left + 32 / 7
            If pic_ghost.Left >= 650 Then txt_code.Enabled = True: timer_go.Enabled = False
    End If
End Sub
Private Sub timer_re_Timer()
    If pic_ghost.Left <= 950 And pic_ghost.Left >= 250 Then
        pic_ghost.Left = pic_ghost.Left - 4
            If pic_ghost.Left < 400 Then txt_code.Enabled = True: timer_re.Enabled = False
    ElseIf pic_ghost.Left <= 250 Then
    pic_ghost.Left = 400: pic_ghost.Top = 350: txt_code.Enabled = True
    End If
End Sub


