VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "录制鼠标拖动及播放"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12225
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   12225
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8280
      TabIndex        =   8
      Text            =   "20"
      Top             =   157
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "使用画图工具 播放间隔:"
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   240
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Text            =   "20"
      Top             =   180
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "清空"
      Height          =   495
      Left            =   10200
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CheckBox chkRecord 
      BackColor       =   &H00C0FFC0&
      Caption         =   "录制鼠标"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtRecordContent 
      Appearance      =   0  'Flat
      Height          =   6135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   720
      Width           =   11895
   End
   Begin VB.Timer tmrRecord 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   1320
      Top             =   120
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "播放下面鼠标动作"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "毫秒"
      Height          =   180
      Left            =   8880
      TabIndex        =   9
      Top             =   277
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "毫秒"
      Height          =   180
      Left            =   3240
      TabIndex        =   6
      Top             =   277
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "抓捕间隔:"
      Height          =   180
      Left            =   1800
      TabIndex        =   5
      Top             =   277
      Width           =   810
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim wRecord As New clsWindow '录制用

Private Sub Form_Load()
    Dim w As New clsWindow
    w.hWnd = Me.hWnd
    w.SetTop '将本窗口设置为置顶
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtRecordContent.Width = Me.ScaleWidth - txtRecordContent.Left - 90
    txtRecordContent.Height = Me.ScaleHeight - txtRecordContent.Top - 90
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub chkRecord_Click()
    wRecord.Wait 500
    tmrRecord.Interval = Text1.Text
    tmrRecord.Enabled = (chkRecord.Value = 1)
    chkRecord.Caption = IIf(chkRecord.Value = 1, "录制中...", "录制鼠标")
    txtRecordContent.SetFocus
End Sub

Private Sub Command1_Click()
    txtRecordContent.Text = ""
End Sub
'执行鼠标拖动代码
Private Sub cmdDone_Click()
    Dim w As New clsWindow
    If Check1.Value = 1 Then '在画布上操作
        If w.GetWindowByPID(Shell("mspaint.exe", vbMaximizedFocus), 3).hWnd = 0 Then Exit Sub '未启动成功画图则退出
    End If
    
    w.Wait 1000
    w.DragToEx txtRecordContent.Text, , Text2.Text, Text2.Text
End Sub
'记录鼠标拖动
Private Sub tmrRecord_Timer()
'    Static status_LButton, isLButtonClicking As Boolean
'    Static status_RButton, isRButtonClicking As Boolean
'
    Static status_LButton, status_RButton, isButtonClicking As Boolean
    Static newPos$, oldPos$, isWrite As Boolean, strLineCode$
    status_LButton = GetAsyncKeyState(vbKeyLButton)
    status_RButton = GetAsyncKeyState(vbKeyRButton)
    If status_LButton < 0 Or status_RButton < 0 Then '参考：https://tieba.baidu.com/p/1829831956
        isButtonClicking = True
        If Not isWrite Then '表示写过开头了
            isWrite = True
            If status_LButton < 0 Then
                strLineCode = wRecord.GetCursorPoint & ":"
            ElseIf status_RButton < 0 Then
                strLineCode = "R" & wRecord.GetCursorPoint & ":" '区别右键按下
            End If
        End If
        
        newPos = wRecord.GetCursorPoint
        If newPos <> oldPos Then '当有变化的时候才记录
            strLineCode = strLineCode & newPos & ":"
            oldPos = newPos
        End If
    Else '未按下
        If isButtonClicking Then '如果当前正在记录中，出现了未按下的情况，那么就表示结束
            isButtonClicking = False
            isWrite = False
            strLineCode = Left(strLineCode, Len(strLineCode) - 1)
            If chkRecord.Value = 1 Then '防止点击停止录制的按钮也被录制
                txtRecordContent.SelStart = Len(txtRecordContent.Text)
                txtRecordContent.SelText = strLineCode & vbCrLf
            End If
            strLineCode = ""
        End If
    End If
    
    '左键拖动的处理
'    Static status_RButton, isRButtonClicking As Boolean
'    Static newPosR$, oldPosR$, isWriteR As Boolean, strLineCodeR$
'    status_RButton = GetAsyncKeyState(vbKeyRButton)
'    If status_RButton < 0 Then '参考：https://tieba.baidu.com/p/1829831956
'        isRButtonClicking = True
'        If Not isWriteR Then '表示写过开头了
'            isWriteR = True
'            strLineCodeR = "R" & wRecord.GetCursorPoint & ":"
'        End If
'
'        newPosR = wRecord.GetCursorPoint
'        If newPosR <> oldPosR Then '当有变化的时候才记录
'            strLineCodeR = strLineCodeR & newPosR & ":"
'            oldPosR = newPosR
'        End If
'    Else '未按下
'        If isRButtonClicking Then '如果当前正在记录中，出现了未按下的情况，那么就表示结束
'            isRButtonClicking = False
'            isWriteR = False
'            strLineCodeR = Left(strLineCodeR, Len(strLineCodeR) - 1)
'            If chkRecord.Value = 1 Then '防止点击停止录制的按钮也被录制
'                txtRecordContent.SelStart = Len(txtRecordContent.Text)
'                txtRecordContent.SelText = strLineCodeR & vbCrLf
'            End If
'            strLineCodeR = ""
'        End If
'    End If
End Sub
