VERSION 5.00
Object = "{A2ED65B5-7DB0-4716-871E-53783A22D5D9}#66.0#0"; "JinDuT.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AirWar"
   ClientHeight    =   8130
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   12060
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   12060
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer BHB 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   10
      Left            =   11040
      Top             =   6600
   End
   Begin VB.Timer BHB 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   11520
      Top             =   6600
   End
   Begin VB.Timer HGC 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9720
      Top             =   7080
   End
   Begin VB.Timer ChangeLock 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10680
      Top             =   7080
   End
   Begin VB.Timer Spare 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11520
      Top             =   7080
   End
   Begin VB.Timer Ftime 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   6240
      Top             =   7560
   End
   Begin VB.Frame Frame1 
      Caption         =   "PLAYER2"
      Height          =   3135
      Index           =   1
      Left            =   9120
      TabIndex        =   15
      Top             =   0
      Width           =   2775
      Begin JinDuTiao.JinDuT JinDuT1 
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   16
         Top             =   945
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         ProgressColor   =   255
      End
      Begin JinDuTiao.JinDuT JinDuT2 
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   17
         Top             =   1305
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
      End
      Begin JinDuTiao.JinDuT JinDuT3 
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   18
         Top             =   1665
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         ProgressColor   =   33023
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         Height          =   405
         Index           =   1
         Left            =   120
         Top             =   2640
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "LV:"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "EXP:"
         Height          =   315
         Index           =   1
         Left            =   135
         TabIndex        =   24
         Top             =   2025
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CD:"
         Height          =   315
         Index           =   1
         Left            =   135
         TabIndex        =   23
         Top             =   1665
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "E:"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   1305
         Width           =   195
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "HP:"
         Height          =   315
         Index           =   1
         Left            =   135
         TabIndex        =   21
         Top             =   945
         Width           =   390
      End
      Begin VB.Label SkOn1 
         Alignment       =   2  'Center
         Caption         =   "盾"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   600
         TabIndex        =   20
         Top             =   2655
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label SkOn1 
         Alignment       =   2  'Center
         Caption         =   "核"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   2655
         Width           =   405
      End
      Begin VB.Label SkOn1 
         Alignment       =   2  'Center
         Caption         =   "血"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   7
         Left            =   1080
         TabIndex        =   28
         Top             =   2640
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label SkOn1 
         Alignment       =   2  'Center
         Caption         =   "黑"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   8
         Left            =   1560
         TabIndex        =   31
         Top             =   2640
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PLAYER1"
      Height          =   3135
      Index           =   0
      Left            =   6240
      TabIndex        =   4
      Top             =   0
      Width           =   2775
      Begin JinDuTiao.JinDuT JinDuT1 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   945
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         ProgressColor   =   255
      End
      Begin JinDuTiao.JinDuT JinDuT2 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   1305
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
      End
      Begin JinDuTiao.JinDuT JinDuT3 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   7
         Top             =   1665
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         ProgressColor   =   33023
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         Height          =   405
         Index           =   0
         Left            =   120
         Top             =   2640
         Width           =   405
      End
      Begin VB.Label SkOn1 
         Alignment       =   2  'Center
         Caption         =   "核"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   2655
         Width           =   405
      End
      Begin VB.Label SkOn1 
         Alignment       =   2  'Center
         Caption         =   "盾"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   600
         TabIndex        =   13
         Top             =   2655
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "HP:"
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   12
         Top             =   945
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "E:"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1305
         Width           =   195
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CD:"
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   1665
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "EXP:"
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   9
         Top             =   2025
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "LV:"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   435
      End
      Begin VB.Label SkOn1 
         Alignment       =   2  'Center
         Caption         =   "血"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   1080
         TabIndex        =   27
         Top             =   2640
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label SkOn1 
         Alignment       =   2  'Center
         Caption         =   "黑"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   1560
         TabIndex        =   30
         Top             =   2640
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.Timer Ftime 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   6720
      Top             =   7560
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   7560
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7920
      Top             =   7560
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8520
      Top             =   7560
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9120
      Top             =   7560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9720
      Top             =   7560
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8000
      Left            =   120
      ScaleHeight     =   8060.721
      ScaleMode       =   0  'User
      ScaleWidth      =   4990.596
      TabIndex        =   0
      Top             =   0
      Width           =   6000
   End
   Begin VB.Label Label10 
      Height          =   495
      Left            =   6240
      TabIndex        =   29
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "单机模式"
      Height          =   315
      Left            =   10800
      TabIndex        =   26
      Top             =   7680
      Width           =   960
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Timing:"
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Menu 菜单 
      Caption         =   "菜单"
      Begin VB.Menu 开始 
         Caption         =   "开始"
         Shortcut        =   ^K
      End
      Begin VB.Menu 重来 
         Caption         =   "重来"
         Shortcut        =   ^C
      End
      Begin VB.Menu 暂停 
         Caption         =   "暂停"
         Shortcut        =   ^Z
      End
      Begin VB.Menu 双人 
         Caption         =   "双人"
      End
   End
   Begin VB.Menu 网络 
      Caption         =   "网络"
      Begin VB.Menu 启动本地服务器 
         Caption         =   "启动本地服务器"
      End
      Begin VB.Menu 连接远程服务器 
         Caption         =   "连接远程服务器"
      End
   End
   Begin VB.Menu 设置 
      Caption         =   "设置"
      Begin VB.Menu 键位设置 
         Caption         =   "键位设置"
      End
      Begin VB.Menu 手柄模式 
         Caption         =   "手柄模式"
      End
   End
   Begin VB.Menu 帮助 
      Caption         =   "帮助"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, _
        lpList As Long) As Long
Private Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" _
        (ByVal pwszKLID As String) As Long
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Private Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal _
        hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal _
        flags As Long) As Long
Const IME_CONFIG_GENERAL = 1
Const KLF_REORDER = &H8
Const KLF_ACTIVATE = &H1
Dim la(1 To 16) As Long
Dim ActIme, BigFoePlaneTime, CunningFoePlaneTime, FrozenFoePlaneTime, BlackHolefoePlaneTime As Long
Private T_s As Single
Private TimeTired As Long
Dim myJoy As JOYINFOEX

Private Sub BHB_Timer(Index As Integer)
PSkillID_BHB(Index) = PSkillID_BHB(Index) + 300
End Sub

Private Sub ChangeLock_Timer()
TimeTired = TimeTired + 1
If TimeTired >= 1 Then PlaneWYKZ_Skill_SwitchLock = False: TimeTired = 0: ChangeLock.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Picture1_KeyDown KeyCode, Shift
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Picture1_KeyUp KeyCode, Shift
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub Ftime_Timer(Index As Integer)
PSkillID_Ft(Index) = PSkillID_Ft(Index) + 0.1
End Sub
Private Sub HGC_Timer()
Dim dX, dY, db As Long
Dim R&
myJoy.dwSize = 64
myJoy.dwFlags = JOY_RETURNALL
R& = joyGetPosEx(JOYSTICKID1, myJoy)
If R = 0 Then
    dX = myJoy.dwXpos     'As Long                '  x position
    dY = myJoy.dwYpos     'As Long                '  y position
    db = myJoy.dwButtons   ' As Long             '  button states
    If ReinCodeVis = True Then Label10.Caption = dX & ";" & dY & ";" & db
    Select Case dX
        Case Is < 20000
            Picture1_KeyDown PC(0).Left, 0
        Case Is > 40000
            Picture1_KeyDown PC(0).Right, 0
        Case Else
            Picture1_KeyUp PC(0).Left, 0
            Picture1_KeyUp PC(0).Right, 0
    End Select
    Select Case dY
        Case Is < 20000
            Picture1_KeyDown PC(0).Up, 0
        Case Is > 40000
            Picture1_KeyDown PC(0).Down, 0
        Case Else
            Picture1_KeyUp PC(0).Up, 0
            Picture1_KeyUp PC(0).Down, 0
    End Select
    Select Case db
        Case 1
            Picture1_KeyDown PC(0).Attack, 0
        Case 2
            Picture1_KeyDown PC(0).Ultimate_Skill, 0
        Case 4
            Picture1_KeyDown PC(0).Skill_Switch, 0
        Case 8
            World_Stop
        Case 64
            重来_Click
        Case 128
            World_Start
        Case Else
            Picture1_KeyUp PC(0).Attack, 0
            Picture1_KeyUp PC(0).Skill_Switch, 0
            Picture1_KeyUp PC(0).Ultimate_Skill, 0
    End Select
End If
End Sub
Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
'37_40:4865
Dim i As Long
If KeyCode = 192 Then
    If CMDShow = False Then CMDShow = True: World_Stop: Form2.Show Else CMDShow = False: World_Start: Unload Form2
End If
Select Case Local_State
    Case 0
        For i = 0 To 5
            If KCTemp(i) = KeyCode Then Exit Sub
        Next
        For i = 0 To 4
            If KCTemp(i) = 0 Then KCTemp(i) = KeyCode: Exit Sub
        Next
        KCTemp(5) = KeyCode
    Case 1
        If KeyCode_Filtration(KeyCode) = True Then Exit Sub
        For i = 0 To 2
            If KCTemp(i) = KeyCode Then Exit Sub
        Next
        For i = 0 To 1
            If KCTemp(i) = 0 Then KCTemp(i) = KeyCode: Exit Sub
        Next
        KCTemp(2) = KeyCode
    Case 2
        If KeyCode_Filtration(KeyCode) = True Then Exit Sub
        For i = 3 To 5
            If KCTemp(i) = KeyCode Then Exit Sub
        Next
        For i = 3 To 4
            If KCTemp(i) = 0 Then KCTemp(i) = KeyCode: Exit Sub
        Next
        KCTemp(5) = KeyCode
End Select
End Sub
Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Long
If KeyCode_Filtration(KeyCode) = True Then Exit Sub
Select Case Local_State
    Case 0
        For i = 0 To 5
            If KCTemp(i) = KeyCode Then KCTemp(i) = 0
        Next
    Case 1
        For i = 0 To 2
            If KCTemp(i) = KeyCode Then KCTemp(i) = 0
        Next
    Case 2
        For i = 3 To 5
            If KCTemp(i) = KeyCode Then KCTemp(i) = 0
        Next
End Select
End Sub
Private Sub Form_Load()
Dim R&
Dim hwnd&
Static TheX As Long
Static TheY As Long
' Tell form to receive joystick functions.
R& = joySetCapture(hwnd, JOYSTICKID1, 1, 0)
R& = joyReleaseCapture(JOYSTICKID1)
' Get joystick position coordinates and fill in the TheX and TheY
' variables.
R& = joyGetPosEx(JOYSTICKID1, myJoy)
TheX = myJoy.dwXpos
TheY = myJoy.dwYpos
'-----------------------------------------------------------------------------
Key_Config_Def
Picture1.Scale (0, Picture1.Height)-(Picture1.Width, 0)
World_Load
End Sub

Private Sub Spare_Timer()
Dim LVTemp: Dim i As Long
If Local_State = 2 Then
    For i = 0 To 1
        Form1.Shape1(i).Left = Form1.SkOn1(PSkillID(i) + i * 5).Left: Form1.Shape1(i).Top = Form1.SkOn1(PSkillID(i) + i * 5).Top - 20
    Next
End If
If Form1.SkOn1(1).Visible = True And Form1.SkOn1(6).Visible = True Then Exit Sub
For i = 0 To 1
    LVTemp = Split(Label9(i), ":")
    If Val(LVTemp(1)) > 2 Then
        Form1.SkOn1(1 + i * 5).Visible = True
    End If
    If Val(LVTemp(1)) > 4 Then
        Form1.SkOn1(2 + i * 5).Visible = True
    End If
Next
End Sub
Private Sub Timer1_Timer()
F5
End Sub
Private Sub Timer2_Timer()
Dim i, c As Long
If KeyboardVis = True Then Label3.Caption = KCTemp(0) & "," & KCTemp(1) & "," & KCTemp(2) & "," & KCTemp(3) & "," & KCTemp(4) & "," & KCTemp(5)
'---------------------------Test------------------------------
For i = 0 To 1
    If Pg(i).a = True Then
        PBCD(i) = True
        If PBSkillCD(i) = False Then PBSkillCDTime(i) = PBSkillCDTime(i) + 1
        If PBSkillCDTime(i) > 99 Then PBSkillCD(i) = True: PBSkillCDTime(i) = 0
        Label8(i).Caption = "EXP:" & Pg(i).EMP & "/" & Pg(i).MxEmp
        Label9(i).Caption = "LV:" & Pg(i).Rank
    Else
        JinDuT1(i).Progress = 0
    End If
    For c = 1 To 3
        If PSkill(i, c) = True Then SkOn1(c + i * 5).Visible = True Else SkOn1(c + i * 5).Visible = False
    Next
Next
'-------------------------------------------------------------
ActIme = "134481924" '这里是英文输入法的码。
ActivateKeyboardLayout ActIme, 1
End Sub
Private Sub Timer3_Timer()
If Pg(0).a = True Or Pg(1).a = True Then
    T_s = Format(T_s + 0.1, "0.0")
    Form1.Label2.Caption = T_s
    Diff = T_s
    If Diff > 300 Then Timer5.Interval = 500
    If Diff > 600 Then Timer5.Interval = 100
End If
End Sub
Private Sub Timer4_Timer()
Dim i As Long
For i = 0 To SgSum - 1
    If Sg(i).a = False Then
        Sg_Add (i): Exit Sub
    End If
Next
Sg_Add (SgSum)
SgSum = SgSum + 1
End Sub

Private Sub Timer5_Timer()
'---------------------------Code------------------------------
BigFoePlaneTime = BigFoePlaneTime + 1
CunningFoePlaneTime = CunningFoePlaneTime + 1
FrozenFoePlaneTime = FrozenFoePlaneTime + 1
BlackHolefoePlaneTime = BlackHolefoePlaneTime + 1
PCEsp_Recovery
If FoePlaneLifeSum() <= Diff / 10 Then
    Test_FoePlane
End If
If BigFoePlaneTime > 30 - Diff / 10 Then
    Test_FoePlane 1: BigFoePlaneTime = 0
End If
If CunningFoePlaneTime > 60 - Diff / 10 Then
    Test_FoePlane 2: CunningFoePlaneTime = 0
End If
If FrozenFoePlaneTime > 90 - Diff / 8 Then
    Test_FoePlane 3: FrozenFoePlaneTime = 0
End If
If BlackHolefoePlaneTime > 180 - Diff / 10 Then
    Test_FoePlane 4: BlackHolefoePlaneTime = 0
End If
Test_Bullet_2
End Sub
Private Sub 帮助_Click()

MsgBox "操作说明：" _
& vbCrLf & "方向键(上下左右):" _
& vbCrLf & "Player1: WSAD" _
& vbCrLf & "Player2:↑↓←→" _
& vbCrLf & "普通弹药: " _
& vbCrLf & "Player1: J" _
& vbCrLf & "Player2: 1" _
& vbCrLf & "特殊弹药: " _
& vbCrLf & "Player1: K" _
& vbCrLf & "Player2: 2" _
& vbCrLf & "特殊弹药切换:" _
& vbCrLf & "Player1: L" _
& vbCrLf & "Player2: 3"

End Sub

Private Sub 键位设置_Click()
Form5.Show
End Sub

Private Sub 开始_Click()
World_Start
End Sub

Private Sub 连接远程服务器_Click()
If Local_State = 0 Then Local_State = 2: 双人_Click: Form4.Show: Spare.Enabled = True Else MsgBox "请先关闭本地服务器再连接远程服务器！"
Local_State_Vision
End Sub

Private Sub 启动本地服务器_Click()
If Local_State = 0 Then Local_State = 1: 双人_Click: Form3.Show: World_Stop Else MsgBox "请先关闭远程服务器连接再启动本地服务器！"
Local_State_Vision
End Sub
Private Sub 手柄模式_Click()
If 手柄模式.Checked = False Then 手柄模式.Checked = True: HGC.Enabled = True Else 手柄模式.Checked = False: HGC.Enabled = False
End Sub
Private Sub 双人_Click()
If DuoPlayer = True Then Exit Sub
PC_2_Def
DuoPlayer = True
End Sub
Private Sub 暂停_Click()
World_Stop
End Sub

Private Sub 重来_Click()
BgSum = 0: T_s = 0: SgSum = 0
World_Load
End Sub
