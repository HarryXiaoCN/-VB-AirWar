VERSION 5.00
Object = "{A2ED65B5-7DB0-4716-871E-53783A22D5D9}#66.0#0"; "JinDuT.ocx"
Begin VB.Form Form1 
   Caption         =   "AirWar"
   ClientHeight    =   8130
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8805
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
   ScaleHeight     =   8130
   ScaleWidth      =   8805
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Ftime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7560
      Top             =   7080
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8160
      Top             =   7080
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6360
      Top             =   7560
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6960
      Top             =   7560
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7560
      Top             =   7560
   End
   Begin JinDuTiao.JinDuT JinDuT1 
      Height          =   255
      Left            =   6720
      TabIndex        =   3
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      ProgressColor   =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8160
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
      TabIndex        =   2
      Top             =   0
      Width           =   6000
   End
   Begin JinDuTiao.JinDuT JinDuT2 
      Height          =   255
      Left            =   6720
      TabIndex        =   6
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
   End
   Begin JinDuTiao.JinDuT JinDuT3 
      Height          =   255
      Left            =   6720
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      ProgressColor   =   33023
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      Height          =   400
      Left            =   6360
      Top             =   2260
      Width           =   400
   End
   Begin VB.Label SkOn 
      Alignment       =   2  'Center
      Caption         =   "盾"
      Enabled         =   0   'False
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
      Left            =   6840
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label SkOn 
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
      Height          =   400
      Index           =   0
      Left            =   6360
      TabIndex        =   1
      Top             =   2280
      Width           =   400
   End
   Begin VB.Label Label9 
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
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   10
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "EXP:"
      Height          =   615
      Left            =   6255
      TabIndex        =   13
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "CD:"
      Height          =   315
      Left            =   6255
      TabIndex        =   12
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "E:"
      Height          =   315
      Left            =   6360
      TabIndex        =   11
      Top             =   840
      Width           =   195
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "HP:"
      Height          =   315
      Left            =   6255
      TabIndex        =   10
      Top             =   480
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   $"AirWar.frx":0000
      Height          =   2835
      Left            =   6360
      TabIndex        =   9
      Top             =   3960
      Width           =   1995
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Timing:"
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   2760
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
Dim ActIme, BigFovPlaneTime, CunningFovPlaneTime As Long
Private T_s As Single

Private Sub Ftime_Timer()
PSkillID_Ft = PSkillID_Ft + 1
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
'37_40:4865
Dim i As Long
If KeyCode = 70 Or KeyCode = 82 Then
    If PSkillID = 0 And PSkill(1) = True Then PSkillID = 1: Pg(0).Blt = 2 Else PSkillID = 0: Pg(0).Blt = 1
    Form1.Shape1.Left = Form1.SkOn(PSkillID).Left: Form1.Shape1.Top = Form1.SkOn(PSkillID).Top - 20
End If
For i = 0 To 5
    If KCTemp(i) = KeyCode Then Exit Sub
Next
For i = 0 To 4
    If KCTemp(i) = 0 Then KCTemp(i) = KeyCode: Exit Sub
Next
KCTemp(5) = KeyCode
End Sub
Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
'1
Dim i As Long
For i = 0 To 5
    If KCTemp(i) = KeyCode Then KCTemp(i) = 0
Next
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'37_40:4865
Dim i As Long
If KeyCode = 76 Then
    If PSkillID = 0 And PSkill(1) = True Then PSkillID = 1: Pg(0).Blt = 2 Else PSkillID = 0: Pg(0).Blt = 1
    Form1.Shape1.Left = Form1.SkOn(PSkillID).Left: Form1.Shape1.Top = Form1.SkOn(PSkillID).Top - 20
    Exit Sub
End If
For i = 0 To 5
    If KCTemp(i) = KeyCode Then Exit Sub
Next
For i = 0 To 4
    If KCTemp(i) = 0 Then KCTemp(i) = KeyCode: Exit Sub
Next
KCTemp(5) = KeyCode
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'1
Dim i As Long
For i = 0 To 5
    If KCTemp(i) = KeyCode Then KCTemp(i) = 0
Next
End Sub
Private Sub Form_Load()
Picture1.Scale (0, Picture1.Height)-(Picture1.Width, 0)
World_Load
End Sub

Private Sub Timer_PC_Skill_Flash_Timer()

End Sub


Private Sub Test1_Click()

End Sub
Private Sub Timer1_Timer()
F5
End Sub
Private Sub Timer2_Timer()
Dim i As Long
Label3.Caption = KCTemp(0) & "," & KCTemp(1) & "," & KCTemp(2) & "," & KCTemp(3) & "," & KCTemp(4) & "," & KCTemp(5)
PBCD = True
If PBSkillCD = False Then PBSkillCDTime = PBSkillCDTime + 1
If PBSkillCDTime > 99 Then PBSkillCD = True: PBSkillCDTime = 0
'---------------------------Test------------------------------
Label8.Caption = "EXP:" & Pg(0).Emp & "/" & Pg(0).MxEmp
Label9.Caption = "LV:" & Pg(0).Rank
'-------------------------------------------------------------
ActIme = "134481924" '这里是英文输入法的码。
ActivateKeyboardLayout ActIme, 1
'---------------------------Skon------------------------------
If PSkill(1) = True Then SkOn(1).Enabled = True: SkOn(1).Visible = True Else SkOn(1).Enabled = False: SkOn(1).Visible = False
End Sub
Private Sub Timer3_Timer()
If Pg(0).a = True Then
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
BigFovPlaneTime = BigFovPlaneTime + 1
CunningFovPlaneTime = CunningFovPlaneTime + 1
PCEsp_Recovery
If FoePlaneLifeSum() <= Diff / 10 Then
    If BigFovPlaneTime < 15 - Diff / 10 Then
        Test_FoePlane
    Else
        Test_FoePlane 1: BigFovPlaneTime = 0
    End If
End If
If CunningFovPlaneTime > 60 - Diff / 10 Then
    Test_FoePlane 2: CunningFovPlaneTime = 0
End If
Test_Bullet_2
End Sub

Private Sub Timer6_Timer()
Dim i As Long
For i = 0 To FPSum - 1
    If FPg(i).a = False Then
        FPg_Add i, 1: Exit Sub
    End If
Next
FPg_Add FPSum, 1
FPSum = FPSum + 1
End Sub

Private Sub 开始_Click()
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
End Sub

Private Sub 双人_Click()
PC_2_Def
End Sub

Private Sub 暂停_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
End Sub

Private Sub 重来_Click()
BgSum = 0: T_s = 0: SgSum = 0
World_Load
End Sub
