VERSION 5.00
Object = "{A2ED65B5-7DB0-4716-871E-53783A22D5D9}#66.0#0"; "JinDuT.ocx"
Begin VB.Form Form1 
   Caption         =   "AirWar"
   ClientHeight    =   8130
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12060
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   12060
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Ftime 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   1000
      Left            =   6960
      Top             =   7080
   End
   Begin VB.Frame Frame1 
      Caption         =   "PLAYER2"
      Height          =   3135
      Index           =   1
      Left            =   9120
      TabIndex        =   16
      Top             =   0
      Width           =   2775
      Begin JinDuTiao.JinDuT JinDuT1 
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   17
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
         TabIndex        =   18
         Top             =   1305
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
      End
      Begin JinDuTiao.JinDuT JinDuT3 
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   19
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
            Name            =   "Î¢ÈíÑÅºÚ"
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
         TabIndex        =   26
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "EXP:"
         Height          =   315
         Index           =   1
         Left            =   135
         TabIndex        =   25
         Top             =   2025
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CD:"
         Height          =   315
         Index           =   1
         Left            =   135
         TabIndex        =   24
         Top             =   1665
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "E:"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   1305
         Width           =   195
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "HP:"
         Height          =   315
         Index           =   1
         Left            =   135
         TabIndex        =   22
         Top             =   945
         Width           =   390
      End
      Begin VB.Label SkOn2 
         Alignment       =   2  'Center
         Caption         =   "¶Ü"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
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
         TabIndex        =   21
         Top             =   2655
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label SkOn2 
         Alignment       =   2  'Center
         Caption         =   "ºÚ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
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
         TabIndex        =   20
         Top             =   2655
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PLAYER1"
      Height          =   3135
      Index           =   0
      Left            =   6240
      TabIndex        =   5
      Top             =   0
      Width           =   2775
      Begin JinDuTiao.JinDuT JinDuT1 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   6
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
         TabIndex        =   7
         Top             =   1305
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
      End
      Begin JinDuTiao.JinDuT JinDuT3 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   8
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
         Caption         =   "ºÚ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
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
         TabIndex        =   15
         Top             =   2655
         Width           =   405
      End
      Begin VB.Label SkOn1 
         Alignment       =   2  'Center
         Caption         =   "¶Ü"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   945
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "E:"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   1305
         Width           =   195
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CD:"
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   1665
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "EXP:"
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   2025
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "LV:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
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
         TabIndex        =   9
         Top             =   480
         Width           =   435
      End
   End
   Begin VB.Timer Ftime 
      Enabled         =   0   'False
      Index           =   0
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
         Name            =   "ËÎÌå"
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   $"AirWar.frx":0000
      Height          =   2835
      Left            =   9120
      TabIndex        =   4
      Top             =   3240
      Width           =   2715
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Timing:"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Menu ²Ëµ¥ 
      Caption         =   "²Ëµ¥"
      Begin VB.Menu ¿ªÊ¼ 
         Caption         =   "¿ªÊ¼"
         Shortcut        =   ^K
      End
      Begin VB.Menu ÖØÀ´ 
         Caption         =   "ÖØÀ´"
         Shortcut        =   ^C
      End
      Begin VB.Menu ÔÝÍ£ 
         Caption         =   "ÔÝÍ£"
         Shortcut        =   ^Z
      End
      Begin VB.Menu Ë«ÈË 
         Caption         =   "Ë«ÈË"
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

Private Sub Ftime_Timer(Index As Integer)
PSkillID_Ft(Index) = PSkillID_Ft(Index) + 1
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
'37_40:4865
Dim i As Long
If KeyCode = 76 Then
    If PSkillID(0) = 0 And PSkill(0, 1) = True Then PSkillID(0) = 1: Pg(0).Blt = 2 Else PSkillID(0) = 0: Pg(0).Blt = 1
    Form1.Shape1(0).Left = Form1.SkOn1(PSkillID(0)).Left: Form1.Shape1(0).Top = Form1.SkOn1(PSkillID(0)).Top - 20
    Exit Sub
End If
If KeyCode = 99 Then
    If PSkillID(1) = 0 And PSkill(1, 1) = True Then PSkillID(1) = 1: Pg(1).Blt = 2 Else PSkillID(1) = 0: Pg(1).Blt = 1
    Form1.Shape1(1).Left = Form1.SkOn2(PSkillID(1)).Left: Form1.Shape1(1).Top = Form1.SkOn2(PSkillID(1)).Top - 20
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
Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
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
Private Sub Timer1_Timer()
F5
End Sub
Private Sub Timer2_Timer()
Dim i As Long
Label3.Caption = KCTemp(0) & "," & KCTemp(1) & "," & KCTemp(2) & "," & KCTemp(3) & "," & KCTemp(4) & "," & KCTemp(5)

'---------------------------Test------------------------------
For i = 0 To 1
    If Pg(i).a = True Then
        PBCD(i) = True
        If PBSkillCD(i) = False Then PBSkillCDTime(i) = PBSkillCDTime(i) + 1
        If PBSkillCDTime(i) > 99 Then PBSkillCD(i) = True: PBSkillCDTime(i) = 0
        Label8(i).Caption = "EXP:" & Pg(i).Emp & "/" & Pg(i).MxEmp
        Label9(i).Caption = "LV:" & Pg(i).Rank
    End If
Next
'-------------------------------------------------------------
ActIme = "134481924" 'ÕâÀïÊÇÓ¢ÎÄÊäÈë·¨µÄÂë¡£
ActivateKeyboardLayout ActIme, 1
'---------------------------Skon------------------------------
If PSkill(0, 1) = True Then SkOn1(1).Enabled = True: SkOn1(1).Visible = True Else SkOn1(1).Enabled = False: SkOn1(1).Visible = False
If PSkill(1, 1) = True Then SkOn2(1).Enabled = True: SkOn2(1).Visible = True Else SkOn2(1).Enabled = False: SkOn2(1).Visible = False
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

Private Sub ¿ªÊ¼_Click()
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
End Sub

Private Sub Ë«ÈË_Click()
PC_2_Def
DuoPlayer = True
End Sub

Private Sub ÔÝÍ£_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
End Sub

Private Sub ÖØÀ´_Click()
BgSum = 0: T_s = 0: SgSum = 0
World_Load
End Sub
