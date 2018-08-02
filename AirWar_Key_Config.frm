VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Key Config"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   4155
   ScaleWidth      =   7935
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7560
      Top             =   3960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Player2"
      Height          =   3855
      Index           =   1
      Left            =   4000
      TabIndex        =   8
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         Caption         =   "ÉÏ£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   13
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "ÏÂ£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "×ó£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "ÓÒ£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "¹¥»÷£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "´óÕÐ£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "´óÕÐÇÐ»»£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Player1"
      Height          =   3855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         Caption         =   "´óÕÐÇÐ»»£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   3240
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "´óÕÐ£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "¹¥»÷£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "ÓÒ£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "×ó£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "ÏÂ£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "ÉÏ£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private KeyTargetLock As Boolean
Private KeyTarget As Long
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyTargetLock = True Then
    Select Case KeyTarget
        Case 0
            PC(0).up = KeyCode
        Case 1
            PC(0).down = KeyCode
        Case 2
            PC(0).left = KeyCode
        Case 3
            PC(0).right = KeyCode
        Case 4
            PC(0).attack = KeyCode
        Case 5
            PC(0).ultimate_skill = KeyCode
        Case 6
            PC(0).skill_switch = KeyCode
        Case 7
            PC(1).up = KeyCode
        Case 8
            PC(1).down = KeyCode
        Case 9
            PC(1).left = KeyCode
        Case 10
            PC(1).right = KeyCode
        Case 11
            PC(1).attack = KeyCode
        Case 12
            PC(1).ultimate_skill = KeyCode
        Case 13
            PC(1).skill_switch = KeyCode
    End Select
End If
KeyTargetLock = False
Label1(KeyTarget).BackColor = &H8000000F
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
If KeyTargetLock = True Then Exit Sub
For i = 0 To 13
    Label1(i).BackColor = &H8000000F
Next
End Sub
Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
If KeyTargetLock = True Then Exit Sub
For i = 0 To 13
    Label1(i).BackColor = &H8000000F
Next
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
KeyTargetLock = True: KeyTarget = Index
Label1(Index).BackColor = RGB(255, 215, 0)
End Sub
Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
If KeyTargetLock = True Then Exit Sub
Label1(Index).BackColor = RGB(127, 127, 127)
For i = 0 To Index - 1
    Label1(i).BackColor = &H8000000F
Next
For i = Index + 1 To 13
    Label1(i).BackColor = &H8000000F
Next
End Sub

Private Sub Timer1_Timer()
Dim i As Long
For i = 0 To 1
    With PC(i)
        Label1(i * 7).Caption = "ÉÏ£º" & KeyCodeToStr(.up)
        Label1(i * 7 + 1).Caption = "ÏÂ£º" & KeyCodeToStr(.down)
        Label1(i * 7 + 2).Caption = "×ó£º" & KeyCodeToStr(.left)
        Label1(i * 7 + 3).Caption = "ÓÒ£º" & KeyCodeToStr(.right)
        Label1(i * 7 + 4).Caption = "¹¥»÷£º" & KeyCodeToStr(.attack)
        Label1(i * 7 + 5).Caption = "´óÕÐ£º" & KeyCodeToStr(.ultimate_skill)
        Label1(i * 7 + 6).Caption = "´óÕÐÇÐ»»£º" & KeyCodeToStr(.skill_switch)
    End With
Next
End Sub
