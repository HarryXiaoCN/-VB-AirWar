VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Desk"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9015
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   9015
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Execute"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   8055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
History(HistorySum) = Text2.Text
CMD_Execute (Text2.Text)
HistorySum = HistorySum + 1
HistoryNow = HistorySum
End Sub
Private Sub Form_Load()
HistoryNow = HistorySum
End Sub
Private Sub Form_Unload(Cancel As Integer)
If CMDShow = False Then CMDShow = True: World_Stop: Form2.Show Else CMDShow = False: World_Start: Unload Form2
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Er
    If KeyCode = 13 Then
        Command1_Click
    End If
    If KeyCode = 192 Or KeyCode = 27 Then
        Unload Form2
    End If
    If KeyCode = 38 And HistorySum > 0 Then
        HistoryNow = HistoryNow + 1
        Text2.Text = History(HistorySum - HistoryNow)
        Text2.SelStart = Len(Text2.Text)
    End If
    If KeyCode = 40 And HistoryNow > 0 Then
        HistoryNow = HistoryNow - 1
        Text2.Text = History(HistorySum - HistoryNow)
        Text2.SelStart = Len(Text2.Text)
    End If
Er:
End Sub
