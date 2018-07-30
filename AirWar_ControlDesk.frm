VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "ControlDesk"
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9060
   LinkTopic       =   "Form2"
   ScaleHeight     =   1350
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
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
Private History(100) As String
Private HistorySum, HistoryNow As Long
Private Sub Command1_Click()
History(HistorySum) = Text2.Text
CMD_Execute (Text2.Text)
HistorySum = HistorySum + 1
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Command1_Click
    End If
    If KeyCode = 192 Or KeyCode = 27 Then
        If CMDShow = False Then CMDShow = True: World_Stop: Form2.Show Else CMDShow = False: World_Start: Unload Form2
    End If
    If KeyCode = 38 And HistorySum > 0 Then
        HistoryNow = HistoryNow + 1
        Text2.Text = History(HistorySum - HistoryNow)
    End If
    If KeyCode = 40 And HistoryNow > 0 Then
        HistoryNow = HistoryNow - 1
        Text2.Text = History(HistorySum - HistoryNow)
    End If
End Sub
