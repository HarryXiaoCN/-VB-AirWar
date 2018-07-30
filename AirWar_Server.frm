VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AirWar Server"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6495
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer OlTime 
      Interval        =   100
      Left            =   120
      Top             =   3960
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   6000
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2018
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3180
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客户端列表："
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TCPNum As Long
Private Sub Form_Load()
Winsock1(0).Listen
End Sub
Private Sub Form_Unload(Cancel As Integer)
Local_State = 0
End Sub
Private Sub OlTime_Timer()
For i = 1 To TCPNum
    List1.List(i - 1) = "客户机IP：" & Winsock1(i).RemoteHostIP & vbTab & 连接状态反馈(Winsock1(i).State)
Next
End Sub
Private Sub Winsock1_Close(Index As Integer)
Winsock1(Index).Close
End Sub
Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal RequestID As Long)
If Index = 0 Then
    For i = 1 To TCPNum
        If Winsock1(i).State = 0 Then
            Winsock1(i).Accept RequestID
            Exit Sub
        End If
    Next
    TCPNum = TCPNum + 1
    Load Winsock1(TCPNum)
    Winsock1(TCPNum).LocalPort = 2018
    Winsock1(TCPNum).Accept RequestID
End If
End Sub
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim TCPData As String
Winsock1(Index).GetData TCPData
Server_GetData TCPData
End Sub
