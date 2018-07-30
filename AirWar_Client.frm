VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AirWar Client"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5520
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   4920
      Top             =   1560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "断开"
      Height          =   375
      Left            =   2815
      TabIndex        =   5
      Top             =   1080
      Width           =   2600
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Text            =   "2018"
      Top             =   600
      Width           =   3855
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4440
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "连接"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2600
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "服务器端口："
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "服务器地址："
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Client_GetData_tcpTemp_KG As Boolean
Private Sub Command1_Click()
On Error GoTo Er
Winsock1.RemoteHost = Text1.Text
Winsock1.RemotePort = Text2.Text
Winsock1.Connect
Exit Sub
Er:
MsgBox "连接失败！请检查与服务器间网络是否通畅或重复连接！"
End Sub
Private Sub Command2_Click()
Winsock1.Close
End Sub
Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
Local_State = 0
Form1.Spare.Enabled = False
End Sub
Private Sub Timer1_Timer()
If Winsock1.State = 7 Then Label3.Caption = "已连接;服务器IP：" & Winsock1.RemoteHostIP Else Label3.Caption = 连接状态反馈(Winsock1.State)
If Winsock1.State = 9 Then Winsock1.Close
End Sub
Private Sub Winsock1_Close()
Winsock1.Close
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData Client_GetData_tcpTemp
If TCPGetDataShow = True Then Form2.Text1.Text = Client_GetData_tcpTemp
Client_GetData
End Sub
