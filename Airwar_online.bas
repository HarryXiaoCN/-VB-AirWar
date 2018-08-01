Attribute VB_Name = "Airwar_online"
Public Local_State, TCPNum, Client_GetData_Temp As Long
Public Client_SendData_Coding_Temp, Client_GetData_Key, Client_GetData_tcpTemp As String
Public Function Server_GetData(ByVal GetData As String)
Dim DataTemp, Temp
Dim i, c As Long
If ServerGetDataShow = True Then Form2.Text1.Text = GetData
Temp = Split(GetData, "|")
On Error GoTo Er
For i = 0 To UBound(Temp) - 1
    If Temp(i) = "Echo" Then
        Form2.Text1.Text = Form2.Text1.Text & Timer - NetworkEcho & "ms" & vbCrLf
        GoTo BreakHanlder
    End If
    DataTemp = Split(Temp(i), ",")
    For c = 3 To 5
        KCTemp(c) = DataTemp(c - 3)
    Next
BreakHanlder:
Next
Er:
End Function
Public Function Server_SendData(ByVal SK As Long, ByVal X As Single, ByVal Y As Single, ByVal R As Long, _
ByVal FR As Long, ByVal FG As Long, ByVal FB As Long, _
ByVal CR As Long, ByVal CG As Long, ByVal CB As Long, _
Optional Break As Long, Optional PlayerID As Long, _
Optional LV As String, Optional HpP As Long, Optional EP As Long, _
Optional SCD As Long, Optional EMP As String, Optional SkillOpt)
Dim i As Long: Dim StrData As String
Select Case Break
    Case 0
        StrData = "H;" & SK & ";" & X & ";" & Y & ";" & R & ";" & FR & ";" & FG & ";" & FB & ";" & CR & ";" & CG & ";" & CB & "|"
    Case 1
        StrData = "Break|"
    Case 2
        StrData = "G;" & PlayerID & ";" & LV & ";" & HpP & ";" & EP & ";" & SCD & ";" & EMP & ";" & SkillOpt & ";" & Form1.Label2.Caption & "|"
    Case 3
        StrData = "Renew|"
End Select
For i = 1 To TCPNum
    If Form3.Winsock1(i).State = 7 Then
        Form3.Winsock1(i).SendData StrData
    End If
Next
If ServerSendDataShow = True Then Form2.Text1.Text = StrData
End Function
Public Function Client_GetData()
Dim DataTemp, Temp: Dim i As Long
Temp = Split(Client_GetData_tcpTemp, "|")
On Error GoTo BreakHanlder
If ClientGetDataShow = True Then Form2.Text1.Text = Client_GetData_tcpTemp
For i = 0 To UBound(Temp) - 1
If Temp(i) = "Renew" Then
    Form1.SkOn1(1).Visible = False: Form1.SkOn2(1).Visible = False
    GoTo BreakHanlder
End If
If Temp(i) = "Break" Then
    Form1.Picture1.Cls
    GoTo BreakHanlder
End If
If Temp(i) = "Echo" Then
    Form4.Winsock1.SendData "Echo|"
    GoTo BreakHanlder
End If
DataTemp = Split(Temp(i), ";")
Select Case DataTemp(0)
    Case "H"
        Form1.Picture1.FillStyle = Val(DataTemp(1))
        Form1.Picture1.FillColor = RGB(Val(DataTemp(5)), Val(DataTemp(6)), Val(DataTemp(7)))
        Form1.Picture1.Circle (Val(DataTemp(2)), Val(DataTemp(3))), Val(DataTemp(4)), RGB(Val(DataTemp(8)), Val(DataTemp(9)), Val(DataTemp(10)))
    Case "G"
        Form1.Label9(Val(DataTemp(1))).Caption = DataTemp(2)
        Form1.JinDuT1(Val(DataTemp(1))).Progress = Val(DataTemp(3))
        Form1.JinDuT2(Val(DataTemp(1))).Progress = Val(DataTemp(4))
        Form1.JinDuT3(Val(DataTemp(1))).Progress = Val(DataTemp(5))
        Form1.Label8(Val(DataTemp(1))).Caption = DataTemp(6)
        PSkillID(Val(DataTemp(1))) = Val(DataTemp(7))
        Form1.Label2.Caption = DataTemp(8)
End Select
BreakHanlder:
Next
End Function
Public Function Client_SendData()
Dim CSendData As String: Dim i As Long
For i = 3 To 4
    CSendData = CSendData & KCTemp(i) & ","
Next
CSendData = CSendData & KCTemp(5) & "|"
If ClientSendDataShow = True Then Form2.Text1.Text = CSendData
If Form4.Winsock1.State = 7 Then Form4.Winsock1.SendData CSendData
End Function
Public Function Local_State_Vision()
Select Case Local_State
    Case 0
        Form1.Label4.Caption = "单机模式"
    Case 1
        Form1.Label4.Caption = "服务器联机"
    Case 2
        Form1.Label4.Caption = "客户端联机"
End Select
End Function
Public Function KeyCode_Filtration(ByRef KeyCode As Integer) As Boolean
Select Case Local_State
    Case 1
        Select Case KeyCode
            Case 37, 38, 39, 40, 97, 98
            KeyCode_Filtration = True
        End Select
    Case 2
        Select Case KeyCode
            Case 65, 87, 68, 83, 74, 75
            KeyCode_Filtration = True
        End Select
End Select
End Function
Public Function 连接状态反馈(ByVal SID As Long) As String
Select Case SID
    Case 0
        连接状态反馈 = "未连接"
    Case 1
        连接状态反馈 = "打开状态"
    Case 2
        连接状态反馈 = "等待连接"
    Case 3
        连接状态反馈 = "连接挂起"
    Case 4
        连接状态反馈 = "域名解析中..."
    Case 5
        连接状态反馈 = "域名解析成功"
    Case 6
        连接状态反馈 = "正在连接..."
    Case 7
        连接状态反馈 = "已连接"
    Case 9
        连接状态反馈 = "连接错误"
End Select
End Function
