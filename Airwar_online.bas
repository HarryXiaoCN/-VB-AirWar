Attribute VB_Name = "Airwar_online"
Public Local_State As Long
Public Client_SendData_Coding_Temp As String
Public Function Server_GetData(ByVal GetData As String)
Dim DataTemp
Dim i As Long
DataTemp = Split(GetData, ",")
Select Case Val(DataTemp(2))
    Case 0
        KCTemp(DataTemp(0)) = DataTemp(1)
    Case 1
        KCTemp(DataTemp(0)) = 0
End Select
End Function
Public Function Server_SendData()

End Function
Public Function Client_GetData(ByVal GetData As String)

End Function
Public Function Client_SendData(ByVal SendDataKeyID As Integer, ByVal SendDataBtye As Integer, ByVal DownUp As Integer)
Form4.Winsock1.SendData SendDataKeyID & "," & SendDataBtye & "," & DownUp
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
