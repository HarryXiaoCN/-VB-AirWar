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
        Form1.Label4.Caption = "����ģʽ"
    Case 1
        Form1.Label4.Caption = "����������"
    Case 2
        Form1.Label4.Caption = "�ͻ�������"
End Select
End Function
Public Function ����״̬����(ByVal SID As Long) As String
Select Case SID
    Case 0
        ����״̬���� = "δ����"
    Case 1
        ����״̬���� = "��״̬"
    Case 2
        ����״̬���� = "�ȴ�����"
    Case 3
        ����״̬���� = "���ӹ���"
    Case 4
        ����״̬���� = "����������..."
    Case 5
        ����״̬���� = "���������ɹ�"
    Case 6
        ����״̬���� = "��������..."
    Case 7
        ����״̬���� = "������"
    Case 9
        ����״̬���� = "���Ӵ���"
End Select
End Function
