Attribute VB_Name = "Airwar_ControlDsek"
Public CMDShow, ClientGetDataShow, ServerSendDataShow, ClientSendDataShow, ServerGetDataShow As Boolean
Public NetworkEcho As Single
Public Function CMD_Execute(ByRef CMD As String)
Form2.Text1 = Form2.Text1 & CMD_Execute_Interpreter(CMD_Big_Small(CMD)) & vbCrLf
Form2.Text1.SelStart = Len(Form2.Text1.Text)
Form2.Text2 = ""
End Function
Public Function CMD_Execute_Interpreter(ByRef CMD As String)
Dim CMD_Temp
On Error GoTo ErrHandler
CMD_Temp = Split(CMD, " ")
Select Case CMD_Temp(0)
    Case "player"
        Select Case CMD_Temp(1)
            Case "hp_mx"
                Pg(CMD_Temp(2)).HP = Pg(CMD_Temp(2)).MxHP: GoTo Successfully
            Case "e_mx"
                Pg(CMD_Temp(2)).E = Pg(CMD_Temp(2)).MxE: GoTo Successfully
            Case "mxhp_add"
                Pg(CMD_Temp(2)).MxHP = Pg(CMD_Temp(2)).MxHP + Val(CMD_Temp(3)): GoTo Successfully
            Case "mxe_add"
                Pg(CMD_Temp(2)).MxE = Pg(CMD_Temp(2)).MxE + Val(CMD_Temp(3)): GoTo Successfully
            Case "sp_add"
                Pg(CMD_Temp(2)).Sp = Pg(CMD_Temp(2)).Sp + Val(CMD_Temp(3)): GoTo Successfully
        End Select
    Case "network"
        Select Case CMD_Temp(1)
            Case "clientget"
                If CMD_Temp(2) = "0" Then ClientGetDataShow = False: GoTo Successfully Else ClientGetDataShow = True: GoTo Successfully
            Case "serversend"
                If CMD_Temp(2) = "0" Then ServerSendDataShow = False: GoTo Successfully Else ServerSendDataShow = True: GoTo Successfully
            Case "clientsend"
                If CMD_Temp(2) = "0" Then ClientSendDataShow = False: GoTo Successfully Else ClientSendDataShow = True: GoTo Successfully
            Case "serverget"
                If CMD_Temp(2) = "0" Then ServerGetDataShow = False: GoTo Successfully Else ServerGetDataShow = True: GoTo Successfully
        End Select
    Case "ping"
        If Form3.Winsock1(Val(CMD_Temp(1))).State = 7 Then
            Form3.Winsock1(Val(CMD_Temp(1))).SendData "Echo|": NetworkEcho = Timer: GoTo Successfully
        Else
            CMD_Execute_Interpreter = CMD & "--Error 404 Found": Exit Function
        End If
End Select
CMD_Execute_Interpreter = CMD & "--Unknown command"
Exit Function
Successfully:
CMD_Execute_Interpreter = CMD & "--Successfully"
Exit Function
ErrHandler:
CMD_Execute_Interpreter = CMD & "--Error"
End Function
Public Function CMD_Big_Small(ByRef CMD As String) As String
Dim i As Long
CMD_Big_Small = CMD
For i = 65 To 90
    CMD_Big_Small = Replace(CMD_Big_Small, Chr(i), Chr(i + 32))
Next
End Function
