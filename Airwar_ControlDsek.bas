Attribute VB_Name = "Airwar_ControlDsek"
Public CMDShow, ClientGetDataShow, ServerSendDataShow, ClientSendDataShow, ServerGetDataShow As Boolean
Public NetworkEcho As Single
Public History(100) As String
Public HistorySum, HistoryNow As Long
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
            Case "revive"
                Pg(CMD_Temp(2)).a = True: Pg(CMD_Temp(2)).HP = Pg(CMD_Temp(2)).MxHP: GoTo Successfully
            Case "hp_add"
                If Pg(CMD_Temp(2)).HP + Val(CMD_Temp(3)) < 100000000 Then Pg(CMD_Temp(2)).HP = Pg(CMD_Temp(2)).HP + Val(CMD_Temp(3)) Else Pg(CMD_Temp(2)).HP = 99999999
                GoTo Successfully
            Case "e_mx"
                Pg(CMD_Temp(2)).E = Pg(CMD_Temp(2)).MxE: GoTo Successfully
            Case "e_add"
                If Pg(CMD_Temp(2)).E + Val(CMD_Temp(3)) < 100000000 Then Pg(CMD_Temp(2)).E = Pg(CMD_Temp(2)).E + Val(CMD_Temp(3)) Else Pg(CMD_Temp(2)).E = 99999999
                GoTo Successfully
            Case "mxhp_add"
                If Pg(CMD_Temp(2)).MxHP + Val(CMD_Temp(3)) < 100000000 Then Pg(CMD_Temp(2)).MxHP = Pg(CMD_Temp(2)).MxHP + Val(CMD_Temp(3)) Else Pg(CMD_Temp(2)).MxHP = 99999999
                GoTo Successfully
            Case "mxe_add"
                If Pg(CMD_Temp(2)).MxE + Val(CMD_Temp(3)) < 100000000 Then Pg(CMD_Temp(2)).MxE = Pg(CMD_Temp(2)).MxE + Val(CMD_Temp(3)) Else Pg(CMD_Temp(2)).MxE = 99999999
                GoTo Successfully
            Case "sp_add"
                If Pg(CMD_Temp(2)).Sp + Val(CMD_Temp(3)) < 0 Then CMD_Execute_Interpreter = CMD & "--Error: Speed cannot be negative": Exit Function
                If Pg(CMD_Temp(2)).Sp + Val(CMD_Temp(3)) < 100000000 Then Pg(CMD_Temp(2)).Sp = Pg(CMD_Temp(2)).Sp + Val(CMD_Temp(3)) Else Pg(CMD_Temp(2)).Sp = 99999999
                GoTo Successfully
            Case "exp_add"
                If Val(CMD_Temp(3)) < 1000000000 Then Pg(CMD_Temp(2)).EMP = Pg(CMD_Temp(2)).EMP + Val(CMD_Temp(3)): Éý¼¶ Val(CMD_Temp(2)) Else Pg(CMD_Temp(2)).EMP = Pg(CMD_Temp(2)).EMP + 999999999
                GoTo Successfully
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
            CMD_Execute_Interpreter = CMD & "--Error: 404 Found": Exit Function
        End If
    Case "foe"
        Select Case CMD_Temp(1)
            Case "plane_add"
                Test_FoePlane Val(CMD_Temp(2)): GoTo Successfully
        End Select
    Case "vis"
        Select Case CMD_Temp(1)
            Case "keyboard"
                If Val(CMD_Temp(2)) = 0 Then KeyboardVis = False: GoTo Successfully Else KeyboardVis = True: GoTo Successfully
            Case "rein"
                If Val(CMD_Temp(2)) = 0 Then ReinCodeVis = False: GoTo Successfully Else ReinCodeVis = True: GoTo Successfully
        End Select
    Case "time"
        T_s = Val(CMD_Temp(1)): GoTo Successfully
End Select
CMD_Execute_Interpreter = CMD & "--Unknown Command"
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
