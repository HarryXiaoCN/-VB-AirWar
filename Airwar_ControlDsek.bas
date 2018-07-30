Attribute VB_Name = "Airwar_ControlDsek"
Public CMDShow, TCPGetDataShow, TCPSendDataShow As Boolean
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
                Pg(CMD_Temp(2)).HP = Pg(CMD_Temp(2)).MxHP
            Case "e_mx"
                Pg(CMD_Temp(2)).E = Pg(CMD_Temp(2)).MxE
            Case "mxhp_add"
                Pg(CMD_Temp(2)).MxHP = Pg(CMD_Temp(2)).MxHP + Val(CMD_Temp(3))
            Case "mxe_add"
                Pg(CMD_Temp(2)).MxE = Pg(CMD_Temp(2)).MxE + Val(CMD_Temp(3))
        End Select
    Case "network"
        Select Case CMD_Temp(1)
            Case "client"
                If CMD_Temp(2) = "0" Then TCPGetDataShow = False Else TCPGetDataShow = True
            Case "server"
                If CMD_Temp(2) = "0" Then TCPSendDataShow = False Else TCPSendDataShow = True
        End Select
End Select
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
