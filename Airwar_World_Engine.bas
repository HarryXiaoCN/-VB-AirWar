Attribute VB_Name = "Airwar_World_Engine"
Public Function World_Stop()
Form1.Timer1.Enabled = False
Form1.Timer2.Enabled = False
Form1.Timer3.Enabled = False
Form1.Timer4.Enabled = False
Form1.Timer5.Enabled = False
End Function
Public Function World_Start()
Form1.Timer1.Enabled = True
Form1.Timer2.Enabled = True
Form1.Timer3.Enabled = True
Form1.Timer4.Enabled = True
Form1.Timer5.Enabled = True
If Local_State = 1 Then Server_SendData_Ignition "Renew|"
End Function
Public Function F5()
Select Case Local_State
    Case 0
        Form1.Picture1.Cls
        物理事件检测
        移动事件处理
        F5_Foe_Bullet
        F5_PC_Supply
        F5_PC_Bullet
        F5_Foe_Plane
        F5_PC_Plane
    Case 1
        Form1.Picture1.Cls
        Server_SendData_Ignition "Break|"
        物理事件检测
        移动事件处理
        F5_NetworkMode
        F5_Foe_Bullet
        F5_PC_Supply
        F5_PC_Bullet
        F5_Foe_Plane
        F5_PC_Plane
End Select
End Function
Public Function F5_NetworkMode()
Dim i As Long
For i = 0 To 1
    Server_SendData_Ui i, Form1.Label9(i).Caption, Form1.JinDuT1(i).Progress, Form1.JinDuT2(i).Progress, Form1.JinDuT3(i).Progress, Form1.Label8(i).Caption, PSkillID(i)
Next
End Function
Public Function F5_Foe_Plane()
Form1.Picture1.FillColor = RGB(143, 188, 143)
For i = 0 To FPSum - 1
    If FPg(i).A = True Then
        Select Case FPg(i).AiRank
            Case 0
                Form1.Picture1.Circle (FPg(i).X, FPg(i).Y), FPg(i).Ar, RGB(143, 188, 143)
                If Local_State = 1 Then Server_SendData_Circle 0, FPg(i).X, FPg(i).Y, FPg(i).Ar, 143, 188, 143, 143, 188, 143
            Case 1
                Form1.Picture1.FillColor = RGB(123, 104, 238)
                Form1.Picture1.Circle (FPg(i).X, FPg(i).Y), FPg(i).Ar, RGB(123, 104, 238)
                If Local_State = 1 Then Server_SendData_Circle 0, FPg(i).X, FPg(i).Y, FPg(i).Ar, 123, 104, 238, 123, 104, 238
                Form1.Picture1.FillColor = RGB(143, 188, 143)
            Case 2
                Form1.Picture1.FillColor = RGB(255, 69, 0)
                Form1.Picture1.Circle (FPg(i).X, FPg(i).Y), FPg(i).Ar, RGB(255, 69, 0)
                If Local_State = 1 Then Server_SendData_Circle 0, FPg(i).X, FPg(i).Y, FPg(i).Ar, 255, 69, 0, 255, 69, 0
                Form1.Picture1.FillColor = RGB(143, 188, 143)
            Case 3
                Form1.Picture1.FillColor = RGB(175, 238, 238)
                Form1.Picture1.Circle (FPg(i).X, FPg(i).Y), FPg(i).Ar, RGB(175, 238, 238)
                If Local_State = 1 Then Server_SendData_Circle 0, FPg(i).X, FPg(i).Y, FPg(i).Ar, 175, 238, 238, 175, 238, 238
                Form1.Picture1.FillColor = RGB(143, 188, 143)
            Case 4
                Form1.Picture1.FillColor = RGB(0, 0, 0)
                Form1.Picture1.Circle (FPg(i).X, FPg(i).Y), FPg(i).Ar, RGB(0, 0, 0)
                If Local_State = 1 Then Server_SendData_Circle 0, FPg(i).X, FPg(i).Y, FPg(i).Ar, 0, 0, 0, 0, 0, 0
                Form1.Picture1.FillColor = RGB(143, 188, 143)
            Case 5
                Form1.Picture1.FillColor = RGB(75, 0, 130)
                Form1.Picture1.Circle (FPg(i).X, FPg(i).Y), FPg(i).Ar, RGB(75, 0, 130)
                If Local_State = 1 Then Server_SendData_Circle 0, FPg(i).X, FPg(i).Y, FPg(i).Ar, 75, 0, 130, 75, 0, 130
                Form1.Picture1.FillColor = RGB(143, 188, 143)
        End Select
    End If
Next
End Function
Public Function F5_PC_Bullet()
Dim i As Long: Dim SYSE As 三原色
Form1.Picture1.FillColor = RGB(0, 255, 255)
For i = 0 To PBSum - 1
    If PBg(i).A = True Then
        Select Case PBg(i).Trl
            Case 0
                Form1.Picture1.Circle (PBg(i).X, PBg(i).Y), PBg(i).Ar, RGB(0, 255, 255)
                If Local_State = 1 Then Server_SendData_Circle 0, PBg(i).X, PBg(i).Y, PBg(i).Ar, 0, 255, 255, 0, 255, 255
            Case 1
                Form1.Picture1.FillColor = RGB(128, 0, 0)
                Form1.Picture1.Circle (PBg(i).X, PBg(i).Y), PBg(i).Ar, RGB(128, 0, 0)
                If Local_State = 1 Then Server_SendData_Circle 0, PBg(i).X, PBg(i).Y, PBg(i).Ar, 128, 0, 0, 128, 0, 0
                Form1.Picture1.FillColor = RGB(0, 255, 255)
            Case 2
                Randomize
                Form1.Picture1.FillStyle = 1
                Form1.Picture1.Circle (PBg(i).X, PBg(i).Y), PBg(i).Ar, RGB(100 + PSkillID_Ft(PBg(i).Source) * 30, 149 + PSkillID_Ft(PBg(i).Source) * 20, 237)
                If Local_State = 1 Then Server_SendData_Circle 1, PBg(i).X, PBg(i).Y, PBg(i).Ar, 100 + PSkillID_Ft(PBg(i).Source) * 30, 149 + PSkillID_Ft(PBg(i).Source) * 20, 237, 100 + PSkillID_Ft(PBg(i).Source) * 30, 149 + PSkillID_Ft(PBg(i).Source) * 20, 237
                Form1.Picture1.Circle (PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20), PBg(i).Ar + Rnd * 20 - 20, RGB(176, 196, 222)
                If Local_State = 1 Then Server_SendData_Circle 1, PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20, PBg(i).Ar + Rnd * 20 - 20, 176, 196, 222, 176, 196, 222
                Form1.Picture1.Circle (PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20), PBg(i).Ar + Rnd * 20 - 20, RGB(176, 196, 222)
                If Local_State = 1 Then Server_SendData_Circle 1, PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20, PBg(i).Ar + Rnd * 20 - 20, 176, 196, 222, 176, 196, 222
                Form1.Picture1.Circle (PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20), PBg(i).Ar + Rnd * 20 - 20, RGB(176, 196, 222)
                If Local_State = 1 Then Server_SendData_Circle 1, PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20, PBg(i).Ar + Rnd * 20 - 20, 176, 196, 222, 176, 196, 222
                Form1.Picture1.FillStyle = 0
            Case 3
                Form1.Picture1.FillColor = RGB(127, 255, 0)
                Form1.Picture1.Circle (PBg(i).X, PBg(i).Y), PBg(i).Ar, RGB(127, 255, 0)
                If Local_State = 1 Then Server_SendData_Circle 0, PBg(i).X, PBg(i).Y, PBg(i).Ar, 127, 255, 0, 127, 255, 0
                Form1.Picture1.FillColor = RGB(0, 255, 255)
            Case 4
                Form1.Picture1.FillColor = RGB(0, 0, 0)
                Form1.Picture1.Circle (PBg(i).X, PBg(i).Y), PBg(i).Ar, RGB(0, 0, 0)
                If Local_State = 1 Then Server_SendData_Circle 0, PBg(i).X, PBg(i).Y, PBg(i).Ar, 0, 0, 0, 0, 0, 0
                Form1.Picture1.FillColor = RGB(0, 255, 255)
            Case 5
                SYSE = 彩虹函数
                Form1.Picture1.FillColor = RGB(SYSE.R, SYSE.G, SYSE.B)
                Form1.Picture1.Circle (PBg(i).X, PBg(i).Y), PBg(i).Ar, RGB(SYSE.R, SYSE.G, SYSE.B)
                If Local_State = 1 Then Server_SendData_Circle 0, PBg(i).X, PBg(i).Y, PBg(i).Ar, SYSE.R, SYSE.G, SYSE.B, SYSE.R, SYSE.G, SYSE.B
                Form1.Picture1.FillColor = RGB(0, 255, 255)
        End Select
    End If
Next
End Function
Public Function F5_PC_Supply()
Form1.Picture1.FillColor = RGB(0, 255, 0)
For i = 0 To SgSum - 1
    If Sg(i).A = True Then
        Select Case Sg(i).Tp
            Case 0
                Form1.Picture1.Circle (Sg(i).X, Sg(i).Y), 50, RGB(255, 0, 0)
                If Local_State = 1 Then Server_SendData_Circle 0, Sg(i).X, Sg(i).Y, 50, 0, 255, 0, 255, 0, 0
            Case 1
                Form1.Picture1.Circle (Sg(i).X, Sg(i).Y), 50, RGB(0, 0, 255)
                If Local_State = 1 Then Server_SendData_Circle 0, Sg(i).X, Sg(i).Y, 50, 0, 255, 0, 0, 0, 255
            Case 2
                Form1.Picture1.Circle (Sg(i).X, Sg(i).Y), 50, RGB(30, 144, 255)
                If Local_State = 1 Then Server_SendData_Circle 0, Sg(i).X, Sg(i).Y, 50, 0, 255, 0, 30, 144, 255
        End Select
    End If
Next
End Function
Public Function F5_Foe_Bullet()
Dim i As Long
Form1.Picture1.FillColor = RGB(0, 0, 255)
For i = 0 To BgSum - 1
    If Bg(i).A = True Then
        Select Case Bg(i).Trl
            Case 0
                Form1.Picture1.Circle (Bg(i).X, Bg(i).Y), Bg(i).Ar, RGB(0, 0, 255)
                If Local_State = 1 Then Server_SendData_Circle 0, Bg(i).X, Bg(i).Y, Bg(i).Ar, 0, 0, 255, 0, 0, 255
            Case 1
                Form1.Picture1.FillColor = RGB(255, 255, 0)
                Form1.Picture1.Circle (Bg(i).X, Bg(i).Y), Bg(i).Ar, RGB(255, 255, 0)
                If Local_State = 1 Then Server_SendData_Circle 0, Bg(i).X, Bg(i).Y, Bg(i).Ar, 255, 255, 0, 255, 255, 0
                Form1.Picture1.FillColor = RGB(0, 0, 255)
            Case 2
                Form1.Picture1.FillColor = RGB(255, 20, 147)
                Form1.Picture1.Circle (Bg(i).X, Bg(i).Y), Bg(i).Ar, RGB(255, 20, 147)
                If Local_State = 1 Then Server_SendData_Circle 0, Bg(i).X, Bg(i).Y, Bg(i).Ar, 255, 20, 147, 255, 20, 147
                Form1.Picture1.FillColor = RGB(0, 0, 255)
            Case 3
                Form1.Picture1.Line (Bg(i).X, Bg(i).Y)-(Pg(Bg(i).Target).X, Pg(Bg(i).Target).Y), RGB(225, 255, 255)
                If Local_State = 1 Then Server_SendData_Line Bg(i).X, Bg(i).Y, Pg(Bg(i).Target).X, Pg(Bg(i).Target).Y, 225, 255, 255
            Case 5
                Form1.Picture1.Line (Bg(i).X, 0)-(Bg(i).X, 8000), RGB(255, 255 * TeleportingFoePlaneLockTime, 255 * TeleportingFoePlaneLockTime)
                If Local_State = 1 Then Server_SendData_Line Bg(i).X, 0, Bg(i).X, 8000, 255, 255 * TeleportingFoePlaneLockTime, 255 * TeleportingFoePlaneLockTime
                Form1.Picture1.Line (0, Bg(i).Y)-(6000, Bg(i).Y), RGB(255, 255 * TeleportingFoePlaneLockTime, 255 * TeleportingFoePlaneLockTime)
                If Local_State = 1 Then Server_SendData_Line 0, Bg(i).Y, 6000, Bg(i).Y, 255, 255 * TeleportingFoePlaneLockTime, 255 * TeleportingFoePlaneLockTime
                
                Form1.Picture1.Line (Bg(i).X - 7000, Bg(i).Y - 7000)-(Bg(i).X + 7000, Bg(i).Y + 7000), RGB(255, 255 * TeleportingFoePlaneLockTime, 255 * TeleportingFoePlaneLockTime)
                If Local_State = 1 Then Server_SendData_Line Bg(i).X - 7000, Bg(i).Y - 7000, Bg(i).X + 7000, Bg(i).Y + 7000, 255, 255 * TeleportingFoePlaneLockTime, 255 * TeleportingFoePlaneLockTime
                Form1.Picture1.Line (Bg(i).X - 7000, Bg(i).Y + 7000)-(Bg(i).X + 7000, Bg(i).Y - 7000), RGB(255, 255 * TeleportingFoePlaneLockTime, 255 * TeleportingFoePlaneLockTime)
                If Local_State = 1 Then Server_SendData_Line Bg(i).X - 7000, Bg(i).Y + 7000, Bg(i).X + 7000, Bg(i).Y - 7000, 255, 255 * TeleportingFoePlaneLockTime, 255 * TeleportingFoePlaneLockTime
        End Select
    End If
Next
End Function
Public Function F5_PC_Plane()
Form1.Picture1.FillColor = RGB(255, 0, 0)
If Pg(0).A = True Then
    Form1.Picture1.Circle (Pg(0).X, Pg(0).Y), Pg(0).Ar, RGB(255, 0, 0)
    If Local_State = 1 Then Server_SendData_Circle 0, Pg(0).X, Pg(0).Y, Pg(0).Ar, 255, 0, 0, 255, 0, 0
End If
If Pg(1).A = True Then
    Form1.Picture1.FillColor = RGB(0, 100, 0)
    Form1.Picture1.Circle (Pg(1).X, Pg(1).Y), Pg(1).Ar, RGB(0, 100, 0)
    If Local_State = 1 Then Server_SendData_Circle 0, Pg(1).X, Pg(1).Y, Pg(1).Ar, 0, 100, 0, 0, 100, 0
End If
End Function
Public Sub World_Load()
Dim i As Long
ReDim Bg(1000)
Erase PSkill, PBg, Sg, FPg
Diff = 0: BgSum = 0: SgSum = 0: PBSum = 0: FPSum = 0
For i = 0 To 1
    PSkillID(i) = 0
    PSkill(i, 0) = True
    PBSkillCD(i) = True
    PBCD(i) = True
    Form1.Shape1(i).Left = Form1.SkOn1(PSkillID(i) + i * 5).Left: Form1.Shape1(i).Top = Form1.SkOn1(PSkillID(i) + i * 5).Top - 20
Next
Form1.Timer5.Interval = 1000
Form1.Label2.Caption = "0"
PC_Def
End Sub
