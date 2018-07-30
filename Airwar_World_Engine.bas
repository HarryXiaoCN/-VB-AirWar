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
If Local_State = 1 Then Server_SendData 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3
End Function
Public Function F5()
Select Case Local_State
    Case 0
        Form1.Picture1.Cls
        物理事件检测
        移动事件处理
        F5_PC_Plane
        F5_Foe_Plane
        F5_Foe_Bullet
        F5_PC_Supply
        F5_PC_Bullet
    Case 1
        Form1.Picture1.Cls
        Server_SendData 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1
        物理事件检测
        移动事件处理
        F5_NetworkMode
        F5_PC_Plane
        F5_Foe_Plane
        F5_Foe_Bullet
        F5_PC_Supply
        F5_PC_Bullet
End Select
End Function
Public Function F5_NetworkMode()
Dim i As Long
For i = 0 To 1
    Server_SendData 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, i, Form1.Label9(i).Caption, Form1.JinDuT1(i).Progress, Form1.JinDuT2(i).Progress, Form1.JinDuT3(i).Progress, Form1.Label8(i).Caption, PSkillID(i)
Next
End Function
Public Function F5_Foe_Plane()
Form1.Picture1.FillColor = RGB(143, 188, 143)
For i = 0 To FPSum - 1
    If FPg(i).a = True Then
        Select Case FPg(i).AiRank
            Case 0
                Form1.Picture1.Circle (FPg(i).X, FPg(i).Y), FPg(i).Ar, RGB(143, 188, 143)
                If Local_State = 1 Then Server_SendData 0, FPg(i).X, FPg(i).Y, FPg(i).Ar, 143, 188, 143, 143, 188, 143
            Case 1
                Form1.Picture1.FillColor = RGB(123, 104, 238)
                Form1.Picture1.Circle (FPg(i).X, FPg(i).Y), FPg(i).Ar, RGB(123, 104, 238)
                If Local_State = 1 Then Server_SendData 0, FPg(i).X, FPg(i).Y, FPg(i).Ar, 123, 104, 238, 123, 104, 238
                Form1.Picture1.FillColor = RGB(143, 188, 143)
            Case 2
                Form1.Picture1.FillColor = RGB(255, 69, 0)
                Form1.Picture1.Circle (FPg(i).X, FPg(i).Y), FPg(i).Ar, RGB(255, 69, 0)
                If Local_State = 1 Then Server_SendData 0, FPg(i).X, FPg(i).Y, FPg(i).Ar, 255, 69, 0, 255, 69, 0
                Form1.Picture1.FillColor = RGB(143, 188, 143)
        End Select
    End If
Next
End Function
Public Function F5_PC_Bullet()
Form1.Picture1.FillColor = RGB(0, 255, 255)
For i = 0 To PBSum - 1
    If PBg(i).a = True Then
        Select Case PBg(i).Trl
            Case 0
                Form1.Picture1.Circle (PBg(i).X, PBg(i).Y), PBg(i).Ar, RGB(0, 255, 255)
                If Local_State = 1 Then Server_SendData 0, PBg(i).X, PBg(i).Y, PBg(i).Ar, 0, 255, 255, 0, 255, 255
            Case 1
                Form1.Picture1.FillColor = RGB(0, 0, 0)
                If PBg(i).X = PBg(i).mX And PBg(i).Y = PBg(i).mY Then
                    For j = 100 To 3000 Step 50
                        Form1.Picture1.Circle (PBg(i).X, PBg(i).Y), j, RGB(0, 0, 0)
                        If Local_State = 1 Then Server_SendData 0, PBg(i).X, PBg(i).Y, j, 0, 0, 0, 0, 0, 0
                        DoEvents
                    Next
                Else
                    Form1.Picture1.Circle (PBg(i).X, PBg(i).Y), PBg(i).Ar, RGB(0, 0, 0)
                    If Local_State = 1 Then Server_SendData 0, PBg(i).X, PBg(i).Y, PBg(i).Ar, 0, 0, 0, 0, 0, 0
                End If
                Form1.Picture1.FillColor = RGB(0, 255, 255)
                DoEvents
            Case 2
                Randomize
                Form1.Picture1.FillStyle = 1
                Form1.Picture1.Circle (PBg(i).X, PBg(i).Y), PBg(i).Ar, RGB(100 + PSkillID_Ft(i) * 30, 149 + PSkillID_Ft(i) * 20, 237)
                If Local_State = 1 Then Server_SendData 1, PBg(i).X, PBg(i).Y, PBg(i).Ar, 100 + PSkillID_Ft(i) * 30, 149 + PSkillID_Ft(i) * 20, 237, 100 + PSkillID_Ft(i) * 30, 149 + PSkillID_Ft(i) * 20, 237
                Form1.Picture1.Circle (PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20), PBg(i).Ar + Rnd * 20 - 20, RGB(176, 196, 222)
                If Local_State = 1 Then Server_SendData 1, PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20, PBg(i).Ar + Rnd * 20 - 20, 176, 196, 222, 176, 196, 222
                Form1.Picture1.Circle (PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20), PBg(i).Ar + Rnd * 20 - 20, RGB(176, 196, 222)
                If Local_State = 1 Then Server_SendData 1, PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20, PBg(i).Ar + Rnd * 20 - 20, 176, 196, 222, 176, 196, 222
                Form1.Picture1.Circle (PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20), PBg(i).Ar + Rnd * 20 - 20, RGB(176, 196, 222)
                If Local_State = 1 Then Server_SendData 1, PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20, PBg(i).Ar + Rnd * 20 - 20, 176, 196, 222, 176, 196, 222
                Form1.Picture1.FillStyle = 0
        End Select
    End If
Next
End Function
Public Function F5_PC_Supply()
Form1.Picture1.FillColor = RGB(0, 255, 0)
For i = 0 To SgSum - 1
    If Sg(i).a = True Then
        Select Case Sg(i).Tp
            Case 0
                Form1.Picture1.Circle (Sg(i).X, Sg(i).Y), 50, RGB(255, 0, 0)
                If Local_State = 1 Then Server_SendData 0, Sg(i).X, Sg(i).Y, 50, 0, 255, 0, 255, 0, 0
            Case 1
                Form1.Picture1.Circle (Sg(i).X, Sg(i).Y), 50, RGB(0, 0, 255)
                If Local_State = 1 Then Server_SendData 0, Sg(i).X, Sg(i).Y, 50, 0, 255, 0, 0, 0, 255
            Case 2
                Form1.Picture1.Circle (Sg(i).X, Sg(i).Y), 50, RGB(30, 144, 255)
                If Local_State = 1 Then Server_SendData 0, Sg(i).X, Sg(i).Y, 50, 0, 255, 0, 30, 144, 255
        End Select
    End If
Next
End Function
Public Function F5_Foe_Bullet()
Form1.Picture1.FillColor = RGB(0, 0, 255)
For i = 0 To BgSum - 1
    If Bg(i).a = True Then
        Select Case Bg(i).Trl
            Case 0, 1
                Form1.Picture1.Circle (Bg(i).X, Bg(i).Y), Bg(i).Ar, RGB(0, 0, 255)
                If Local_State = 1 Then Server_SendData 0, Bg(i).X, Bg(i).Y, Bg(i).Ar, 0, 0, 255, 0, 0, 255
            Case 2
                Form1.Picture1.FillColor = RGB(255, 255, 0)
                Form1.Picture1.Circle (Bg(i).X, Bg(i).Y), Bg(i).Ar, RGB(255, 255, 0)
                If Local_State = 1 Then Server_SendData 0, Bg(i).X, Bg(i).Y, Bg(i).Ar, 255, 255, 0, 255, 255, 0
                Form1.Picture1.FillColor = RGB(0, 0, 255)
            Case 3
                Form1.Picture1.FillColor = RGB(255, 20, 147)
                Form1.Picture1.Circle (Bg(i).X, Bg(i).Y), Bg(i).Ar, RGB(255, 20, 147)
                If Local_State = 1 Then Server_SendData 0, Bg(i).X, Bg(i).Y, Bg(i).Ar, 255, 20, 147, 255, 20, 147
                Form1.Picture1.FillColor = RGB(0, 0, 255)
        End Select
    End If
Next
End Function
Public Function F5_PC_Plane()
Form1.Picture1.FillColor = RGB(255, 0, 0)
If Pg(0).a = True Then
    Form1.Picture1.Circle (Pg(0).X, Pg(0).Y), Pg(0).Ar, RGB(255, 0, 0)
    If Local_State = 1 Then Server_SendData 0, Pg(0).X, Pg(0).Y, Pg(0).Ar, 255, 0, 0, 255, 0, 0
End If
If Pg(1).a = True Then
    Form1.Picture1.FillColor = RGB(0, 100, 0)
    Form1.Picture1.Circle (Pg(1).X, Pg(1).Y), Pg(1).Ar, RGB(0, 100, 0)
    If Local_State = 1 Then Server_SendData 0, Pg(1).X, Pg(1).Y, Pg(1).Ar, 0, 100, 0, 0, 100, 0
End If
End Function
Public Sub World_Load()
Dim i As Long
Erase PSkill
Diff = 0
For i = 0 To 1
    PBSkillCD(i) = True
    PBCD(i) = True
Next
Form1.Shape1(0).Left = Form1.SkOn1(PSkillID(0)).Left: Form1.Shape1(0).Top = Form1.SkOn1(PSkillID(0)).Top - 20
    Form1.Shape1(1).Left = Form1.SkOn2(PSkillID(1)).Left: Form1.Shape1(1).Top = Form1.SkOn2(PSkillID(1)).Top - 20
Form1.Timer5.Interval = 1000
Form1.Label2.Caption = "0"
PC_Def
If DuoPlayer = True Then PC_2_Def
End Sub
