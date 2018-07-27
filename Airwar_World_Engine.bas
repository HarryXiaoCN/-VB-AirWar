Attribute VB_Name = "Airwar_World_Engine"
Public Function F5()
Form1.Picture1.Cls
物理事件检测
移动事件处理
F5_PC_Plane
F5_Foe_Plane
F5_Foe_Bullet
F5_PC_Supply
F5_PC_Bullet

End Function
Public Function F5_Foe_Plane()
Form1.Picture1.FillColor = RGB(143, 188, 143)
For i = 0 To FPSum - 1
    If FPg(i).a = True Then
        Select Case FPg(i).AiRank
            Case 0
                Form1.Picture1.Circle (FPg(i).X, FPg(i).Y), FPg(i).Ar, RGB(143, 188, 143)
            Case 1
                Form1.Picture1.FillColor = RGB(123, 104, 238)
                Form1.Picture1.Circle (FPg(i).X, FPg(i).Y), FPg(i).Ar, RGB(123, 104, 238)
                Form1.Picture1.FillColor = RGB(143, 188, 143)
            Case 2
                Form1.Picture1.FillColor = RGB(255, 69, 0)
                Form1.Picture1.Circle (FPg(i).X, FPg(i).Y), FPg(i).Ar, RGB(255, 69, 0)
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
            Case 1
                Form1.Picture1.FillColor = RGB(0, 0, 0)
                If PBg(i).X = PBg(i).mX And PBg(i).Y = PBg(i).mY Then
                    For j = 100 To 3000 Step 50
                        Form1.Picture1.Circle (PBg(i).X, PBg(i).Y), j, RGB(0, 0, 0)
                        DoEvents
                    Next
                Else
                    Form1.Picture1.Circle (PBg(i).X, PBg(i).Y), PBg(i).Ar, RGB(0, 0, 0)
                End If
                Form1.Picture1.FillColor = RGB(0, 255, 255)
                DoEvents
            Case 2
                Randomize
                Form1.Picture1.FillStyle = 1
                Form1.Picture1.Circle (PBg(i).X, PBg(i).Y), PBg(i).Ar, RGB(100, 149, 237)
                Form1.Picture1.Circle (PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20), PBg(i).Ar + Rnd * 20 - 20, RGB(176, 196, 222)
                Form1.Picture1.Circle (PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20), PBg(i).Ar + Rnd * 20 - 20, RGB(176, 196, 222)
                Form1.Picture1.Circle (PBg(i).X + Rnd * 20 - 20, PBg(i).Y + Rnd * 20 - 20), PBg(i).Ar + Rnd * 20 - 20, RGB(176, 196, 222)
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
            Case 1
                Form1.Picture1.Circle (Sg(i).X, Sg(i).Y), 50, RGB(0, 0, 255)
            Case 2
                Form1.Picture1.Circle (Sg(i).X, Sg(i).Y), 50, &H8000000D
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
            Case 2
                Form1.Picture1.FillColor = RGB(255, 255, 0)
                Form1.Picture1.Circle (Bg(i).X, Bg(i).Y), Bg(i).Ar, RGB(255, 255, 0)
                Form1.Picture1.FillColor = RGB(0, 0, 255)
            Case 3
                Form1.Picture1.FillColor = RGB(255, 20, 147)
                Form1.Picture1.Circle (Bg(i).X, Bg(i).Y), Bg(i).Ar, RGB(255, 20, 147)
                Form1.Picture1.FillColor = RGB(0, 0, 255)
        End Select
    End If
Next
End Function
Public Function F5_PC_Plane()
Form1.Picture1.FillColor = RGB(255, 0, 0)
For i = 0 To 1
    If Pg(i).a = True Then
        Form1.Picture1.Circle (Pg(i).X, Pg(i).Y), Pg(i).Ar, RGB(255, 0, 0)
    Else
        Exit For
    End If
Next
End Function
Public Sub World_Load()
Erase PSkill
Diff = 0: PBSkillCD = True: Timer5.Interval = 1000
Form1.Label2.Caption = "0"
PBCD = True
PC_Def
End Sub
