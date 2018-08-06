Attribute VB_Name = "Airwar_ruler_engine_Move"
Public Function FoePlaneWYKZ(ByRef FPID As Long)
Dim XXJie As 二元解
Select Case FPg(FPID).AiRank
    Case 0, 1, 2, 3, 4
        If FPg(FPID).Da = False Then
            XXJie = 线性求解(FPg(FPID).X, FPg(FPID).Y, FPg(FPID).mX, FPg(FPID).mY, FPg(FPID).Sp)
            FPg(FPID).dX = XXJie.a: FPg(FPID).dY = XXJie.b: FPg(FPID).Da = True
        End If
        FPg(FPID).X = FPg(FPID).X + FPg(FPID).dX: FPg(FPID).Y = FPg(FPID).Y + FPg(FPID).dY
        If FPg(FPID).X < 0 Or FPg(FPID).Y < 0 Or FPg(FPID).X > 6000 Or FPg(FPID).Y > 8000 Then
            FPg(FPID).a = False: FPg(FPID).Da = False
        End If
End Select
End Function
Public Function BulletWYKZ(ByRef BID As Long, TrID As Long)
Dim XXJie As 二元解
Select Case TrID
    Case 0, 1
        If Bg(BID).Da = False Then
            XXJie = 线性求解(Bg(BID).X, Bg(BID).Y, Bg(BID).mX, Bg(BID).mY, Bg(BID).Sp)
            Bg(BID).dX = XXJie.a: Bg(BID).dY = XXJie.b: Bg(BID).Da = True
        End If
        Bg(BID).X = Bg(BID).X + Bg(BID).dX: Bg(BID).Y = Bg(BID).Y + Bg(BID).dY
        If Bg(BID).X < 0 Or Bg(BID).Y < 0 Or Bg(BID).X > 6000 Or Bg(BID).Y > 8000 Then
            Bg(BID).a = False: Bg(BID).Da = False
        End If
    Case 2
        XXJie = 线性求解(Bg(BID).X, Bg(BID).Y, Pg(Bg(BID).Target).X, Pg(Bg(BID).Target).Y, Bg(BID).Sp)
        Bg(BID).dX = XXJie.a: Bg(BID).dY = XXJie.b
        Bg(BID).X = Bg(BID).X + Bg(BID).dX: Bg(BID).Y = Bg(BID).Y + Bg(BID).dY
        If Bg(BID).X < 0 Or Bg(BID).Y < 0 Or Bg(BID).X > 6000 Or Bg(BID).Y > 8000 Then
            Bg(BID).a = False: Bg(BID).Da = False
        End If
    Case 3, 4
        Bg(BID).X = FPg(Bg(BID).Source).X: Bg(BID).Y = FPg(Bg(BID).Source).Y
        If FPg(Bg(BID).Source).a = False Then Bg(BID).a = False
End Select
End Function
Public Function PlBtWYKZ(ByRef PBID As Long, TrID As Long)
Dim XXJie As 二元解
Select Case TrID
    Case 0
        If PBg(PBID).Da = False Then
            XXJie = 线性求解(PBg(PBID).X, PBg(PBID).Y, PBg(PBID).mX, PBg(PBID).mY, PBg(PBID).Sp)
            PBg(PBID).dX = XXJie.a: PBg(PBID).dY = XXJie.b: PBg(PBID).Da = True
        End If
        PBg(PBID).X = PBg(PBID).X + PBg(PBID).dX: PBg(PBID).Y = PBg(PBID).Y + PBg(PBID).dY
        If PBg(PBID).X < 0 Or PBg(PBID).Y < 0 Or PBg(PBID).X > 6000 Or PBg(PBID).Y > 8000 Then
            PBg(PBID).a = False: PBg(PBID).Da = False
        End If
    Case 1
        If PBg(PBID).Y < PBg(PBID).mY Then
            If PBg(PBID).Da = False Then
                XXJie = 线性求解(PBg(PBID).X, PBg(PBID).Y, PBg(PBID).mX, PBg(PBID).mY, PBg(PBID).Sp)
                PBg(PBID).dX = XXJie.a: PBg(PBID).dY = XXJie.b: PBg(PBID).Da = True
            End If
            PBg(PBID).X = PBg(PBID).X + PBg(PBID).dX: PBg(PBID).Y = PBg(PBID).Y + PBg(PBID).dY
        Else
            '----------------------------
            If Form1.BHB(PBg(PBID).Source).Enabled = False Then
                PSkillID_BHB(PBg(PBID).Source) = PBg(PBID).Ar
                Form1.BHB((PBg(PBID).Source)).Enabled = True
            Else
                If PBg(PBID).Ar > 3000 Then
                    Form1.BHB((PBg(PBID).Source)).Enabled = False
                    PBg(PBID).a = False: PBg(PBID).Da = False
                Else
                    PBg(PBID).Ar = PSkillID_BHB(PBg(PBID).Source)
                End If
            End If
            '----------------------------
        End If
'        If PBg(PBID).X = PBg(PBID).mX And PBg(PBID).Y >= PBg(PBID).mY Then PBg(PBID).Ar = 3000
'        If PBg(PBID).X < 0 Or PBg(PBID).Y < 0 Or PBg(PBID).X > 6000 Or PBg(PBID).Y > 11000 Then
'            PBg(PBID).a = False: PBg(PBID).Da = False
'        End If
    Case 2
        If Form1.Ftime(PBg(PBID).Source).Enabled = False Then
            PSkillID_Ft(PBg(PBID).Source) = 0: Form1.Ftime((PBg(PBID).Source)).Enabled = True: PBg(PBID).X = Pg((PBg(PBID).Source)).X: PBg(PBID).Y = Pg((PBg(PBID).Source)).Y
        Else
            If PSkillID_Ft(PBg(PBID).Source) > 3 + 0.5 * Pg((PBg(PBID).Source)).Rank Then
                Form1.Ftime((PBg(PBID).Source)).Enabled = False
                PBg(PBID).a = False
            Else
                PBg(PBID).X = Pg((PBg(PBID).Source)).X: PBg(PBID).Y = Pg((PBg(PBID).Source)).Y
            End If
        End If
    Case 3
        If Form1.Ftime(PBg(PBID).Source).Enabled = False Then
            PSkillID_Ft(PBg(PBID).Source) = 0: Form1.Ftime((PBg(PBID).Source)).Enabled = True: PBg(PBID).X = Pg((PBg(PBID).Source)).X: PBg(PBID).Y = Pg((PBg(PBID).Source)).Y
        Else
            If PSkillID_Ft(PBg(PBID).Source) > 1 + 0.1 * Pg((PBg(PBID).Source)).Rank Then
                Form1.Ftime((PBg(PBID).Source)).Enabled = False
                PBg(PBID).a = False
            Else
                PBg(PBID).X = Pg((PBg(PBID).Source)).X: PBg(PBID).Y = Pg((PBg(PBID).Source)).Y
                PBg(PBID).Ar = PSkillID_Ft(PBg(PBID).Source) * 700
            End If
        End If
    Case 4
        If PBg(PBID).Y < PBg(PBID).mY Then
            If PBg(PBID).Da = False Then
                XXJie = 线性求解(PBg(PBID).X, PBg(PBID).Y, PBg(PBID).mX, PBg(PBID).mY, PBg(PBID).Sp)
                PBg(PBID).dX = XXJie.a: PBg(PBID).dY = XXJie.b: PBg(PBID).Da = True
            End If
            PBg(PBID).X = PBg(PBID).X + PBg(PBID).dX: PBg(PBID).Y = PBg(PBID).Y + PBg(PBID).dY
        Else
            '----------------------------
            If Form1.Ftime(PBg(PBID).Source).Enabled = False Then
                PSkillID_Ft(PBg(PBID).Source) = 0
                Form1.Ftime((PBg(PBID).Source)).Enabled = True
            Else
                If PSkillID_Ft(PBg(PBID).Source) > 5 + Pg((PBg(PBID).Source)).Rank Then
                    Form1.Ftime((PBg(PBID).Source)).Enabled = False
                    PBg(PBID).a = False: PBg(PBID).Da = False
                End If
            End If
            '----------------------------
        End If
End Select
End Function
Public Function PlaneWYKZ_Up(ByRef PlID As Long)
If Pg(PlID).Y + Pg(PlID).Sp + Pg(PlID).Ar >= Form1.Picture1.Height Then Pg(PlID).Y = Form1.Picture1.Height - Pg(PlID).Ar: Exit Function
If Pg(PlID).Y + Pg(PlID).Sp + Pg(PlID).Ar <= 0 Then Pg(PlID).Y = Pg(PlID).Ar: Exit Function
Pg(PlID).Y = Pg(PlID).Y + Pg(PlID).Sp
'If Pg(PlID).Y > Form1.Picture1.Height Then Pg(PlID).Y = Form1.Picture1.Height
End Function
Public Function PlaneWYKZ_Down(ByRef PlID As Long)
If Pg(PlID).Y - Pg(PlID).Sp - Pg(PlID).Ar <= 0 Then Pg(PlID).Y = Pg(PlID).Ar: Exit Function
If Pg(PlID).Y - Pg(PlID).Sp - Pg(PlID).Ar >= Form1.Picture1.Height Then Pg(PlID).Y = Form1.Picture1.Height - Pg(PlID).Ar: Exit Function
Pg(PlID).Y = Pg(PlID).Y - Pg(PlID).Sp
'If Pg(PlID).Y < 0 Then Pg(PlID).Y = 0
End Function
Public Function PlaneWYKZ_Left(ByRef PlID As Long)
If Pg(PlID).X - Pg(PlID).Sp - Pg(PlID).Ar <= 0 Then Pg(PlID).X = Pg(PlID).Ar: Exit Function
If Pg(PlID).X - Pg(PlID).Sp - Pg(PlID).Ar >= Form1.Picture1.Width Then Pg(PlID).X = Form1.Picture1.Width - Pg(PlID).Ar: Exit Function
Pg(PlID).X = Pg(PlID).X - Pg(PlID).Sp
'If Pg(PlID).X < 0 Then Pg(PlID).X = 0
End Function
Public Function PlaneWYKZ_Right(ByRef PlID As Long)
If Pg(PlID).X + Pg(PlID).Sp + Pg(PlID).Ar >= Form1.Picture1.Width Then Pg(PlID).X = Form1.Picture1.Width - Pg(PlID).Ar: Exit Function
If Pg(PlID).X + Pg(PlID).Sp + Pg(PlID).Ar <= 0 Then Pg(PlID).X = Pg(PlID).Ar: Exit Function
Pg(PlID).X = Pg(PlID).X + Pg(PlID).Sp
'If Pg(PlID).X >= Form1.Picture1.Width Then Pg(PlID).X = Form1.Picture1.Width
End Function
Public Function PlaneWYKZ_Skill_Switch(ByRef PlID As Long)
If PlaneWYKZ_Skill_SwitchLock = True Then Form1.ChangeLock = True: Exit Function
If PSkill(PlID, PSkillID(PlID) + 1) = True Then
    PSkillID(PlID) = PSkillID(PlID) + 1: Pg(PlID).Blt = PSkillID(PlID) + 1
Else
    PSkillID(PlID) = 0: Pg(PlID).Blt = 1
End If
Form1.Shape1(PlID).Left = Form1.SkOn1(PSkillID(PlID) + PlID * 5).Left: Form1.Shape1(PlID).Top = Form1.SkOn1(PSkillID(PlID) + PlID * 5).Top - 20
PlaneWYKZ_Skill_SwitchLock = True
End Function
Public Function PlaneWYKZ(ByRef PlID As Long, KeyCode As Integer, KeyBtye As Long)
Select Case KeyCode
    Case PC(PlID).Left
        PlaneWYKZ_Left PlID
    Case PC(PlID).Up
        PlaneWYKZ_Up PlID
    Case PC(PlID).Right
        PlaneWYKZ_Right PlID
    Case PC(PlID).Down
        PlaneWYKZ_Down PlID
    Case PC(PlID).Attack
        PBg_Add PlID
    Case PC(PlID).Ultimate_Skill
        PBg_Add PlID, Pg(PlID).Blt
    Case PC(PlID).Skill_Switch
        PlaneWYKZ_Skill_Switch PlID
End Select
End Function
Public Function SupplyWYKZ(ByRef BID As Long)
Dim XXJie As 二元解
If Sg(SID).Da = False Then
    XXJie = 线性求解(Sg(SID).X, Sg(SID).Y, Sg(SID).mX, Sg(SID).mY, Sg(SID).Sp)
    Sg(SID).dX = XXJie.a: Sg(SID).dY = XXJie.b: Sg(SID).Da = True
End If
Sg(SID).X = Sg(SID).X + Sg(SID).dX: Sg(SID).Y = Sg(SID).Y + Sg(SID).dY
If Sg(SID).X < 0 Or Sg(SID).Y < 0 Or Sg(SID).X > 6000 Or Sg(SID).Y > 8000 Then
    Sg(SID).a = False: Sg(SID).Da = False
End If
End Function
