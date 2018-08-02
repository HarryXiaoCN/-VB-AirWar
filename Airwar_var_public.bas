Attribute VB_Name = "Airwar_var_public"
Public Pg(10) As Plane
Public Bg() As Bullet
Public PBg(1000) As Bullet
Public Sg(10) As Supply
Public FPg(1000) As Plane
Public PC(10) As KeyConfig
Public Diff, PBSkillCDTime(3) As Long
Public KCTemp(7) As Integer
Public PBCD(3), PBSkillCD(3), PSkill(2, 4), DuoPlayer, PlaneWYKZ_Skill_SwitchLock As Boolean
Public BgSum, SgSum, PBSum, FPSum, PSkillID(3) As Long
Public PSkillID_Ft(3) As Single
Public Function PBg_Add_2(ByVal PBID As Long, ByVal PID As Long, Optional Ty As Long = 0)
Select Case Ty
    Case 0
        PBg(PBID).a = True: PBg(PBID).Ar = 70: PBg(PBID).Atk = 10 + 2 * Pg(PID).Rank: PBg(PBID).Sb = True
        PBg(PBID).Sp = 50: PBg(PBID).X = Pg(PID).X: PBg(PBID).Y = Pg(PID).Y: PBg(PBID).PenHp = Pg(PID).Rank
        PBg(PBID).Pen = False: PBg(PBID).Trl = 0
        PBg(PBID).mX = Pg(PID).X
        PBg(PBID).mY = 8000
        Pg(PID).E = Pg(PID).E - 1
        PBCD(PID) = False
    Case 1
        PBg(PBID).a = True: PBg(PBID).Ar = 50: PBg(PBID).Atk = 100 + 20 * Pg(PID).Rank: PBg(PBID).Sb = True
        PBg(PBID).Sp = 20: PBg(PBID).X = Pg(PID).X: PBg(PBID).Y = Pg(PID).Y
        PBg(PBID).Pen = False: PBg(PBID).Trl = 1
        PBg(PBID).mX = Pg(PID).X
        PBg(PBID).mY = Pg(PID).Y + 3200
        Pg(PID).E = Pg(PID).E - 20
        PBSkillCD(PID) = False
    Case 2
        PBg(PBID).a = True: PBg(PBID).Ar = 500 + 10 * Pg(PID).Rank: PBg(PBID).Atk = 5 + 10 * Pg(PID).Rank: PBg(PBID).Sb = False
        PBg(PBID).Sp = 0: PBg(PBID).X = Pg(PID).X: PBg(PBID).Y = Pg(PID).Y
        PBg(PBID).Pen = False: PBg(PBID).Trl = 2
        PBg(PBID).mX = Pg(PID).X
        PBg(PBID).mY = Pg(PID).Y
        Pg(PID).E = Pg(PID).E - 40
        PBSkillCD(PID) = False
    Case 3
        PBg(PBID).a = True: PBg(PBID).Ar = 10 * Pg(PID).Rank: PBg(PBID).Atk = Pg(PID).MxHP * 0.007: PBg(PBID).Sb = False
        PBg(PBID).Sp = 0: PBg(PBID).X = Pg(PID).X: PBg(PBID).Y = Pg(PID).Y
        PBg(PBID).Pen = False: PBg(PBID).Trl = 3
        PBg(PBID).mX = Pg(PID).X
        PBg(PBID).mY = Pg(PID).Y
        Pg(PID).E = Pg(PID).E - 50
        PBSkillCD(PID) = False
End Select
PBg(PBID).Source = PID
End Function
Public Function PBg_Add(ByVal PID As Long, Optional Ty As Long = 0)
Select Case Ty
    Case 0
        If Pg(PID).E <= 0 Or PBCD(PID) = False Then Exit Function
    Case 1, 2, 3
        If PSkillID(PID) = 0 Then If Pg(PID).E < 20 Or PBSkillCD(PID) = False Then Exit Function
        If PSkillID(PID) = 1 Then If Pg(PID).E < 40 Or PBSkillCD(PID) = False Then Exit Function
        If PSkillID(PID) = 2 Then If Pg(PID).E < 50 Or PBSkillCD(PID) = False Then Exit Function
End Select
For i = 0 To PBSum - 1
    If PBg(i).a = False Then
        PBg_Add_2 i, PID, Ty
        Exit Function
    End If
Next
PBg_Add_2 PBSum, PID, Ty
PBSum = PBSum + 1
End Function
Public Function FPg_Add(ByVal FPID As Long, Optional Arnk As Long = 0)
Select Case Arnk
    Case 0
        FPg(FPID).a = True: FPg(FPID).HP = 10: FPg(FPID).MxHP = 10
        FPg(FPID).Ar = 100: FPg(FPID).Sb = True: FPg(FPID).AiRank = 0
        FPg(FPID).Sp = 5
        FPg_Add_0_Def FPID
    Case 1
        FPg(FPID).a = True: FPg(FPID).HP = 200: FPg(FPID).MxHP = 200
        FPg(FPID).Ar = 150: FPg(FPID).Sb = True: FPg(FPID).AiRank = 1
        FPg(FPID).Sp = 3
        FPg_Add_0_Def FPID
    Case 2
        FPg(FPID).a = True: FPg(FPID).HP = 50: FPg(FPID).MxHP = 50
        FPg(FPID).Ar = 80: FPg(FPID).Sb = True: FPg(FPID).AiRank = 2
        FPg(FPID).Sp = 8
        FPg_Add_0_Def FPID
    Case 3
        FPg(FPID).a = True: FPg(FPID).HP = 250: FPg(FPID).MxHP = 250
        FPg(FPID).Ar = 120: FPg(FPID).Sb = True: FPg(FPID).AiRank = 3
        FPg(FPID).Sp = 5
        FPg_Add_0_Def FPID
End Select
End Function
Public Function FPg_Add_0_Def(ByVal FPID As Long)
Randomize
Select Case Int(Rnd * 4)
    Case 0
        FPg(FPID).X = 0: FPg(FPID).Y = Int(Rnd * (8001))
        FPg(FPID).mX = 6000: FPg(FPID).mY = Int(Rnd * (8001))
    Case 1
        FPg(FPID).X = 6000: FPg(FPID).Y = Int(Rnd * (8001))
        FPg(FPID).mX = 0: FPg(FPID).mY = Int(Rnd * (8001))
    Case 2
        FPg(FPID).X = Int(Rnd * (6001)): FPg(FPID).Y = 0
        FPg(FPID).mX = Int(Rnd * (6001)): FPg(FPID).mY = 8000
    Case 3
        FPg(FPID).X = Int(Rnd * (6001)): FPg(FPID).Y = 8000
        FPg(FPID).mX = Int(Rnd * (6001)): FPg(FPID).mY = 0
End Select
End Function
Public Sub PC_Def()
Pg(0).a = True: Pg(0).HP = 100: Pg(0).MxHP = 100: Pg(0).MxEmp = 10
Pg(0).EMP = 0: Pg(0).Rank = 1: Pg(0).Blt = 1
Pg(0).Ar = 100: Pg(0).E = 100: Pg(0).MxE = 100: Pg(0).Sb = True
Pg(0).Sp = 30: Pg(0).Esp = 1: Pg(0).X = 2000: Pg(0).Y = 1000
End Sub
Public Sub PC_2_Def()
Pg(1).a = True: Pg(1).HP = 100: Pg(1).MxHP = 100: Pg(1).MxEmp = 10
Pg(1).EMP = 0: Pg(1).Rank = 1: Pg(1).Blt = 1
Pg(1).Ar = 100: Pg(1).E = 100: Pg(1).MxE = 100: Pg(1).Sb = True
Pg(1).Sp = 30: Pg(1).Esp = 1: Pg(1).X = 4000: Pg(1).Y = 1000
End Sub
Public Function Sg_Add(ByVal SgID As Long)
Randomize
Sg(SgID).a = True: Sg(SgID).Tp = Int(Rnd * 3): Sg(SgID).Sp = 20
Sg(SgID).X = Int(Rnd * (6001)): Sg(SgID).Y = 8000
Sg(SgID).mX = Int(Rnd * (6001)): Sg(SgID).mY = 0
End Function
Public Function Bg_Add_JustHitYou() As Long
Dim i, Aggregate, Division, Verification As Long
Randomize
For i = 0 To 1
    If Pg(i).a = True Then
        Aggregate = Aggregate + Pg(i).Rank
    End If
Next
Division = Int(Rnd * (Aggregate + 1))
For i = 0 To 1
    If Pg(i).a = True Then
        Verification = Verification + Pg(i).Rank
        If Verification > Division Then Bg_Add_JustHitYou = i: Exit Function
    End If
Next
End Function
Public Function Bg_Add(ByVal BgID As Long, Optional Ty As Long = 0, Optional FPID)
Randomize
Bg(BgID).Trl = Ty
Bg(BgID).Target = Bg_Add_JustHitYou
Select Case Ty
    Case 0
        Bg(BgID).a = True: Bg(BgID).Ar = 30: Bg(BgID).Atk = 10: Bg(BgID).Sb = True
        Bg(BgID).Sp = 10 + Diff / 30: Bg(BgID).X = Int(Rnd * (6001)): Bg(BgID).Y = 8000: Bg(BgID).mX = Pg(Bg(BgID).Target).X
        Bg(BgID).mY = Pg(Bg(BgID).Target).Y
    Case 1
        Bg(BgID).a = True: Bg(BgID).Ar = 30: Bg(BgID).Atk = 10: Bg(BgID).Sb = True
        Bg(BgID).Sp = 10 + Diff / 30: Bg(BgID).X = FPg(FPID).X
        Bg(BgID).Y = FPg(FPID).Y: Bg(BgID).mX = Pg(Bg(BgID).Target).X
        Bg(BgID).mY = Pg(Bg(BgID).Target).Y
    Case 2
        Bg(BgID).a = True: Bg(BgID).Ar = 40: Bg(BgID).Atk = 20: Bg(BgID).Sb = True
        Bg(BgID).Sp = 5 + Diff / 30: Bg(BgID).X = FPg(FPID).X
        Bg(BgID).Y = FPg(FPID).Y: Bg(BgID).mX = Pg(Bg(BgID).Target).X
        Bg(BgID).mY = Pg(Bg(BgID).Target).Y
    Case 3
        Bg(BgID).a = True: Bg(BgID).Ar = 20: Bg(BgID).Atk = 15: Bg(BgID).Sb = True
        Bg(BgID).Sp = 10 + Diff / 25: Bg(BgID).X = FPg(FPID).X
        Bg(BgID).Y = FPg(FPID).Y: Bg(BgID).mX = Pg(Bg(BgID).Target).X
        Bg(BgID).mY = Pg(Bg(BgID).Target).Y
    Case 4
        Bg(BgID).a = True: Bg(BgID).Ar = 10: Bg(BgID).Atk = 0.01: Bg(BgID).Sb = False
        Bg(BgID).Sp = 10 + Diff / 25: Bg(BgID).X = FPg(FPID).X
        Bg(BgID).Y = FPg(FPID).Y: Bg(BgID).mX = Pg(Bg(BgID).Target).X
        Bg(BgID).mY = Pg(Bg(BgID).Target).Y
        Bg(BgID).Source = FPID
End Select
End Function
