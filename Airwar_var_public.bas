Attribute VB_Name = "Airwar_var_public"
Public Pg(1000) As Plane
Public Bg(1000) As Bullet
Public PBg(1000) As Bullet
Public Sg(10000) As Supply
Public FPg(1000) As Plane
Public Diff, PBSkillCDTime As Long
Public KCTemp(5) As Integer
Public PBCD, PBSkillCD, PSkill(3) As Boolean
Public BgSum, SgSum, PBSum, FPSum, PSkillID, PSkillID_Ft As Long
Public Function PBg_Add_2(ByVal PBID As Long, ByVal PID As Long, Optional Ty As Long = 0)
Select Case Ty
    Case 0
        PBg(PBID).a = True: PBg(PBID).Ar = 70: PBg(PBID).Atk = 10 + 2 * Pg(PID).Rank: PBg(PBID).Sb = True
        PBg(PBID).Sp = 50: PBg(PBID).X = Pg(PID).X: PBg(PBID).Y = Pg(PID).Y
        PBg(PBID).Pen = False: PBg(PBID).Trl = 0
        PBg(PBID).mX = Pg(PID).X
        PBg(PBID).mY = 8000
        Pg(PID).E = Pg(PID).E - 1
        PBCD = False
    Case 1
        PBg(PBID).a = True: PBg(PBID).Ar = 50: PBg(PBID).Atk = 100 + 20 * Pg(PID).Rank: PBg(PBID).Sb = True
        PBg(PBID).Sp = 20: PBg(PBID).X = Pg(PID).X: PBg(PBID).Y = Pg(PID).Y
        PBg(PBID).Pen = False: PBg(PBID).Trl = 1
        PBg(PBID).mX = Pg(PID).X
        PBg(PBID).mY = Pg(PID).Y + 3200
        Pg(PID).E = Pg(PID).E - 20
        PBSkillCD = False
    Case 2
        PBg(PBID).a = True: PBg(PBID).Ar = 300: PBg(PBID).Atk = 5 + 10 * Pg(PID).Rank: PBg(PBID).Sb = False
        PBg(PBID).Sp = 0: PBg(PBID).X = Pg(PID).X: PBg(PBID).Y = Pg(PID).Y
        PBg(PBID).Pen = False: PBg(PBID).Trl = 2
        PBg(PBID).mX = Pg(PID).X
        PBg(PBID).mY = Pg(PID).Y
        Pg(PID).E = Pg(PID).E - 50
        PBSkillCD = False
End Select
End Function
Public Function PBg_Add(ByVal PID As Long, Optional Ty As Long = 0)
Select Case Ty
    Case 0
        If Pg(PID).E <= 0 Or PBCD = False Then Exit Function
    Case 1, 2
        If Pg(PID).E < 20 Or PBSkillCD = False Then Exit Function
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
Pg(0).Emp = 0: Pg(0).Rank = 1: Pg(0).Blt = 1
Pg(0).Ar = 100: Pg(0).E = 100: Pg(0).MxE = 100: Pg(0).Sb = True
Pg(0).Sp = 30: Pg(0).Esp = 1: Pg(0).X = 3000: Pg(0).Y = 1000
End Sub
Public Function Sg_Add(ByVal SgID As Long)
Randomize
Sg(SgID).a = True: Sg(SgID).Tp = Int(Rnd * 3): Sg(SgID).Sp = 20
Sg(SgID).X = Int(Rnd * (6001)): Sg(SgID).Y = 8000
Sg(SgID).mX = Int(Rnd * (6001)): Sg(SgID).mY = 0
End Function
Public Function Bg_Add(ByVal BgID As Long, Optional Ty As Long = 0, Optional FPID)
Randomize
Bg(BgID).Trl = Ty
Select Case Ty
    Case 0
        Bg(BgID).a = True: Bg(BgID).Ar = 30: Bg(BgID).Atk = 10: Bg(BgID).Sb = True
        Bg(BgID).Sp = 10 + Diff / 30: Bg(BgID).X = Int(Rnd * (6001)): Bg(BgID).Y = 8000: Bg(BgID).mX = Pg(0).X
        Bg(BgID).mY = Pg(0).Y
    Case 1
        Bg(BgID).a = True: Bg(BgID).Ar = 30: Bg(BgID).Atk = 10: Bg(BgID).Sb = True
        Bg(BgID).Sp = 10 + Diff / 30: Bg(BgID).X = FPg(FPID).X
        Bg(BgID).Y = FPg(FPID).Y: Bg(BgID).mX = Pg(0).X
        Bg(BgID).mY = Pg(0).Y
    Case 2
        Bg(BgID).a = True: Bg(BgID).Ar = 40: Bg(BgID).Atk = 20: Bg(BgID).Sb = True
        Bg(BgID).Sp = 5 + Diff / 30: Bg(BgID).X = FPg(FPID).X
        Bg(BgID).Y = FPg(FPID).Y: Bg(BgID).mX = Pg(0).X
        Bg(BgID).mY = Pg(0).Y
    Case 3
        Bg(BgID).a = True: Bg(BgID).Ar = 20: Bg(BgID).Atk = 15: Bg(BgID).Sb = True
        Bg(BgID).Sp = 10 + Diff / 25: Bg(BgID).X = FPg(FPID).X
        Bg(BgID).Y = FPg(FPID).Y: Bg(BgID).mX = Pg(0).X
        Bg(BgID).mY = Pg(0).Y
End Select
End Function
