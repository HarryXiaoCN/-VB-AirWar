Attribute VB_Name = "Airwar_var_public"
Public Pg(10) As Plane
Public Bg() As Bullet
Public PBg(1000) As Bullet
Public Sg(10) As Supply
Public FPg(1000) As Plane
Public PC(10) As KeyConfig
Public Diff, PBSkillCDTime(3) As Long 'Diff 是难度，随时间变化；PBSkillCDTime 是技能冷却时间
Public KCTemp(7) As Integer
Public PBCD(3), PBSkillCD(3), PSkill(2, 5), DuoPlayer, PlaneWYKZ_Skill_SwitchLock As Boolean
'PBCD(玩家号) 控制玩家子弹频率；PBSkillCD 由PBSkillCDTime决定的大招开关
'PSkill（玩家号，技能号） 玩家大招有无控制；DuoPlayer 双人玩家控制
'PlaneWYKZ_Skill_SwitchLock 大招切换间隔控制
Public KeyboardVis, ReinCodeVis, TeleportingFoePlaneDisplacementLock As Boolean '键盘及手柄可视化控制
Public PgSum, BgSum, SgSum, PBSum, FPSum, PSkillID(3) As Long 'PSkillID(玩家号) 当前玩家大招指向
Public PSkillID_Ft(3), PSkillID_BHB(3), T_s, TeleportingFoePlaneLockTime As Single 'T_s 游戏计时变量
Public Sub PC_Def()
'[Engine]World_Load input
Dim i As Long
For i = 0 To PgSum
    Pg_Def_AttributeAssignment i
Next
End Sub
Public Function PBg_Add(ByVal PID As Long, Optional Ty As Long = 0)
'[Move] PlaneWYKZ input
Select Case Ty
    Case 0
        If Pg(PID).E < 1 Or PBCD(PID) = False Then Exit Function
    Case 1, 2, 3, 4
        If PSkillID(PID) = 0 Then If Pg(PID).E < 10 Or PBSkillCD(PID) = False Then Exit Function
        If PSkillID(PID) = 1 Then If Pg(PID).E < 20 Or PBSkillCD(PID) = False Then Exit Function
        If PSkillID(PID) = 2 Then If Pg(PID).E < 30 Or PBSkillCD(PID) = False Then Exit Function
        If PSkillID(PID) = 3 Then If Pg(PID).E < 40 Or PBSkillCD(PID) = False Then Exit Function
End Select
For i = 0 To PBSum - 1
    If PBg(i).a = False Then
        PBg_Add_ClassificationAndEntry i, PID, Ty
        Exit Function
    End If
Next
PBg_Add_ClassificationAndEntry PBSum, PID, Ty
PBSum = PBSum + 1
End Function
Public Function PBg_Add_ClassificationAndEntry(ByVal PBID As Long, ByVal PID As Long, Optional Ty As Long = 0)
'PBg_Add input
Select Case Ty
    Case 0 '普通子弹
        PBg_add_AttributeAssignment PBID, PID, 70, 10 + 2 * Pg(PID).Rank, True, 50, Pg(PID).X, Pg(PID).Y, Pg(PID).Rank, False, 0, Pg(PID).X, 8000, 1
    Case 1 '核弹
        PBg_add_AttributeAssignment PBID, PID, 50, 100 + 20 * Pg(PID).Rank, False, 20, Pg(PID).X, Pg(PID).Y, Pg(PID).Rank, False, 1, Pg(PID).X, Pg(PID).Y + 3200, 10
    Case 2 '力场护盾
        PBg_add_AttributeAssignment PBID, PID, 500 + 10 * Pg(PID).Rank, 5 + 10 * Pg(PID).Rank, False, 0, Pg(PID).X, Pg(PID).Y, Pg(PID).Rank, False, 2, Pg(PID).X, Pg(PID).Y, 20
    Case 3 '治疗力场
        PBg_add_AttributeAssignment PBID, PID, 10 * Pg(PID).Rank, Pg(PID).MxHP * 0.007, False, 0, Pg(PID).X, Pg(PID).Y, Pg(PID).Rank, False, 3, Pg(PID).X, Pg(PID).Y, 30
    Case 4 '黑洞炸弹
        PBg_add_AttributeAssignment PBID, PID, 50, Pg(PID).Rank, False, 30, Pg(PID).X, Pg(PID).Y, Pg(PID).Rank, False, 4, Pg(PID).X, Pg(PID).Y + 3200, 40
End Select
PBg(PBID).Source = PID
End Function
Public Function FPg_Add(ByVal FPID As Long, Optional Arnk As Long = 0)
'[Case]Test_FoePlane input
FPg(FPID).AiRank = Arnk
Select Case Arnk
    Case 0 '小飞机
        FPg_add_AttributeAssignment FPID, 10 + Diff / 10, 100, True, 5
    Case 1 '大飞机
        FPg_add_AttributeAssignment FPID, 200, 150, True, 3
    Case 2 '跟踪飞机
        FPg_add_AttributeAssignment FPID, 50, 80, True, 8
    Case 3 '冰冻飞机
        FPg_add_AttributeAssignment FPID, 250, 120, True, 5
    Case 4 '黑洞飞机
        FPg_add_AttributeAssignment FPID, 2000, 500, True, 3
    Case 5 '瞬移飞机
        FPg_add_AttributeAssignment FPID, 3000, 200, True, 0
End Select
FPg_Add_TrajectoryInitialization FPID, Arnk
End Function
Public Function FPg_Add_TrajectoryInitialization(ByVal FPID As Long, Optional Arnk As Long = 0)
'FPg_Add input
Randomize
Select Case Arnk
    Case 0, 1, 2, 3, 4 '四周随机出现
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
    Case 5
        FPg(FPID).X = Int(Rnd * (4001)) + 1000
        FPg(FPID).Y = Int(Rnd * (6001)) + 1000
End Select
End Function
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
Public Function Bg_Add(ByRef BID As Long, ByRef Arnk As Long, Optional FPID)
'[Case]FoeBullet_Shoot input
Bg(BID).Trl = Arnk
Bg(BID).Target = Bg_Add_JustHitYou
Bg(BID).Source = FPID
Select Case Arnk
    Case 0 '小飞机子弹
        Bg_add_AttributeAssignment BID, 30, 10 + Diff / 10, True, 10 + Diff / 30, FPg(FPID).X, FPg(FPID).Y, Pg(Bg(BID).Target).X, Pg(Bg(BID).Target).Y
    Case 1 '大飞机子弹
        Bg_add_AttributeAssignment BID, 40, 20, True, 5 + Diff / 30, FPg(FPID).X, FPg(FPID).Y, Pg(Bg(BID).Target).X, Pg(Bg(BID).Target).Y
    Case 2 '跟踪弹
        Bg_add_AttributeAssignment BID, 20, 15, True, 10 + Diff / 25, FPg(FPID).X, FPg(FPID).Y, Pg(Bg(BID).Target).X, Pg(Bg(BID).Target).Y
    Case 3 '冰冻射线
        Bg_add_AttributeAssignment BID, 1, 0.01, True, 0, FPg(FPID).X, FPg(FPID).Y, Pg(Bg(BID).Target).X, Pg(Bg(BID).Target).Y
    Case 4 '引力撕裂
        Bg_add_AttributeAssignment BID, 1, 10, True, 10, FPg(FPID).X, FPg(FPID).Y, Pg(Bg(BID).Target).X, Pg(Bg(BID).Target).Y
    Case 5 '纵横激光炮
        Bg_add_AttributeAssignment BID, 1, 3, True, 10, FPg(FPID).X, FPg(FPID).Y, Pg(Bg(BID).Target).X, Pg(Bg(BID).Target).Y
End Select
End Function
Public Function PBg_add_AttributeAssignment( _
ByRef PBID As Long, ByRef PID As Long, ByRef Ar As Long, _
ByVal Atk As Single, ByRef Sb As Boolean, _
ByRef Sp As Long, ByRef X As Single, _
ByRef Y As Single, ByRef PenHp As Long, _
ByRef Pen As Boolean, ByRef Trl As Long, _
ByRef mX As Single, ByRef mY As Single, _
ByRef dE As Long)
'PBg_Add_ClassificationAndEntry input
PBg(PBID).a = True: PBg(PBID).Ar = Ar
PBg(PBID).Atk = Atk
PBg(PBID).Sb = Sb
PBg(PBID).Sp = Sp
PBg(PBID).X = X: PBg(PBID).Y = Y
PBg(PBID).PenHp = PenHp
PBg(PBID).Pen = Pen: PBg(PBID).Trl = Trl
PBg(PBID).mX = mX
PBg(PBID).mY = mY
Pg(PID).E = Pg(PID).E - dE
If Trl = 0 Then PBCD(PID) = False Else PBSkillCD(PID) = False
End Function
Public Function FPg_add_AttributeAssignment( _
ByRef FPID As Long, ByRef HP As Long, _
ByRef Ar As Long, ByRef Sb As Boolean, _
ByRef Sp As Long)
'FBg_Add input
FPg(FPID).a = True
FPg(FPID).HP = HP
FPg(FPID).MxHP = HP
FPg(FPID).Ar = Ar
FPg(FPID).Sb = Sb
FPg(FPID).Sp = Sp
End Function
Public Function Pg_Def_AttributeAssignment(ByRef PID As Long)
'Pg_Def input
Pg(PID).a = True
Pg(PID).HP = 100
Pg(PID).MxHP = 100
Pg(PID).MxEmp = 10
Pg(PID).EMP = 0
Pg(PID).Rank = 1
Pg(PID).Blt = 1
Pg(PID).Ar = 100
Pg(PID).E = 100
Pg(PID).MxE = 100
Pg(PID).Sb = True
Pg(PID).Sp = 30
Pg(PID).Esp = 1
Pg(PID).X = Pg_Def_AttributeAssignment_XDetermination(PID)
Pg(PID).Y = Pg_Def_AttributeAssignment_YDetermination(PID)
End Function
Public Function Pg_Def_AttributeAssignment_XDetermination(ByRef PID As Long) As Long
'Pg_Def_AttributeAssignment input
Pg_Def_AttributeAssignment_XDetermination = 6000 / (PgSum + 2) * (PID + 1)
End Function
Public Function Pg_Def_AttributeAssignment_YDetermination(ByRef PID As Long) As Long
'Pg_Def_AttributeAssignment input
Pg_Def_AttributeAssignment_YDetermination = 1000
End Function
Public Function Bg_add_AttributeAssignment( _
ByRef BID As Long, ByRef Ar As Long, _
ByRef Atk As Single, ByRef Sb As Boolean, _
ByRef Sp As Long, ByRef X As Single, _
ByRef Y As Single, ByRef mX As Single, ByRef mY As Single)
'Bg_Add input
Bg(BID).a = True
Bg(BID).Ar = Ar
Bg(BID).Atk = Atk
Bg(BID).Sb = Sb
Bg(BID).Sp = Sp
Bg(BID).X = X
Bg(BID).Y = Y
Bg(BID).mX = mX
Bg(BID).mY = mY
End Function

