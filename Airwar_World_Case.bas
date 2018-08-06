Attribute VB_Name = "Airwar_World_Case"
Public Function Éý¼¶(ByRef PID As Long)
Randomize
Do While Pg(PID).EMP >= Pg(PID).MxEmp
    Pg(PID).EMP = Pg(PID).EMP - Pg(PID).MxEmp
    Pg(PID).MxEmp = Int(Rnd * (10 ^ Pg(PID).Rank - 10 * Pg(PID).Rank + 1)) + 10 * Pg(PID).Rank + Pg(PID).MxEmp
    Pg(PID).Rank = Pg(PID).Rank + 1
    Pg(PID).MxHP = Int(Rnd * (10 * Pg(PID).Rank - 9 + Pg(PID).Rank)) + 10 + Pg(PID).Rank + Pg(PID).MxHP
    Pg(PID).HP = Pg(PID).MxHP
    Pg(PID).E = Pg(PID).MxE
    Pg(PID).Esp = Pg(PID).Esp + 0.5
    Pg(PID).Sp = 3 + Pg(PID).Sp
    If Pg(PID).Rank >= 3 Then PSkill(PID, 1) = True
    If Pg(PID).Rank >= 5 Then PSkill(PID, 2) = True
Loop
End Function
Public Function PCEsp_Recovery()
Dim i As Long
For i = 0 To 1
    If Pg(i).a = True Then
        If Pg(i).E < Pg(i).MxE - Pg(i).Esp Then
            Pg(i).E = Pg(i).E + Pg(i).Esp
        Else
            If Pg(i).E > Pg(i).MxE - Pg(i).Esp And Pg(i).E < Pg(i).MxE Then
                Pg(i).E = Pg(i).MxE
            End If
        End If
    End If
Next
End Function
Public Function FoeBullet_Shoot(ByRef FPID As Long, Optional Arnk As Long = 0)
Dim c As Long
If FoeBullet_Shoot_Chance() = False Then Exit Function
For c = 0 To BgSum - 1
    If Bg(c).a = False Then
        Bg_Add c, Arnk, FPID
        Exit Function
    End If
Next
Bg_Add c, Arnk, FPID
BgSum = BgSum + 1
If BgSum + 100 > UBound(Bg) Then ReDim Preserve Bg(BgSum + 1100)
End Function
Public Function FoeBullet_Shoot_Chance() As Boolean
Randomize
If Int(Rnd * (31)) - 20 + Diff / 10 > 0 Then FoeBullet_Shoot_Chance = True
End Function
Public Function FoePlaneLifeSum() As Long
Dim i As Long
For i = 0 To FPSum
    If FPg(i).a = True Then
        FoePlaneLifeSum = FoePlaneLifeSum + 1
    End If
Next
End Function
Public Sub Test_Bullet_2()
'100ms
Dim i As Long
For i = 0 To FPSum - 1
    If FPg(i).a = True Then
        FoeBullet_Shoot i, FPg(i).AiRank
    End If
Next
End Sub
Public Function Key_Config_Def()
Key_Config_Def_Player1
Key_Config_Def_Player2
End Function
Public Function Key_Config_Def_Player1()
With PC(0)
    .Up = 87
    .Down = 83
    .Left = 65
    .Right = 68
    .Attack = 74
    .Ultimate_Skill = 75
    .Skill_Switch = 76
End With
End Function
Public Function Key_Config_Def_Player2()
With PC(1)
    .Up = 38
    .Down = 40
    .Left = 37
    .Right = 39
    .Attack = 97
    .Ultimate_Skill = 98
    .Skill_Switch = 99
End With
End Function
Public Function Test_FoePlane(Optional Arnk As Long = 0)
'[Form1]Timer5
For i = 0 To FPSum - 1
    If FPg(i).a = False Then
        FPg_Add i, Arnk: Exit Function
    End If
Next
FPg_Add FPSum, Arnk
FPSum = FPSum + 1
End Function
