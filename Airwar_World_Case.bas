Attribute VB_Name = "Airwar_World_Case"
Public Function Éý¼¶(ByRef PID As Long)
Randomize
If Pg(PID).Emp >= Pg(PID).MxEmp Then
    Pg(PID).MxEmp = Int(Rnd * (10 ^ Pg(PID).Rank - 10 * Pg(PID).Rank + 1)) + 10 * Pg(PID).Rank + Pg(PID).Emp
    Pg(PID).Emp = 0: Pg(PID).Rank = Pg(PID).Rank + 1
    Pg(PID).MxHP = Int(Rnd * (10 * Pg(PID).Rank - 9 + Pg(PID).Rank)) + 10 + Pg(PID).Rank + Pg(PID).MxHP
    Pg(PID).HP = Pg(PID).MxHP
    Pg(PID).E = Pg(PID).MxE
    Pg(PID).Esp = Pg(PID).Esp + 0.5
    Pg(PID).Sp = 3 + Pg(PID).Sp
    If Pg(PID).Rank >= 3 Then PSkill(PID, 1) = True
End If
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
If FoeBullet_Shoot_Chance() = False Then Exit Function
For c = 0 To BgSum - 1
    If Bg(c).a = False Then
        Bg_Add c, Arnk + 1, FPID
        Exit Function
    End If
Next
Bg_Add c, Arnk + 1, FPID
BgSum = BgSum + 1
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
Public Function Test_FoePlane(Optional Arnk As Long = 0)
'100ms
'If FPSum > 4 Then Exit Sub
For i = 0 To FPSum - 1
    If FPg(i).a = False Then
        FPg_Add i, Arnk: Exit Function
    End If
Next
FPg_Add FPSum, Arnk
FPSum = FPSum + 1
End Function
