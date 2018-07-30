Attribute VB_Name = "Airwar_ruler_engine"
Public Sub 物理事件检测()
Dim i As Long
物理事件检测_Plane主体
物理事件检测_PlaneBullet主体
For i = 0 To 1
    If Pg(i).a = True Then
        Form1.JinDuT1(i).Progress = Pg(i).HP / Pg(i).MxHP * 100
        Form1.JinDuT2(i).Progress = Pg(i).E / Pg(i).MxE * 100
        If PBSkillCD(i) = False Then Form1.JinDuT3(i).Progress = PBSkillCDTime(i) Else Form1.JinDuT3(i).Progress = 100
    End If
Next
End Sub
Public Sub 移动事件处理()
移动事件处理_Plane主体
移动事件处理_FoePlane主体
移动事件处理_Bullet主体
移动事件处理_Supply主体
移动事件处理_PlaneBullet主体
End Sub
Public Sub 物理事件检测_PlaneBullet主体()
Dim i As Long
For i = 0 To PBSum
    If PBg(i).a = True Then
        Select Case PBg(i).Trl
            Case 0
                物理事件检测_PlaneBullet主体_Trl_1 i
            Case 1
                物理事件检测_PlaneBullet主体_Trl_2 i
            Case 2
                物理事件检测_PlaneBullet主体_Trl_3 i
        End Select
    End If
Next
End Sub
Public Sub 物理事件检测_Plane主体()
Dim i As Long
For i = 0 To 1
    If Pg(i).a = True Then
        物理事件检测_Plane主体_PaB i
        物理事件检测_Plane主体_PaS i
        物理事件检测_Plane主体_PaFP i
    End If
Next
End Sub
Public Sub 移动事件处理_FoePlane主体()
Dim i As Long
For i = 0 To FPSum
    If FPg(i).a = True Then
        FoePlaneWYKZ i
    End If
Next
End Sub
Public Sub 移动事件处理_Plane主体()
Dim i As Long
If Pg(0).a = True Then
    For i = 0 To 5
        PlaneWYKZ 0, KCTemp(i)
    Next
End If
If Pg(1).a = True Then
    For i = 0 To 5
        PlaneWYKZ 1, KCTemp(i)
    Next
End If
End Sub
Public Sub 移动事件处理_Bullet主体()
Dim i As Long
For i = 0 To BgSum - 1
    If Bg(i).a = True Then
        BulletWYKZ i, Bg(i).Trl
    End If
Next
End Sub
Public Sub 移动事件处理_Supply主体()
Dim i As Long
For i = 0 To SgSum - 1
    If Sg(i).a = True Then
        SupplyWYKZ i
    End If
Next
End Sub
Public Sub 移动事件处理_PlaneBullet主体()
Dim i As Long
For i = 0 To PBSum - 1
    If PBg(i).a = True Then
        PlBtWYKZ i, PBg(i).Trl
    End If
Next
End Sub


