Attribute VB_Name = "Airwar_ruler_engine"
Public Sub �����¼����()
Dim i As Long
�����¼����_Plane����
�����¼����_PlaneBullet����
For i = 0 To 1
    If Pg(i).a = True Then
        Form1.JinDuT1(i).Progress = Pg(i).HP / Pg(i).MxHP * 100
        Form1.JinDuT2(i).Progress = Pg(i).E / Pg(i).MxE * 100
        If PBSkillCD(i) = False Then Form1.JinDuT3(i).Progress = PBSkillCDTime(i) Else Form1.JinDuT3(i).Progress = 100
    End If
Next
End Sub
Public Sub �ƶ��¼�����()
�ƶ��¼�����_Plane����
�ƶ��¼�����_FoePlane����
�ƶ��¼�����_Bullet����
�ƶ��¼�����_Supply����
�ƶ��¼�����_PlaneBullet����
End Sub
Public Sub �����¼����_PlaneBullet����()
Dim i As Long
For i = 0 To PBSum
    If PBg(i).a = True Then
        Select Case PBg(i).Trl
            Case 0
                �����¼����_PlaneBullet����_Trl_1 i
            Case 1
                �����¼����_PlaneBullet����_Trl_2 i
            Case 2
                �����¼����_PlaneBullet����_Trl_3 i
        End Select
    End If
Next
End Sub
Public Sub �����¼����_Plane����()
Dim i As Long
For i = 0 To 1
    If Pg(i).a = True Then
        �����¼����_Plane����_PaB i
        �����¼����_Plane����_PaS i
        �����¼����_Plane����_PaFP i
    End If
Next
End Sub
Public Sub �ƶ��¼�����_FoePlane����()
Dim i As Long
For i = 0 To FPSum
    If FPg(i).a = True Then
        FoePlaneWYKZ i
    End If
Next
End Sub
Public Sub �ƶ��¼�����_Plane����()
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
Public Sub �ƶ��¼�����_Bullet����()
Dim i As Long
For i = 0 To BgSum - 1
    If Bg(i).a = True Then
        BulletWYKZ i, Bg(i).Trl
    End If
Next
End Sub
Public Sub �ƶ��¼�����_Supply����()
Dim i As Long
For i = 0 To SgSum - 1
    If Sg(i).a = True Then
        SupplyWYKZ i
    End If
Next
End Sub
Public Sub �ƶ��¼�����_PlaneBullet����()
Dim i As Long
For i = 0 To PBSum - 1
    If PBg(i).a = True Then
        PlBtWYKZ i, PBg(i).Trl
    End If
Next
End Sub


