Attribute VB_Name = "Airwar_ruler_engine_Phy"
Public Function �����¼����_Plane����_PaB(ByRef i As Long)
Dim Min(2), MiTemp As Long
Min(0) = 99999
For c = 0 To BgSum - 1
    If Bg(c).a = True Then
        MiTemp = �������(Pg(i).X, Pg(i).Y, Bg(c).X, Bg(c).Y)
        If Min(0) > MiTemp Then
            Min(0) = MiTemp: Min(1) = c
        End If
    End If
Next
If Min(0) <= Pg(i).Ar + Bg(Min(1)).Ar Then �����¼����_Plane����_PaB�¼� i, Min(1)
End Function
Public Function �����¼����_Plane����_PaFP(ByRef i As Long)
Dim Min(2), MiTemp As Long
Min(0) = 99999
For c = 0 To FPSum - 1
    If FPg(c).a = True Then
        MiTemp = �������(Pg(i).X, Pg(i).Y, FPg(c).X, FPg(c).Y)
        If Min(0) > MiTemp Then
            Min(0) = MiTemp: Min(1) = c
        End If
    End If
Next
If Min(0) <= Pg(i).Ar + FPg(Min(1)).Ar Then �����¼����_Plane����_PaFP�¼� i, Min(1)
End Function
Public Function �����¼����_Plane����_PaFP�¼�(ByVal PlID As Long, ByVal FPID As Long)
If Pg(PlID).Sb = False Then Exit Function '�ɻ��Ƿ������˺�
Pg(PlID).HP = Pg(PlID).HP - FPg(FPID).HP
If Pg(PlID).HP <= 0 Then Pg(PlID).a = False
FPg(FPID).a = False: FPg(FPID).Da = False
End Function
Public Function �����¼����_Plane����_PaS(ByRef i As Long)
Dim Min(2), MiTemp As Long
Min(0) = 99999
For c = 0 To SgSum - 1
    If Sg(i).a = True Then
        MiTemp = �������(Pg(i).X, Pg(i).Y, Sg(c).X, Sg(c).Y)
        If Min(0) > MiTemp Then
            Min(0) = MiTemp: Min(1) = c
        End If
    End If
Next
If Min(0) <= Pg(i).Ar + 50 Then �����¼����_Plane����_PaS�¼� i, Min(1)
End Function
Public Function �����¼����_PlaneBullet����_Trl_1_Bullet(ByRef i As Long)
Dim Min(2), MiTemp As Long
Min(0) = 99999
For c = 0 To BgSum
    If Bg(c).a = True Then
        MiTemp = �������(PBg(i).X, PBg(i).Y, Bg(c).X, Bg(c).Y)
        If Min(0) > MiTemp Then
            Min(0) = MiTemp: Min(1) = c
        End If
    End If
Next
If Min(0) <= PBg(i).Ar + Bg(Min(1)).Ar Then �����¼����_Plane����_PBaB�¼� i, Min(1)
End Function
Public Function �����¼����_PlaneBullet����_Trl_1_FoePlane(ByRef i As Long)
Dim Min(2), MiTemp As Long
Min(0) = 99999
For c = 0 To FPSum
    If FPg(c).a = True Then
        MiTemp = �������(PBg(i).X, PBg(i).Y, FPg(c).X, FPg(c).Y)
        If Min(0) > MiTemp Then
            Min(0) = MiTemp: Min(1) = c
        End If
    End If
Next
If Min(0) <= PBg(i).Ar + FPg(Min(1)).Ar Then �����¼����_Plane����_PBaFP�¼� i, Min(1)
End Function
Public Function �����¼����_Plane����_PBaFP�¼�(ByVal PBID As Long, ByVal FPID As Long)
If FPg(FPID).Sb = False Then Exit Function
FPg(FPID).HP = FPg(FPID).HP - PBg(PBID).Atk
If FPg(FPID).HP <= 0 Then FPg(FPID).a = False: �����¼����_Plane����_PBaFP�¼�_������� PBID, FPID
If PBg(PBID).Pen = False Then PBg(PBID).a = False: PBg(PBID).Da = False
End Function
Public Function �����¼����_Plane����_PBaFP�¼�_�������(ByVal PBID As Long, ByVal FPID As Long)
Select Case FPg(FPID).AiRank
    Case 0
        Pg(PBID).Emp = Pg(PBID).Emp + 1
    Case 1
        Pg(PBID).Emp = Pg(PBID).Emp + 25
    Case 2
        Pg(PBID).Emp = Pg(PBID).Emp + 50
End Select
���� PBID
End Function
Public Function �����¼����_PlaneBullet����_Trl_1(ByRef i As Long)
�����¼����_PlaneBullet����_Trl_1_Bullet i
�����¼����_PlaneBullet����_Trl_1_FoePlane i
End Function
Public Function �����¼����_PlaneBullet����_Trl_2_Bullet(ByRef i As Long)
'E����
For c = 0 To BgSum
    If Bg(c).a = True Then
        If �������(PBg(i).X, PBg(i).Y, Bg(c).X, Bg(c).Y) <= PBg(i).Ar Then
            Bg(c).a = False: Bg(c).Da = False
        End If
    End If
Next
PBg(i).a = False: PBg(i).Da = False
End Function
Public Function �����¼����_PlaneBullet����_Trl_2_FoePlane(ByRef i As Long)
For c = 0 To FPSum
    If FPg(c).a = True Then
        If �������(PBg(i).X, PBg(i).Y, FPg(c).X, FPg(c).Y) <= PBg(i).Ar Then
            FPg(c).HP = FPg(c).HP - PBg(i).Atk
            If FPg(c).HP <= 0 Then
                FPg(c).a = False: FPg(c).Da = False: �����¼����_Plane����_PBaFP�¼�_������� i, c
            End If
        End If
    End If
Next
PBg(i).a = False: PBg(i).Da = False
End Function
Public Function �����¼����_PlaneBullet����_Trl_2(ByRef i As Long)
If PBg(i).X = PBg(i).mX And PBg(i).Y = PBg(i).mY Then
    �����¼����_PlaneBullet����_Trl_2_Bullet i
    �����¼����_PlaneBullet����_Trl_2_FoePlane i
End If
End Function
Public Function �����¼����_PlaneBullet����_Trl_3(ByRef i As Long)
For c = 0 To BgSum
    If Bg(c).a = True Then
        If �������(PBg(i).X, PBg(i).Y, Bg(c).X, Bg(c).Y) <= PBg(i).Ar Then
            Bg(c).a = False: Bg(c).Da = False
        End If
    End If
Next
For c = 0 To FPSum
    If FPg(c).a = True Then
        If �������(PBg(i).X, PBg(i).Y, FPg(c).X, FPg(c).Y) <= PBg(i).Ar Then
            FPg(c).HP = FPg(c).HP - PBg(i).Atk
            If FPg(c).HP <= 0 Then
                FPg(c).a = False: FPg(c).Da = False: �����¼����_Plane����_PBaFP�¼�_������� i, c
            End If
        End If
    End If
Next
End Function
Public Function �����¼����_Plane����_PBaB�¼�(ByVal PBID As Long, ByVal BID As Long)
If Bg(BID).Sb = False Then Exit Function
Bg(BID).a = False: Bg(BID).Da = False
PBg(PBID).a = False: PBg(PBID).Da = False
End Function
Public Function �����¼����_Plane����_PaS�¼�(ByVal PlID As Long, ByVal SID As Long)
Select Case Sg(SID).Tp
    Case 0
        Pg(PlID).HP = Pg(PlID).MxHP
    Case 1
        Pg(PlID).Sp = Pg(PlID).Sp + 3
    Case 2
        Pg(PlID).E = Pg(PlID).MxE
End Select
Sg(SID).a = False: Sg(SID).Da = False
End Function
Public Function �����¼����_Plane����_PaB�¼�(ByVal PlID As Long, ByVal BID As Long)
If Pg(PlID).Sb = False Then Exit Function '�ɻ��Ƿ������˺�
Pg(PlID).HP = Pg(PlID).HP - Bg(BID).Atk
If Pg(PlID).HP <= 0 Then Pg(PlID).a = False
If Bg(BID).Pen = False Then Bg(BID).a = False: Bg(BID).Da = False
End Function