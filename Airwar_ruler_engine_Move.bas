Attribute VB_Name = "Airwar_ruler_engine_Move"
Public Function FoePlaneWYKZ(ByRef FPID As Long)
Dim XXJie As 二元解
Select Case FPg(FPID).AiRank
    Case 0, 1, 2
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
    Case 0, 1, 2
        If Bg(BID).Da = False Then
            XXJie = 线性求解(Bg(BID).X, Bg(BID).Y, Bg(BID).mX, Bg(BID).mY, Bg(BID).Sp)
            Bg(BID).dX = XXJie.a: Bg(BID).dY = XXJie.b: Bg(BID).Da = True
        End If
        Bg(BID).X = Bg(BID).X + Bg(BID).dX: Bg(BID).Y = Bg(BID).Y + Bg(BID).dY
        If Bg(BID).X < 0 Or Bg(BID).Y < 0 Or Bg(BID).X > 6000 Or Bg(BID).Y > 8000 Then
            Bg(BID).a = False: Bg(BID).Da = False
        End If
    Case 3
        XXJie = 线性求解(Bg(BID).X, Bg(BID).Y, Pg(0).X, Pg(0).Y, Bg(BID).Sp)
        Bg(BID).dX = XXJie.a: Bg(BID).dY = XXJie.b
        Bg(BID).X = Bg(BID).X + Bg(BID).dX: Bg(BID).Y = Bg(BID).Y + Bg(BID).dY
        If Bg(BID).X < 0 Or Bg(BID).Y < 0 Or Bg(BID).X > 6000 Or Bg(BID).Y > 8000 Then
            Bg(BID).a = False: Bg(BID).Da = False
        End If
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
        If PBg(PBID).Da = False Then
            XXJie = 线性求解(PBg(PBID).X, PBg(PBID).Y, PBg(PBID).mX, PBg(PBID).mY, PBg(PBID).Sp)
            PBg(PBID).dX = XXJie.a: PBg(PBID).dY = XXJie.b: PBg(PBID).Da = True
        End If
        PBg(PBID).X = PBg(PBID).X + PBg(PBID).dX: PBg(PBID).Y = PBg(PBID).Y + PBg(PBID).dY
        If PBg(PBID).X = PBg(PBID).mX And PBg(PBID).Y = PBg(PBID).mY Then PBg(PBID).Ar = 3000
        If PBg(PBID).X < 0 Or PBg(PBID).Y < 0 Or PBg(PBID).X > 6000 Or PBg(PBID).Y > 11000 Then
            PBg(PBID).a = False: PBg(PBID).Da = False
        End If
    Case 2
        
        If Form1.Ftime.Enabled = False Then
            PSkillID_Ft = 0: Form1.Ftime.Enabled = True: PBg(PBID).X = Pg(0).X: PBg(PBID).Y = Pg(0).Y
        Else
            If PSkillID_Ft > 3 + 0.5 * Pg(0).Rank Then
                Form1.Ftime.Enabled = False
                PBg(PBID).a = False
            Else
                PBg(PBID).X = Pg(0).X: PBg(PBID).Y = Pg(0).Y
            End If
        End If
End Select
End Function
Public Function PlaneWYKZ(ByRef PlID As Long, KeyCode As Integer)
Select Case KeyCode
    Case 37, 65
        If Pg(PlID).X - Pg(PlID).Sp - Pg(PlID).Ar <= 0 Then Exit Function
        Pg(PlID).X = Pg(PlID).X - Pg(PlID).Sp
    Case 38, 87
        If Pg(PlID).Y + Pg(PlID).Sp + Pg(PlID).Ar >= Form1.Picture1.Height Then Exit Function
        Pg(PlID).Y = Pg(PlID).Y + Pg(PlID).Sp
    Case 39, 68
        If Pg(PlID).X + Pg(PlID).Sp + Pg(PlID).Ar >= Form1.Picture1.Width Then Exit Function
        Pg(PlID).X = Pg(PlID).X + Pg(PlID).Sp
    Case 40, 83
        If Pg(PlID).Y - Pg(PlID).Sp - Pg(PlID).Ar <= 0 Then Exit Function
        Pg(PlID).Y = Pg(PlID).Y - Pg(PlID).Sp
    Case 32, 81
        PBg_Add PlID
    Case 69
        PBg_Add PlID, Pg(0).Blt
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
