Attribute VB_Name = "modCRE"
Option Explicit

Public Type tCRE
    X       As Double
    Y       As Double
    vX      As Double
    vY      As Double
    iX      As Double
    iY      As Double
    color   As Long

    ANG     As Double

End Type


Public C()  As tCRE
Public NC   As Long



Public Sub InitCRE(N As Long)
    Dim I   As Long


    ReDim C(N)
    NC = N

    For I = 1 To NC
        C(I).X = Rnd * MaxX
        C(I).Y = Rnd * MaxY
        C(I).color = RGB(60 + Rnd * 195, 60 + Rnd * 195, 60 + Rnd * 195)
    Next


End Sub

Public Sub DrawCreatures()
    Dim I   As Long
    Dim X   As Long
    Dim Y   As Long
    Dim X2  As Long
    Dim Y2  As Long

    For I = 1 To NC
        ' MyCircle pHDC, C(I).X * 1, C(I).Y * 1, 2, 4, C(I).color
        X = C(I).X
        Y = C(I).Y
        X2 = X + Cos(C(I).ANG) * 6
        Y2 = Y + Sin(C(I).ANG) * 6


        MyCircle pHDC, X, Y, 3, 1, C(I).color
        FastLine pHDC, X, Y, X2, Y2, 1, C(I).color

    Next
    MyCircle pHDC, C(PrevBest).X * 1, C(PrevBest).Y * 1, 7, 2, C(PrevBest).color

End Sub



Public Sub MoveCreatures()
    Dim I   As Long
    Dim D   As Double
    Dim V   As Double
    Dim ATG As Double
    Dim Ainc As Double

    For I = 1 To NC
        With C(I)

            '            .iX = GE.GetOUT(I, 1)
            '            .iY = GE.GetOUT(I, 2)
            '
            '            D = .iX * .iX + .iY * .iY
            '            If D > 1 Then
            '                D = 1 / Sqr(D)
            '                .iX = .iX * D
            '                .iY = .iY * D
            '            End If
            '
            '            .vX = .vX * 0.95 + .iX * 0.05
            '            .vY = .vY * 0.95 + .iY * 0.05

            V = GE.GetOUT(I, 1) * 0.01
            If V < 0.01 Then V = 0.01
            If V > 1 Then V = 1

            Ainc = GE.GetOUT(I, 2) * 0.0001
            If Ainc < -0.1 Then Ainc = -0.1
            If Ainc > 0.1 Then Ainc = 0.1

            .ANG = .ANG + Ainc
            '            While .ANG > PI2: .ANG = .ANG - PI2: Wend
            '            While .ANG < -PI2: .ANG = .ANG + PI2: Wend
            While .ANG > PI2: .ANG = .ANG - PI2: Wend
            While .ANG < 0: .ANG = .ANG + PI2: Wend

            .vX = Cos(.ANG) * V
            .vY = Sin(.ANG) * V

            .X = .X + .vX
            .Y = .Y + .vY

            If .X < 0 Then .X = 0: .vX = -.vX
            If .Y < 0 Then .Y = 0: .vY = -.vY
            If .X > MaxX Then .X = MaxX: .vX = -.vX
            If .Y > MaxY Then .Y = MaxY: .vY = -.vY


        End With

    Next


    TX = TX + (Rnd - 0.5) * 0.7
    TY = TY + (Rnd - 0.5) * 0.7

    If TX < RR Then TX = RR
    If TY < RR Then TY = RR
    If TX > MaxX - RR Then TX = MaxX - RR
    If TY > MaxY - RR Then TY = MaxY - RR



End Sub

Public Sub CreaturesSetInputs()
    Dim I   As Long
    Dim J   As Long


    Dim dx  As Double
    Dim dy  As Double
    Dim A   As Double


    '    For I = 1 To NC
    '        GE.SetINPUT(I, 1) = (TX - C(I).X)
    '        GE.SetINPUT(I, 2) = (TY - C(I).Y)
    '    Next

    For I = 1 To NC
        dx = (TX - C(I).X)
        dy = (TY - C(I).Y)
        A = Atan2(dx, dy)
        GE.SetINPUT(I, 1) = AngleDIFF(A, C(I).ANG)
        GE.SetINPUT(I, 2) = FASTsqr(dx * dx + dy * dy)
    Next





End Sub

Public Sub CreaturesUpdateFitness()
    Dim I   As Long

    Dim dx  As Double
    Dim dy  As Double
    Dim D   As Double


    For I = 1 To NC
        dx = C(I).X - TX
        dy = C(I).Y - TY
        D = (dx * dx + dy * dy)

        If D < RR2 Then
            D = RR2 - D    'To stay at circleRR from target
        Else
            D = D - RR2
        End If

        'D = D * 0.001
        'GE.Fitness(I) = GE.Fitness(I) + D
        D = FASTsqr(D * 0.01)
        GE.Fitness(I) = GE.Fitness(I) + D
    Next

End Sub

Public Sub CeaturesRandPOS()
    Dim I   As Long
    Dim X   As Double
    Dim Y   As Double

    X = TX    'Rnd * MaxX
    Y = TY    'Rnd * MaxY

    For I = 1 To NC
        C(I).X = X
        C(I).Y = Y
        C(I).vX = 0
        C(I).vY = 0
    Next

    TX = Rnd * MaxX
    TY = Rnd * MaxY

End Sub
