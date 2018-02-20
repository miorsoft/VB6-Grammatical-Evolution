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

Private Nearest() As Long
Private MinD() As Double


Public Sub InitCRE(N As Long)
    Dim I   As Long


    ReDim C(N)
    ReDim Nearest(N)
    ReDim MinD(N)

    NC = N

    For I = 1 To NC
        C(I).X = RndM * MaxX
        C(I).Y = RndM * MaxY
        C(I).color = RGB(80 + RndM * 175, 80 + RndM * 175, 80 + RndM * 175)
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
        X2 = X + Cos(C(I).ANG) * 7
        Y2 = Y + Sin(C(I).ANG) * 7


        MyCircle pHDC, X, Y, 3, 1, C(I).color
        FastLine pHDC, X, Y, X2, Y2, 1, C(I).color

    Next

    '    MyCircle pHDC, C(PrevBest).X * 1, C(PrevBest).Y * 1, 7, 2, C(PrevBest).color
    MyCircle pHDC, C(PrevBest).X * 1, C(PrevBest).Y * 1, RR * 1, 1, C(PrevBest).color


    FastLine pHDC, C(PrevBest).X * 1, C(PrevBest).Y * 1, _
             C(Nearest(PrevBest)).X * 1, C(Nearest(PrevBest)).Y * 1, 1, C(PrevBest).color

End Sub
Private Function Clamp(V As Double, Min As Double, Max As Double) As Double
    Clamp = V

    If Clamp < Min Then
        Clamp = Min
    ElseIf Clamp > Max Then
        Clamp = Max
    End If


End Function

Public Sub MoveCreatures()
    Dim I   As Long
    Dim D   As Double
    Dim V   As Double
    Dim ATG As Double
    Dim Ainc As Double



    For I = 1 To NC
        With C(I)

            V = GE.GetOUT(I, 1) * 0.01
            If V < 0.01 Then V = 0.01
            If V > 0.75 Then V = 0.75

            Ainc = GE.GetOUT(I, 2) * 0.0001
            'If Ainc < -0.1 Then Ainc = -0.1
            'If Ainc > 0.1 Then Ainc = 0.1

            Ainc = Clamp(Ainc, -0.08, 0.08)

            .ANG = .ANG + Ainc
            While .ANG > PI2: .ANG = .ANG - PI2: Wend
            While .ANG < 0: .ANG = .ANG + PI2: Wend

            .vX = Cos(.ANG) * V
            .vY = Sin(.ANG) * V

            .X = .X + .vX
            .Y = .Y + .vY

            '            If .X < 0 Then .X = 0: .vX = -.vX
            '            If .Y < 0 Then .Y = 0: .vY = -.vY
            '            If .X > MaxX Then .X = MaxX: .vX = -.vX
            '            If .Y > MaxY Then .Y = MaxY: .vY = -.vY
            If .X < 0 Then .X = MaxX + .X
            If .Y < 0 Then .Y = MaxY + .Y
            If .X > MaxX Then .X = .X - MaxX
            If .Y > MaxY Then .Y = .Y - MaxY


        End With

    Next





End Sub

Public Sub CreaturesSetInputs()
    Dim I   As Long
    Dim J   As Long


    Dim dx  As Double
    Dim dy  As Double
    Dim a   As Double
    Dim D   As Double


    For I = 1 To NC
        MinD(I) = 1E+99
    Next

    For I = 1 To NC - 1
        For J = I + 1 To NC
            dx = C(J).X - C(I).X
            dy = C(J).Y - C(I).Y
            D = (dx * dx + dy * dy)
            If D < MinD(I) Then
                MinD(I) = D
                Nearest(I) = J
            End If
            If D < MinD(J) Then
                MinD(J) = D
                Nearest(J) = I
            End If
        Next
    Next


    For I = 1 To NC

        dx = (C(Nearest(I)).X - C(I).X)
        dy = (C(Nearest(I)).Y - C(I).Y)
        a = Atan2(dx, dy)
        GE.SetINPUT(I, 1) = AngleDIFF(a, C(I).ANG)

        D = dx * dx + dy * dy
        '        If D < RR2 Then
        '            D = RR - sqr(D)
        '        Else
        '            D = 0
        '        End If
        D = Sqr(D)

        GE.SetINPUT(I, 2) = D * 0.1
    Next





End Sub

Public Sub CreaturesUpdateFitness()
    Dim I   As Long
    Dim J   As Long


    Dim dx  As Double
    Dim dy  As Double
    Dim D   As Double


    '    For I = 1 To NC - 1
    '        For J = I + 1 To NC
    '
    '            dx = C(J).X - C(I).X
    '            dy = C(J).Y - C(I).Y
    '            D = (dx * dx + dy * dy)
    '
    '            'If D < RR2 Then
    '                D = RR2 - D
    '                'D = sqr(D) * 0.01
    '                D = sqr(Abs(D)) * 0.01
    '                GE.Fitness(I) = GE.Fitness(I) + D
    '                GE.Fitness(J) = GE.Fitness(J) + D
    '            'Else
    '            'End If
    '        Next
    '    Next

    For I = 1 To NC
        J = Nearest(I)
        dx = C(J).X - C(I).X
        dy = C(J).Y - C(I).Y
        D = (dx * dx + dy * dy)
        D = RR2 - D
        D = Sqr(Abs(D)) * 0.01
        GE.Fitness(I) = GE.Fitness(I) + D
    Next



End Sub

Public Sub CeaturesRandRepos()
    Dim I   As Long
    Dim X   As Double
    Dim Y   As Double
    Dim J   As Long

    Dim a   As Double
    a = RndM * PI

    Dim NotFree() As Boolean
    ReDim NotFree(NC)

    Do
        J = Int(RndM * NC) + 1
        If Not (NotFree(J)) Then
            I = I + 1
            C(I).X = MaxX * 0.5 + Cos(J / NC * PI2) * MaxY * 0.4
            C(I).Y = MaxY * 0.5 + Sin(J / NC * PI2) * MaxY * 0.4
            C(I).vX = 0
            C(I).vY = 0
            C(I).ANG = RndM * PI2
            NotFree(J) = True

        End If
    Loop While I < NC



End Sub
