Attribute VB_Name = "modMAIN"
Option Explicit

Public GE   As clsGE

Public MaxX As Long
Public MaxY As Long
Public pHDC As Long

Public CNT  As Long

Public RR   As Double
Public RR2  As Double

Public PrevBest As Long

Public Const PIh As Double = 1.5707963267949
Public Const PI As Double = 3.14159265358979
Public Const PI2 As Double = 6.28318530717959


Private Const SavePopAtEveryGen As Long = 25


Public Sub INITW()


    Set GE = New clsGE
    '70
    GE.INIT 90, 2, 2, 7, 7, 18, 0.1, 0.3, 0.05

    InitCRE GE.PopSize

    RR = 50    '100 ' 90
    RR2 = RR * RR

End Sub


Public Function Atan2(X As Double, Y As Double) As Double
    If X Then
        Atan2 = -PI + Atn(Y / X) - (X > 0) * PI
    Else
        Atan2 = -PIh - (Y > 0) * PI
    End If
    ' While Atan2 < 0: Atan2 = Atan2 + PI2: Wend
    ' While Atan2 > PI2: Atan2 = Atan2 - PI2: Wend
End Function

Public Function AngleDIFF(ByRef A1 As Double, ByRef a2 As Double) As Double

    AngleDIFF = a2 - A1
    While AngleDIFF < -PI
        AngleDIFF = AngleDIFF + PI2
    Wend
    While AngleDIFF > PI
        AngleDIFF = AngleDIFF - PI2
    Wend
End Function

'Public Function Fastsqr(N As Double) As Double
'    Dim I      As Long
'    Dim X      As Double
'    If N Then
'        X = N * 0.25
'        For I = 1 To 12  '16 '12
'
'
'            X = (X + (N / X)) * 0.5
'        Next
'        Fastsqr = X
'    End If
'End Function




Public Sub MainLOOP()
    Dim I   As Long
    Dim J   As Long


    Dim InstructToRun As Long

    InstructToRun = 1    ' 4

    Do
        CNT = CNT + 1

        If CNT Mod 2 = 0 Then
            CreaturesUpdateFitness

            BitBlt pHDC, 0, 0, MaxX, MaxY, pHDC, 0, 0, 0
            DrawCreatures

            fMain.PIC.Refresh
            DoEvents

        End If

        '.................... BRAINS
        CreaturesSetInputs
        For J = 1 To InstructToRun
            For I = 1 To NC
                GE.RUNstep I
            Next
        Next J

        MoveCreatures
        '....................

        If CNT Mod 2500 = 0 Then

            If GE.Generation Mod SavePopAtEveryGen = 0 Then
                GE.SavePopulation "POP.txt"
            End If


            PrevBest = GE.GetBestIndi
            fMain.Label1 = "Best code (" & PrevBest & ") :"
            fMain.tCode = GE.GetExplicitCode

            GE.EVOLVE
            CeaturesRandRepos
            fMain.Caption = GE.Generation
        End If



    Loop While True

End Sub
