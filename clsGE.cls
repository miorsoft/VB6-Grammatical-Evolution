VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type tPROG

    INP()   As Double
    OUT()   As Double
    RAM()   As Double
    '----
    PRG()   As Double
    ROM()   As Double

    Fitness As Double

    Changed As Boolean


    IP      As Long    'Instrucion Pointer

    CanBeWorst As Boolean    'Can be container for New Child

End Type

Private INDI() As tPROG
Public Ninp As Long
Public Nout As Long
Public Nram As Long
Public Nrom As Long
Public Nprg As Long


Public PopSize As Long



Private Const Ninstr As Long = 9 + 2   'Set of instructions

Private Const InstrLEN As Long = 7

Private Const MaxResult As Double = 1000000
Private Const STARTFIT As Double = 100

Public Generation As Long

Public LinesOfCode As Long

'Private Const Nsons As Long = 4
Private Nsons As Long
Private MUTprob As Double
Private MUTrate As Double


Public Property Get Fitness(wI As Long) As Double
    Fitness = INDI(wI).Fitness
End Property

Public Property Let Fitness(wI As Long, V As Double)
    INDI(wI).Fitness = V
End Property

Public Property Get GetOUT(wI As Long, OutN As Long) As Double
    GetOUT = INDI(wI).OUT(OutN)
End Property

Public Property Let SetINPUT(wI As Long, InpN As Long, V As Double)
    INDI(wI).INP(InpN) = V
End Property
Private Function Clamp(V As Double, Min As Double, Max As Double) As Double
    Clamp = V

    If Clamp < Min Then
        Clamp = Min
    ElseIf Clamp > Max Then
        Clamp = Max
    End If


End Function



Public Sub INIT(PopulationSize As Long, Inputs As Long, Outputs As Long, Rams As Long, Roms As Long, NCodeLines As Long, _
                EVOSonsPerc As Double, EVOChildMutationProb As Double, EVOMutationRate As Double)
    Dim I   As Long

    PopSize = PopulationSize

    Ninp = Inputs
    Nout = Outputs
    Nram = Rams
    Nprg = NCodeLines * InstrLEN
    Nrom = Roms
    LinesOfCode = NCodeLines

    MUTprob = EVOChildMutationProb
    MUTrate = EVOMutationRate

    Nsons = PopSize * EVOSonsPerc
    If Nsons < 1 Then Nsons = 1

    ReDim INDI(PopSize)

    For I = 1 To PopSize
        INDIRandomizeInit I
    Next

    Generation = 0

End Sub

Private Sub INDIRandomizeInit(wI As Long)
    Dim I   As Long

    With INDI(wI)

        ReDim .INP(Ninp)
        ReDim .OUT(Nout)
        ReDim .RAM(Nram)
        ReDim .PRG(Nprg)
        ReDim .ROM(Nrom)

        For I = 1 To Nprg
            .PRG(I) = RandomG
        Next
        For I = 1 To Nrom
            .ROM(I) = RandomG(True)
        Next

        .IP = 1
        .Fitness = STARTFIT
        .Changed = True

    End With

End Sub
Private Sub ResetRAM(wI As Long)
    ReDim INDI(wI).RAM(Nram)
End Sub

Private Function RandomG(Optional NegToo As Boolean = False) As Double
    If NegToo Then
        RandomG = RndM * 2000 * 2 - 2000
    Else
        RandomG = RndM * 2000
    End If
End Function



Private Function SonOf(ByVal A As Long, ByVal b As Long) As tPROG
'CROSSOVER
    Dim Son As tPROG
    Dim I   As Long
    Dim J   As Long
    Dim T   As Long
    Dim X   As Long

    With Son
        ReDim .INP(Ninp)
        ReDim .OUT(Nout)
        ReDim .RAM(Nram)
        ReDim .PRG(Nprg)
        ReDim .ROM(Nrom)
    End With


    If RndM < 0.5 Then T = A: A = b: b = T    'Swap Parents

    If RndM < 0.5 Then
        '2 points rom
        I = Int(RndM * Nrom) + 1
        J = Int(RndM * Nrom) + 1
        If J < I Then
            T = J: J = I: I = T
        End If

        For X = 1 To I
            Son.ROM(X) = INDI(A).ROM(X)
        Next
        For X = I + 1 To J
            Son.ROM(X) = INDI(b).ROM(X)
        Next
        For X = J + 1 To Nrom
            Son.ROM(X) = INDI(A).ROM(X)
        Next
    Else
        'uniform ROM
        For X = 1 To Nrom
            If RndM < 0.5 Then
                Son.ROM(X) = INDI(A).ROM(X)
            Else
                Son.ROM(X) = INDI(b).ROM(X)
            End If
        Next

    End If


    If RndM < 0.5 Then
        '2 points prg
        I = Int(RndM * Nprg) + 1
        J = Int(RndM * Nprg) + 1
        If J < I Then
            T = J: J = I: I = T
        End If

        For X = 1 To I
            Son.PRG(X) = INDI(A).PRG(X)
        Next
        For X = I + 1 To J
            Son.PRG(X) = INDI(b).PRG(X)
        Next
        For X = J + 1 To Nprg
            Son.PRG(X) = INDI(A).PRG(X)
        Next
    Else
        'uniform PRG
        For X = 1 To Nprg
            If RndM < 0.5 Then
                Son.PRG(X) = INDI(A).PRG(X)
            Else
                Son.PRG(X) = INDI(b).PRG(X)
            End If
        Next
    End If

    Son.Changed = True
    Son.Fitness = STARTFIT
    Son.IP = 1
    Son.CanBeWorst = False


    SonOf = Son

End Function





Public Sub AddNewSonToPop(I As Long, J As Long)
    PopSize = PopSize + 1
    ReDim Preserve INDI(PopSize)

    INDI(PopSize) = SonOf(I, J)

End Sub






Public Sub RUNstep(wI As Long)
'(make changes to GetExplicitCode too)

    Dim INSTRU As Long
    Dim OP1type As Long
    Dim op1Addr As Long
    Dim OP2type As Long
    Dim op2Addr As Long
    Dim SaveTo As Long
    Dim SaveAddr As Long
    Dim JumpTO As Long


    Dim A   As Double
    Dim b   As Double
    Dim R   As Double

    Dim iPTR As Long

    With INDI(wI)
        If .IP = 0 Then .IP = 1
        iPTR = .IP


        INSTRU = .PRG(iPTR) Mod Ninstr + 1

        If INSTRU <> 1 Then    'no NOP

            OP1type = .PRG(iPTR + 1) Mod 3 + 1
            Select Case OP1type
                Case 1    'Inp
                    op1Addr = .PRG(iPTR + 2) Mod Ninp + 1
                    A = .INP(op1Addr)
                Case 2    'Ram
                    op1Addr = .PRG(iPTR + 2) Mod Nram + 1
                    A = .RAM(op1Addr)
                Case 3    'ROM
                    op1Addr = .PRG(iPTR + 2) Mod Nrom + 1
                    A = .ROM(op1Addr)
            End Select

            OP2type = .PRG(iPTR + 3) Mod 3 + 1
            Select Case OP2type
                Case 1    'Inp
                    op2Addr = .PRG(iPTR + 4) Mod Ninp + 1
                    b = .INP(op2Addr)
                Case 2    'Ram
                    op2Addr = .PRG(iPTR + 4) Mod Nram + 1
                    b = .RAM(op2Addr)
                Case 3     'Rom
                    op2Addr = .PRG(iPTR + 4) Mod Nrom + 1
                    b = .ROM(op2Addr)

            End Select



            '----------------------------------------------------------------------------------
            Select Case INSTRU

                Case 2    'ADD
                    R = A + b
                Case 3    'SUB
                    R = A - b
                Case 4    'mul
                    R = A * b
                Case 5    'div
                    If b Then R = A / b
                Case 6    'pow
                    While Abs(A) > 100: A = A * 0.1: Wend
                    While Abs(b) > 10: b = b * 0.1: Wend
                    R = Abs(A) ^ (Abs(b))
                Case 7    'greater
                    R = IIf(A > b, A, b)
                Case 8    'smaller
                    R = IIf(A < b, A, b)
                Case 9
                    R = A
                    '            Case 10
                    '                R = Cos(A)
                    '            Case 11
                    '                R = Sin(A)
                    '            Case 12
                    '                R = Tan(A)
                    '            Case 13
                    '                R = Log(Abs(A) + 0.00001)
                    '            Case 14
                    '                R = Exp(Clamp(A, -100, 100))
                    '            Case 15
                    '                R = Abs(A)
                    '------
                Case 10    '16    '10
                    'just JUMP
                Case 11    '17    '11    'JMPgreater
                    R = CLng(A > b)
                Case 12    '18    '12   'JMPsmaller
                    R = CLng(A < b)
            End Select


            If R < -MaxResult Then R = -MaxResult
            If R > MaxResult Then R = MaxResult

            '----------------------------------------------------------------------------------
            If INSTRU < 10 Then    '10 '16
                SaveTo = .PRG(iPTR + 5) Mod 2 + 1
                Select Case SaveTo
                    Case 1    'Save to RAM
                        SaveAddr = .PRG(iPTR + 6) Mod Nram + 1
                        .RAM(SaveAddr) = R
                    Case 2    'Save to OUT
                        SaveAddr = .PRG(iPTR + 6) Mod Nout + 1
                        .OUT(SaveAddr) = R
                End Select
            Else
                Select Case INSTRU
                    Case 10    '16 '10    'jump
                        JumpTO = .PRG(iPTR + 6) Mod LinesOfCode + 1
                        JumpTO = 1 + (JumpTO - 1) * InstrLEN - InstrLEN
                        .IP = JumpTO
                    Case 11, 12    '17, 18
                        If R Then    'JUMP IF > or <
                            JumpTO = .PRG(iPTR + 6) Mod LinesOfCode + 1
                            JumpTO = 1 + (JumpTO - 1) * InstrLEN - InstrLEN
                            .IP = JumpTO
                        End If
                End Select


            End If


        End If

        .IP = .IP + InstrLEN
        If .IP > Nprg Then .IP = 1

    End With

End Sub
Public Function GetExplicitCode(Optional ofIndi As Long = 0) As String
' (Look at RUNstep)
    Dim S   As String
    Dim Be  As Long
    Dim I   As Long

    'ofIndi = 0 means get Best Indi, otherwise individual is specified by ofIndi itself
    If ofIndi = 0 Then
        Be = GetBestIndi
    Else
        Be = ofIndi
    End If


    Dim INSTRU As Long
    Dim OP1type As Long
    Dim op1Addr As Long
    Dim OP2type As Long
    Dim op2Addr As Long
    Dim SaveTo As Long
    Dim SaveAddr As Long
    Dim JumpTO As Long


    Dim A   As String
    Dim b   As String
    Dim R   As String
    Dim vO  As String

    Dim iPTR As Long
    Dim LN  As Long

    With INDI(Be)

        For iPTR = 1 To Nprg Step InstrLEN

            LN = LN + 1

            INSTRU = .PRG(iPTR) Mod Ninstr + 1

            If INSTRU <> 1 Then    'no NOP

                OP1type = .PRG(iPTR + 1) Mod 3 + 1
                Select Case OP1type
                    Case 1    'Inp
                        op1Addr = .PRG(iPTR + 2) Mod Ninp + 1
                        A = "INP(" & op1Addr & ")"
                    Case 2    'Ram
                        op1Addr = .PRG(iPTR + 2) Mod Nram + 1
                        A = "RAM(" & op1Addr & ")"
                    Case 3    'ROM
                        op1Addr = .PRG(iPTR + 2) Mod Nrom + 1
                        A = "ROM(" & op1Addr & ")" & "[" & .ROM(op1Addr) & "]"
                End Select

                OP2type = .PRG(iPTR + 3) Mod 3 + 1
                Select Case OP2type
                    Case 1    'Inp
                        op2Addr = .PRG(iPTR + 4) Mod Ninp + 1
                        b = "INP(" & op2Addr & ")"
                    Case 2    'Ram
                        op2Addr = .PRG(iPTR + 4) Mod Nram + 1
                        b = "RAM(" & op2Addr & ")"
                    Case 3     'Rom
                        op2Addr = .PRG(iPTR + 4) Mod Nrom + 1
                        b = "ROM(" & op2Addr & ")" & "[" & .ROM(op2Addr) & "]"
                End Select



                '----------------------------------------------------------------------------------
                Select Case INSTRU

                    Case 2    'ADD
                        R = A & " + " & b
                    Case 3    'SUB
                        R = A & " - " & b
                    Case 4    'mul
                        R = A & " * " & b
                    Case 5    'div
                        R = A & " / " & b
                    Case 6    'pow
                        R = A & " ^ " & b
                    Case 7    'greater
                        R = "Greater (" & A & "," & b & ")"
                    Case 8    'greater
                        R = "Smaller (" & A & "," & b & ")"

                    Case 9    'Assign
                        R = A
                        '---------
                        '                Case 10
                        '                    R = "Cos(" & A & ")"
                        '                Case 11
                        '                    R = "Sin(" & A & ")"
                        '                Case 12
                        '                    R = "Tan(" & A & ")"
                        '                Case 13
                        '                    R = "Log(Abs(" & A & "))"
                        '                Case 14
                        '                    R = "Exp(" & A & ")"
                        '                Case 15
                        '                    R = "Abs(" & A & ")"


                    Case 10    '16    ' 10
                        R = "JUMP ->"
                    Case 11    '17    ' 11
                        R = "jmpIF (" & A & " > " & b & ") ->"
                    Case 12    '18    ' 12
                        R = "jmpIF (" & A & " < " & b & ") ->"

                    Case Else
                        Stop


                End Select

                'R = "(" & R & ")"

                '----------------------------------------------------------------------------------

                SaveTo = .PRG(iPTR + 5) Mod 2 + 1
                Select Case SaveTo
                    Case 1    'Save to RAM
                        SaveAddr = .PRG(iPTR + 6) Mod Nram + 1
                        vO = "RAM(" & SaveAddr & ")"
                    Case 2    'Save to OUT
                        SaveAddr = .PRG(iPTR + 6) Mod Nout + 1
                        vO = "OUT(" & SaveAddr & ")"
                End Select

                If INSTRU < 10 Then    '10 '16
                    S = S & "[" & CStr(LN) & "] " & vO & " = " & R & vbCrLf
                Else

                    Select Case INSTRU
                        Case 10    '16    '10
                            JumpTO = .PRG(iPTR + 6) Mod LinesOfCode + 1
                            S = S & "[" & CStr(LN) & "] " & R & " {" & JumpTO & "}" & vbCrLf
                        Case 11, 12    '17, 18    '11,12
                            JumpTO = .PRG(iPTR + 6) Mod LinesOfCode + 1
                            S = S & "[" & CStr(LN) & "] " & R & " {" & JumpTO & "}" & vbCrLf
                    End Select
                End If


            Else
                S = S & "[" & CStr(LN) & "] " & "NOP" & vbCrLf
            End If



        Next

    End With

    GetExplicitCode = S

End Function
Private Function getWorst() As Long
    Dim I   As Long
    Dim V   As Double
    For I = 1 To PopSize
        If INDI(I).CanBeWorst Then If INDI(I).Fitness > V Then V = INDI(I).Fitness: getWorst = I
    Next

End Function

Public Sub EVOLVE()
    Dim FS  As Double
    Dim I   As Long
    Dim P() As Double
    Dim CP() As Double
    Dim J   As Long

    Dim totF As Double
    Dim S   As Double
    Dim Max As Double

    Dim P1  As Long
    Dim P2  As Long
    Dim R   As Double
    Dim W   As Long
    Dim Son As tPROG

    ReDim P(PopSize)
    ReDim CP(PopSize)

    For I = 1 To PopSize
        FS = FS + INDI(I).Fitness
        If INDI(I).Fitness > Max Then Max = INDI(I).Fitness
        INDI(I).CanBeWorst = True
    Next
    totF = FS


    For I = 1 To PopSize
        P(I) = (Max - INDI(I).Fitness) / totF
    Next


    FS = 0
    For I = 1 To PopSize
        FS = FS + P(I)
    Next


    If FS Then
        For I = 1 To PopSize
            P(I) = P(I) / FS
            S = S + P(I)
            'Debug.Print P(I), s
            CP(I) = S
        Next
    Else
        For I = 1 To PopSize
            CP(I) = CP(I - 1) + 1 / totF
        Next
    End If



    For J = 1 To Nsons
        '-----------------------------
        Do

            I = 0
            R = RndM
            Do
                I = I + 1
            Loop While CP(I) < R
            P1 = I

            I = 0
            R = RndM
            Do
                I = I + 1
            Loop While CP(I) < R
            P2 = I

        Loop While P1 = P2

        W = getWorst
        Debug.Print P1, P2, W
        Son = SonOf(P1, P2)

        INDI(W) = Son
        If RndM < MUTprob Then IndiMutate (W)


        '------------------------
    Next



    For I = 1 To PopSize - 1
        For J = I + 1 To PopSize
            R = CalcDiversity(I, J)
            If R = 0 Then
                INDIRandomizeInit J
            End If
        Next
    Next


    For I = 1 To PopSize
        INDI(I).Fitness = STARTFIT
        ResetRAM I
    Next

    Generation = Generation + 1

End Sub


Private Sub IndiMutate(wI As Long)
    Dim I   As Long
    Dim J   As Long
    Dim M   As Long
    Dim MT  As Long


    With INDI(wI)

        'Each gene 5%
        '        For I = 1 To Nrom
        '            If rndm < MUTrate Then .ROM(I) = RandomG(True)
        '        Next
        '        For I = 1 To Nprg
        '            If rndm < MUTrate Then .PRG(I) = RandomG
        '        Next

        '5% of all genes
        '        MT = Nrom *MUTrate
        '        If MT < 1 Then MT = 1
        '        M = 0
        '        Do
        '            I = Int(rndm * Nrom) + 1
        '            .ROM(I) = RandomG(True)
        '            M = M + 1
        '        Loop While M < MT
        '
        '
        '        MT = Nprg * 0.02
        '        If MT < 1 Then MT = 1
        '        M = 0
        '        Do
        '            I = Int(rndm * Nprg) + 1
        '            .PRG(I) = RandomG
        '            M = M + 1
        '        Loop While M < MT

        '5% of all genes ... NOT ROM/PRG
        MT = (Nrom + Nprg) * MUTrate
        If MT < 1 Then MT = 1
        M = 0
        Do


            I = Int(RndM * (Nrom + Nprg)) + 1
            If I <= Nrom Then
                .ROM(I) = RandomG(True)
            Else
                .PRG(I - Nrom) = RandomG
            End If

            M = M + 1
        Loop While M < MT

    End With

End Sub

Public Function GetBestIndi() As Long
    Dim F   As Double
    Dim I   As Long

    F = 1E+99

    For I = 1 To PopSize
        If INDI(I).Fitness < F Then
            F = INDI(I).Fitness
            GetBestIndi = I
        End If
    Next
End Function





Private Function CalcDiversity(A As Long, b As Long) As Double
    Dim D   As Double
    Dim I   As Long


    For I = 1 To Nrom
        D = D + Abs(INDI(A).ROM(I) - INDI(b).ROM(I))
    Next

    For I = 1 To Nprg
        D = D + Abs(INDI(A).PRG(I) - INDI(b).PRG(I))
    Next
    CalcDiversity = D

End Function

Private Function CtP(S As Variant) As String
'Comma to point
    CtP = Replace(CStr(S), ",", ".")
End Function

Public Sub SavePopulation(FN As String)


    Open App.Path & "\" & FN For Output As 1
    Dim I   As Long
    Dim J   As Long

    Print #1, "Population size"
    Print #1, PopSize

    Print #1, "N of inputs"
    Print #1, Ninp

    Print #1, "N of outputs"
    Print #1, Nout

    Print #1, "N of RAMs"
    Print #1, Nram

    Print #1, "N of ROMs"
    Print #1, Nrom

    Print #1, "CodeLines"
    Print #1, LinesOfCode

    Print #1, "EVO- Sons Percentage (of Population size)"
    Print #1, CtP(Nsons / PopSize)

    Print #1, "EVO- Mutation Probability (For Each Son)"
    Print #1, CtP(MUTprob)
    Print #1, "EVO- Mutation Rate"
    Print #1, CtP(MUTrate)


    For I = 1 To NC
        Print #1, "------------------------INDI " & I
        Print #1, "ROM:"
        For J = 1 To Nrom
            Print #1, CtP(INDI(I).ROM(J))
        Next
        Print #1, "CODE:"
        For J = 1 To Nprg
            Print #1, CtP(INDI(I).PRG(J))
        Next
    Next


    'Optional Save Explicit Code
    Print #1, " "
    Print #1, "************************************* Explicit CODES : "
    For I = 1 To NC
        Print #1, " "
        Print #1, "------------------------INDI " & I
        Print #1, " "
        Print #1, GetExplicitCode(I)
    Next

    Close 1

End Sub
Public Sub LoadPopulation(FN As String)

    If Len(Dir(App.Path & "\" & FN)) = 0 Then MsgBox "Previous Population not Found! " & vbCrLf & "Starting with a new random one ...", vbInformation: Exit Sub


    Open App.Path & "\" & FN For Input As 1
    Dim I   As Long
    Dim J   As Long
    Dim S   As String
    Dim V   As Double

    Dim pp As Long, ni As Long, no As Long
    Dim nra As Long, nro As Long, nlines As Long
    Dim Es As Double, Emp As Double, Emr As Double



    Input #1, S    '"Popsize"
    Input #1, pp

    Input #1, S    '"Ninp"
    Input #1, ni

    Input #1, S    '"Nout"
    Input #1, no

    Input #1, S    '"Nram"
    Input #1, nra

    Input #1, S    '"Nrom"
    Input #1, nro

    Input #1, S    '"CodeLines"
    Input #1, nlines

    Input #1, S    '"EVOSonsP"
    Input #1, Es

    Input #1, S    '"EVOMutProb"
    Input #1, Emp
    Input #1, S    '"EVOMutRate"
    Input #1, Emr


    'Preform Initialization
    INIT pp, ni, no, nra, nro, nlines, Es, Emp, Emr


    'Load Individuals ....

    For I = 1 To NC
        Input #1, S    ' "------------------------INDI " & I
        Input #1, S    '"ROM:"
        For J = 1 To Nrom
            Input #1, INDI(I).ROM(J)
        Next
        Input #1, S    '"Code:"
        For J = 1 To Nprg
            Input #1, INDI(I).PRG(J)
        Next
    Next

    Close 1

End Sub


