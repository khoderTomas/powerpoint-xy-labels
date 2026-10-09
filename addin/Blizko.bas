Option Explicit

' Placement "Kratsi cary" (shorter leader lines; an option in the dialog, approved on 9 Oct 2026). Port of
' popisky/blizko.py with expressions in the same order so that rounding matches: the tests run the reference
' on the model this module works with (exported by Popisky.ZpracujSoubor bit for bit) and compare the result.
' 1. Start: the approved columns (module Rozlozeni).
' 2. Candidates of every label: next to the point, right and left of the point on the rows of the chart grid,
'    above and below the point, elsewhere in its own column, closer to the point in its own row, and the start
'    position. Cost = leader line length + direction preference + crossing a vertical line. Static conditions:
'    inside the plot area, clear of boxes and points, not across vertical lines or the dividing line, in the
'    same half of the chart as its point, leader line not across a box.
' 3. Threshold accepting: moving a label or swapping the rows of two labels is accepted when the cost grows
'    by less than the threshold. The conflict penalty (overlap, crossing leader lines, leader line across
'    another label) grows and the threshold falls, both linearly. The best state without conflicts is kept.
'    Random numbers are MINSTD and there is no Exp or ^.
'    VBA evaluates an expression in extended precision and rounds only the result, Python rounds after every
'    operation. So that both decide the same, the inputs are rounded to 1/1024 pt (differences, sums and
'    orientation tests are then exact), costs after the square root too, and the threshold and the penalty
'    are integers in 1/1024.
' 4. Remaining conflicts: repair (move the label, possibly also one blocking label, otherwise back to its
'    column). Finally polish: a label moves to a cheaper position only when that creates no conflict.
' Fallback placement (Nouzove, approved on 9 Oct 2026) when the columns do not fit, for both placement options:
' the start is the cheapest column split even outside the plot area. A label with an invalid start gets the
' position "not placed" instead (no rectangle, no leader line, cost NEUMISTEN: above any leader line, below the
' conflict penalty at the end of the search), so the search places as many labels as possible without
' conflicts. A leader line across a vertical line costs PRES_CARU_NOUZE. Then the labels not placed get their
' original position (two original positions never conflict, those conflicts are the user's) and placed labels
' in conflict with them move away. Every column gap is tried, the fewest labels not placed wins; with fewer
' than half of the labels placed the chart is left unchanged.
' Labels and candidate lists are indexed from 1 (Python from 0).

Private Const POSUN_MAX As Double = 120#   ' max vertical offset of a label beside its point
Private Const PRES_CARU As Double = 40#    ' leader line across a vertical line (label in another quadrant)
Private Const VLEVO As Double = 3#         ' preference: right of the point, then left, then above and below
Private Const NAD_POD As Double = 6#
Private Const SIKMO As Double = 0.25       ' per pt of vertical offset of a label beside its point
Private Const OBAL As Double = 1.5         ' padding of the bounding box used for quick rejection of pairs
Private Const KROKU_NA_POPISEK As Long = 600
Private Const PRAH0 As Double = 250#
Private Const PRAH1 As Double = 0#
Private Const POKUTA0 As Double = 100#
Private Const POKUTA1 As Double = 1000#
Private Const PROHOZENI As Double = 0.3    ' share of moves that swap the rows of two labels
Private Const OPRAVA_KANDIDATU As Long = 80
Private Const DOLADENI_KOL As Long = 50
Private Const POCET_DX As Long = 14
Private Const POCET_DV As Long = 8
Private Const MRIZ As Double = 1024#       ' grid of inputs and costs (1/MRIZ pt); integers in it are exact
Private Const REZERVA_MEZERY As Double = 0.01  ' conflict = closer than the column gap minus this (start on the grid)
Private Const NEUMISTEN As Double = 800#   ' fallback: cost of the position "not placed" (the original position)
Private Const PRES_CARU_NOUZE As Double = 400#
Private Const NIKDE As Double = 1E+300     ' rectangle of the position "not placed" (its bounding box meets nothing)

Private Type Poloha
    Cena As Double
    L As Double
    T As Double
    R As Obd               ' label rectangle
    C As Cesta             ' leader line, C.N = 0: none
    Seg(1 To 2) As Obd     ' bounding boxes of the leader line segments
    Obal As Obd            ' label and leader line, padded by OBAL
    Strana As Long         ' grid key: 0 none, 1 right of the point, 2 left, 3 own column
    A As Long              ' index into mDX
    K As Long              ' grid row
    Puvodni As Boolean     ' fallback: the position "not placed" (original position of the label)
End Type

Private Type Moznosti
    P() As Poloha          ' sorted by cost (stable)
    Pocet As Long
    Start As Long          ' the position from the columns
    K0 As Long             ' first grid row of the key table
    NK As Long             ' number of grid rows
    Klice() As Long        ' grid key -> index into P, 0 = none
End Type

Private mK() As Moznosti
Private mVolba() As Long
Private mN As Long
Private mMez As Double
Private mSeed As Double
Private mDX(0 To POCET_DX - 1) As Double
Private mDV(0 To POCET_DV - 1) As Double

Private mLadeni As String       ' JSON for the parity test: candidate counts, starts, choices after each phase
Private mNouze As Boolean       ' fallback placement (Nouzove)
Private mPresCaru As Double
Private mPuvL() As Double, mPuvT() As Double, mPuvW() As Double, mPuvH() As Double  ' original label positions

' generated positions of one label, before the static conditions
Private mGn As Long
Private mGl() As Double, mGt() As Double, mGp() As Double
Private mGs() As Long, mGa() As Long, mGk() As Long

' Pulls the labels from the column placement (gP L/T, W/H) towards their points; mezera = gap used by the columns.
' Works on the chart rounded to the grid; afterwards everything is restored except the new L/T (on the grid).
' nouze: fallback from invalid columns (original positions in mPuv*, set by Nouzove); labels not placed get
' Rezim "N" and their original position on the grid. Returns the number of labels not placed.
Public Function ZlepsiBlizko(ByVal mezera As Double, Optional ByVal nouze As Boolean = False) As Long
    Dim i As Long, vyska As Double, krok As Double, starty() As Long, p As Poloha
    Dim puv() As Popisek, puvPrek() As Obd, puvCary() As Double
    Dim pl As Double, pt As Double, pr As Double, pb As Double, rez As Double, del As Double
    mN = UBound(gP)
    If mN = 0 Then Exit Function
    mNouze = nouze
    If nouze Then mPresCaru = PRES_CARU_NOUZE Else mPresCaru = PRES_CARU
    Konstanty
    puv = gP
    puvPrek = gPrek
    puvCary = gCary
    pl = gPL
    pt = gPT
    pr = gPR
    pb = gPB
    rez = gRezerva
    del = gDelici
    gPL = Mrizka(gPL)
    gPT = Mrizka(gPT)
    gPR = Mrizka(gPR)
    gPB = Mrizka(gPB)
    gRezerva = Mrizka(gRezerva)
    If gMaDelici Then gDelici = Mrizka(gDelici)
    For i = 1 To UBound(gCary)
        gCary(i) = Mrizka(gCary(i))
    Next i
    For i = 1 To UBound(gPrek)
        gPrek(i) = NovyObd(Mrizka(gPrek(i).L), Mrizka(gPrek(i).T), Mrizka(gPrek(i).R), Mrizka(gPrek(i).B))
    Next i
    For i = 1 To mN
        gP(i).X = Mrizka(gP(i).X)
        gP(i).Y = Mrizka(gP(i).Y)
        gP(i).W = Mrizka(gP(i).W)
        gP(i).H = Mrizka(gP(i).H)
        gP(i).L = Mrizka(gP(i).L)
        gP(i).T = Mrizka(gP(i).T)
    Next i
    mMez = mezera - REZERVA_MEZERY  ' conflict = labels closer than the gap the columns used
    vyska = gP(1).H
    For i = 1 To mN
        vyska = Max2(vyska, gP(i).H)
    Next i
    krok = (vyska + mezera) / 2
    ReDim mK(1 To mN)
    ReDim mVolba(1 To mN)
    For i = 1 To mN
        NajdiPolohy i, krok
        mVolba(i) = mK(i).Start
    Next i
    mLadeni = "{""pocty"":" & SeznamJson(0) & ",""starty"":" & SeznamJson(1)
    starty = mVolba
    Hledani
    mLadeni = mLadeni & ",""po_hledani"":" & SeznamJson(2)
    Doladeni
    If nouze Then  ' labels not placed to their original position, the placed ones move away; polish may place them
        For i = 1 To mN
            If mK(i).P(starty(i)).Puvodni Then
                gP(i).L = Mrizka(mPuvL(i))
                gP(i).T = Mrizka(mPuvT(i))
                gP(i).W = Mrizka(mPuvW(i))
                gP(i).H = Mrizka(mPuvH(i))
                VyplnPolohu p, NEUMISTEN, gP(i).L, gP(i).T, gP(i).W, gP(i).H, VodiciCara(i), 0, 0, 0
                p.Puvodni = True
                mK(i).P(starty(i)) = p
            End If
        Next i
        Uvolni
        Doladeni
    End If
    mLadeni = mLadeni & ",""konec"":" & SeznamJson(2) & "}"
    gP = puv
    gPrek = puvPrek
    gCary = puvCary
    gPL = pl
    gPT = pt
    gPR = pr
    gPB = pb
    gRezerva = rez
    gDelici = del
    For i = 1 To mN
        With mK(i).P(mVolba(i))
            gP(i).L = .L
            gP(i).T = .T
            If .Puvodni Then
                gP(i).Rezim = "N"
                ZlepsiBlizko = ZlepsiBlizko + 1
            ElseIf .C.N = 0 Then
                gP(i).Rezim = "D"
            Else
                gP(i).Rezim = "B"
            End If
        End With
    Next i
    Erase mK
End Function

' Fallback placement when the columns do not fit (after Rozmisti failed, labels have single-line W/H).
' puv* = original L/T/W/H of the labels. Returns the gap used, or -1 when fewer than half of the labels would
' be placed (the chart is then left unchanged). Labels not placed get Rezim "N" and their original position
' and size. startL/startT = column start of the chosen attempt (for the parity test).
Public Function Nouzove(puvL() As Double, puvT() As Double, puvW() As Double, puvH() As Double, _
                        startL() As Double, startT() As Double) As Double
    Dim mezera As Variant, i As Long, n As Long, pocet As Long, nej As Long, nejMezera As Double, nejLadeni As String
    Dim stL() As Double, stT() As Double, kL() As Double, kT() As Double, kR() As String
    n = UBound(gP)
    mPuvL = puvL
    mPuvT = puvT
    mPuvW = puvW
    mPuvH = puvH
    ReDim stL(1 To n)
    ReDim stT(1 To n)
    ReDim kL(1 To n)
    ReDim kT(1 To n)
    ReDim kR(1 To n)
    nej = -1
    For Each mezera In Array(1.5, 0.5, 0#)  ' the gaps of Rozlozeni.Rozmisti
        StartSloupce CDbl(mezera)  ' the cheapest column split, even when it does not fit
        For i = 1 To n
            stL(i) = gP(i).L
            stT(i) = gP(i).T
        Next i
        pocet = ZlepsiBlizko(CDbl(mezera), True)
        If nej < 0 Or pocet < nej Then
            nej = pocet
            nejMezera = mezera
            nejLadeni = mLadeni
            startL = stL
            startT = stT
            For i = 1 To n
                kL(i) = gP(i).L
                kT(i) = gP(i).T
                kR(i) = gP(i).Rezim
            Next i
        End If
        If pocet = 0 Then Exit For
    Next mezera
    mLadeni = nejLadeni
    If 2 * nej > n Then
        Nouzove = -1
        Exit Function
    End If
    For i = 1 To n
        gP(i).L = kL(i)
        gP(i).T = kT(i)
        gP(i).Rezim = kR(i)
        If kR(i) = "N" Then
            gP(i).L = puvL(i)
            gP(i).T = puvT(i)
            gP(i).W = puvW(i)
            gP(i).H = puvH(i)
        End If
    Next i
    Nouzove = nejMezera
End Function

' Rounding to the grid 1/MRIZ: v * MRIZ is exact and so is + 0.5 below 2^42, same bits as Python's floor.
Private Function Mrizka(ByVal v As Double) As Double
    Mrizka = Int(v * MRIZ + 0.5) / MRIZ
End Function

' For the parity test (Popisky.ModelJson): counts, starts and choices of the last chart, 0-based like Python.
Public Function Ladeni() As String
    Ladeni = mLadeni
End Function

' co: 0 candidate counts, 1 start indexes, 2 current choices (indexes 0-based like Python)
Private Function SeznamJson(ByVal co As Long) As String
    Dim i As Long, s As String
    For i = 1 To mN
        If i > 1 Then s = s & ","
        Select Case co
            Case 0: s = s & mK(i).Pocet
            Case 1: s = s & (mK(i).Start - 1)
            Case Else: s = s & (mVolba(i) - 1)
        End Select
    Next i
    SeznamJson = "[" & s & "]"
End Function

Private Sub Konstanty()
    mDX(0) = 12#: mDX(1) = 18#: mDX(2) = 25#: mDX(3) = 33#: mDX(4) = 42#: mDX(5) = 53#: mDX(6) = 66#
    mDX(7) = 82#: mDX(8) = 100#: mDX(9) = 122#: mDX(10) = 148#: mDX(11) = 180#: mDX(12) = 220#: mDX(13) = 270#
    mDV(0) = 12#: mDV(1) = 19#: mDV(2) = 28#: mDV(3) = 39#: mDV(4) = 52#: mDV(5) = 68#: mDV(6) = 86#: mDV(7) = 108#
End Sub

' ---------- candidates ----------
Private Sub Generuj(ByVal l As Double, ByVal t As Double, ByVal pref As Double, ByVal strana As Long, _
                    ByVal a As Long, ByVal k As Long)
    mGn = mGn + 1
    If mGn > UBound(mGl) Then
        ReDim Preserve mGl(1 To 2 * mGn)
        ReDim Preserve mGt(1 To 2 * mGn)
        ReDim Preserve mGp(1 To 2 * mGn)
        ReDim Preserve mGs(1 To 2 * mGn)
        ReDim Preserve mGa(1 To 2 * mGn)
        ReDim Preserve mGk(1 To 2 * mGn)
    End If
    mGl(mGn) = l
    mGt(mGn) = t
    mGp(mGn) = pref
    mGs(mGn) = strana
    mGa(mGn) = a
    mGk(mGn) = k
End Sub

' Candidate positions of label i, same order as popisky/blizko.py _kandidati, then a stable sort by cost.
Private Sub NajdiPolohy(ByVal i As Long, ByVal krok As Double)
    Dim x As Double, y As Double, w As Double, h As Double, l0 As Double, t0 As Double, horni As Double
    Dim k0 As Long, k1 As Long, k As Long, a As Long, d As Long, m As Long, c As Long, n As Long
    Dim t As Double, sik As Double, posun(1 To 3) As Double, r As Obd, platny As Boolean
    Dim pp() As Poloha, cena() As Double, idx() As Long, p As Poloha
    x = gP(i).X
    y = gP(i).Y
    w = gP(i).W
    h = gP(i).H
    l0 = gP(i).L
    t0 = gP(i).T
    mGn = 0
    ReDim mGl(1 To 256)
    ReDim mGt(1 To 256)
    ReDim mGp(1 To 256)
    ReDim mGs(1 To 256)
    ReDim mGa(1 To 256)
    ReDim mGk(1 To 256)
    Generuj x + 6, y - h / 2, 0#, 0, 0, 0
    Generuj x - 6 - w, y - h / 2, 0.5, 0, 0, 0
    Generuj x - w / 2, y - 6 - h, 1#, 0, 0, 0
    Generuj x - w / 2, y + 6, 1.5, 0, 0, 0
    Generuj x + 3, y - 3 - h, 2#, 0, 0, 0
    Generuj x + 3, y + 3, 2.5, 0, 0, 0
    Generuj x - 3 - w, y - 3 - h, 3#, 0, 0, 0
    Generuj x - 3 - w, y + 3, 3.5, 0, 0, 0
    horni = gPT + Max2(OKRAJ, gRezerva)
    k0 = -Int(-((y - h / 2 - POSUN_MAX - horni) / krok))
    If k0 < 0 Then k0 = 0
    k1 = Int((y - h / 2 + POSUN_MAX - horni) / krok)
    For a = 0 To POCET_DX - 1
        For k = k0 To k1
            t = horni + k * krok
            sik = SIKMO * Abs(t + h / 2 - y)
            Generuj x + mDX(a), t, sik, 1, a, k
            Generuj x - mDX(a) - w, t, VLEVO + sik, 2, a, k
        Next k
    Next a
    For k = k0 To k1  ' elsewhere in its own column
        t = horni + k * krok
        Generuj l0, t, SIKMO * Abs(t + h / 2 - y), 3, 0, k
    Next k
    sik = SIKMO * Abs(t0 + h / 2 - y)
    If l0 > x Then  ' closer to the point in its own row (column on the right)
        For d = 12 To Int(l0 - x) - 1 Step 3
            Generuj x + d, t0, sik, 0, 0, 0
        Next d
    ElseIf l0 + w < x Then  ' column on the left
        For d = 12 To Int(x - l0 - w) - 1 Step 3
            Generuj x - d - w, t0, VLEVO + sik, 0, 0, 0
        Next d
    End If
    posun(1) = 0#
    posun(2) = -w / 2 + 8
    posun(3) = w / 2 - 8
    For d = 0 To POCET_DV - 1
        For m = 1 To 3
            Generuj x - w / 2 + posun(m), y - mDV(d) - h, NAD_POD, 0, 0, 0
            Generuj x - w / 2 + posun(m), y + mDV(d), NAD_POD, 0, 0, 0
        Next m
    Next d

    ReDim pp(1 To mGn + 1)
    For m = 1 To mGn
        r = NovyObd(mGl(m), mGt(m), mGl(m) + w, mGt(m) + h)
        If Volno(i, r) Then
            If NovaPoloha(i, mGl(m), mGt(m), mGp(m), mGs(m), mGa(m), mGk(m), p) Then
                n = n + 1
                pp(n) = p
            End If
        End If
    Next m
    platny = NovaPoloha(i, l0, t0, 0#, 0, 0, 0, p)
    If mNouze Then  ' fallback: a start outside the static conditions or with its leader across a box is "not placed"
        If Not platny Then
            PolohaNeumisten p
        ElseIf Not Volno(i, NovyObd(l0, t0, l0 + w, t0 + h)) Then
            PolohaNeumisten p
        End If
    ElseIf Not platny Then  ' leader from the column across a box: valid but expensive
        gP(i).L = l0
        gP(i).T = t0
        VyplnPolohu p, 1000000#, l0, t0, w, h, VodiciCara(i), 0, 0, 0
    End If
    n = n + 1
    pp(n) = p

    ReDim cena(1 To n)
    ReDim idx(1 To n)
    For m = 1 To n
        cena(m) = pp(m).Cena
        idx(m) = m
    Next m
    SeradStabilne idx, cena, n
    With mK(i)
        ReDim .P(1 To n)
        .Pocet = n
        For m = 1 To n
            .P(m) = pp(idx(m))
            If idx(m) = n Then .Start = m
        Next m
        .K0 = k0
        .NK = k1 - k0 + 1
        If .NK < 0 Then .NK = 0
        If .NK > 0 Then
            ReDim .Klice(0 To 3 * POCET_DX * .NK - 1)
            For c = 1 To n
                If .P(c).Strana > 0 Then .Klice(((.P(c).Strana - 1) * POCET_DX + .P(c).A) * .NK + (.P(c).K - .K0)) = c
            Next c
        End If
    End With
End Sub

' Static conditions: inside the plot area, not across vertical lines, in its half, clear of boxes and other points.
Private Function Volno(ByVal i As Long, r As Obd) As Boolean
    Dim j As Long
    If Not VPlose(r) Then Exit Function
    If PresCary(r) > 0 Then Exit Function
    If gMaDelici Then
        If gP(i).Y < gDelici Then
            If r.B > gDelici - 2 Then Exit Function
        ElseIf r.T < gDelici + 2 Then
            Exit Function
        End If
    End If
    For j = 1 To UBound(gPrek)
        If Prekryv(r, gPrek(j), 2#) Then Exit Function
    Next j
    For j = 1 To UBound(gP)
        If j <> i Then
            If Prekryv(r, NovyObd(gP(j).X - 3 - 2, gP(j).Y - 3 - 2, gP(j).X + 3 + 2, gP(j).Y + 3 + 2)) Then Exit Function
        End If
    Next j
    Volno = True
End Function

' Position with its cost (preference + leader line length + leader across vertical lines); False = leader across a box.
Private Function NovaPoloha(ByVal i As Long, ByVal l As Double, ByVal t As Double, ByVal pref As Double, _
                            ByVal strana As Long, ByVal a As Long, ByVal k As Long, p As Poloha) As Boolean
    Dim c As Cesta, s As Long, m As Long, cena As Double, xa As Double, xb As Double, pocet As Long
    gP(i).L = l
    gP(i).T = t
    c = VodiciCara(i)
    For s = 1 To c.N - 1
        For m = 1 To UBound(gPrek)
            If UseckaPres(c.X(s), c.Y(s), c.X(s + 1), c.Y(s + 1), gPrek(m), 0#) Then Exit Function
        Next m
    Next s
    cena = pref
    For s = 1 To c.N - 1
        cena = cena + Sqr((c.X(s + 1) - c.X(s)) * (c.X(s + 1) - c.X(s)) + (c.Y(s + 1) - c.Y(s)) * (c.Y(s + 1) - c.Y(s)))
    Next s
    If c.N > 0 Then
        xa = Min2(gP(i).X, l)
        xb = Max2(gP(i).X, l + gP(i).W)
        For m = 1 To UBound(gCary)
            If xa < gCary(m) And gCary(m) < xb Then pocet = pocet + 1
        Next m
        cena = cena + mPresCaru * pocet
    End If
    VyplnPolohu p, Mrizka(cena), l, t, gP(i).W, gP(i).H, c, strana, a, k
    NovaPoloha = True
End Function

Private Sub VyplnPolohu(p As Poloha, ByVal cena As Double, ByVal l As Double, ByVal t As Double, ByVal w As Double, _
                        ByVal h As Double, c As Cesta, ByVal strana As Long, ByVal a As Long, ByVal k As Long)
    Dim s As Long, ol As Double, ot As Double, orr As Double, ob As Double
    p.Cena = cena
    p.L = l
    p.T = t
    p.R = NovyObd(l, t, l + w, t + h)
    p.C = c
    For s = 1 To c.N - 1
        p.Seg(s) = NovyObd(Min2(c.X(s), c.X(s + 1)), Min2(c.Y(s), c.Y(s + 1)), Max2(c.X(s), c.X(s + 1)), Max2(c.Y(s), c.Y(s + 1)))
    Next s
    ol = p.R.L
    ot = p.R.T
    orr = p.R.R
    ob = p.R.B
    For s = 1 To c.N
        ol = Min2(ol, c.X(s))
        ot = Min2(ot, c.Y(s))
        orr = Max2(orr, c.X(s))
        ob = Max2(ob, c.Y(s))
    Next s
    p.Obal = NovyObd(ol - OBAL, ot - OBAL, orr + OBAL, ob + OBAL)
    p.Strana = strana
    p.A = a
    p.K = k
    p.Puvodni = False
End Sub

' Fallback: the position "not placed" for the search, without a rectangle and a leader line (no conflicts).
Private Sub PolohaNeumisten(p As Poloha)
    Dim zadna As Cesta
    VyplnPolohu p, NEUMISTEN, NIKDE, NIKDE, 0#, 0#, zadna, 0, 0, 0
    p.R = NovyObd(NIKDE, NIKDE, -NIKDE, -NIKDE)
    p.Obal = p.R
    p.Puvodni = True
End Sub

' Stable bottom-up merge sort of indexes by key (as Python's sort).
Private Sub SeradStabilne(idx() As Long, klic() As Double, ByVal n As Long)
    Dim tmp() As Long, sirka As Long, lo As Long, st As Long, hi As Long, a As Long, b As Long, k As Long
    If n < 2 Then Exit Sub
    ReDim tmp(1 To n)
    sirka = 1
    Do While sirka < n
        lo = 1
        Do While lo <= n
            st = lo + sirka - 1
            If st > n Then st = n
            hi = lo + 2 * sirka - 1
            If hi > n Then hi = n
            a = lo
            b = st + 1
            k = lo
            Do While a <= st And b <= hi
                If klic(idx(b)) < klic(idx(a)) Then
                    tmp(k) = idx(b)
                    b = b + 1
                Else
                    tmp(k) = idx(a)
                    a = a + 1
                End If
                k = k + 1
            Loop
            Do While a <= st
                tmp(k) = idx(a)
                a = a + 1
                k = k + 1
            Loop
            Do While b <= hi
                tmp(k) = idx(b)
                b = b + 1
                k = k + 1
            Loop
            lo = lo + 2 * sirka
        Loop
        For k = 1 To n
            idx(k) = tmp(k)
        Next k
        sirka = sirka * 2
    Loop
End Sub

Private Function Najdi(ByVal i As Long, ByVal strana As Long, ByVal a As Long, ByVal k As Long) As Long
    With mK(i)
        If .NK = 0 Or k < .K0 Or k >= .K0 + .NK Then Exit Function
        Najdi = .Klice(((strana - 1) * POCET_DX + a) * .NK + (k - .K0))
    End With
End Function

' ---------- conflicts ----------
' Same as Graf.Krizeni (proper intersection of two segments), expressions in the same order.
Private Function Kriz(ByVal p1x As Double, ByVal p1y As Double, ByVal p2x As Double, ByVal p2y As Double, _
                      ByVal q1x As Double, ByVal q1y As Double, ByVal q2x As Double, ByVal q2y As Double) As Boolean
    If ((p2x - p1x) * (q1y - p1y) - (p2y - p1y) * (q1x - p1x)) * _
       ((p2x - p1x) * (q2y - p1y) - (p2y - p1y) * (q2x - p1x)) >= 0 Then Exit Function
    Kriz = ((q2x - q1x) * (p1y - q1y) - (q2y - q1y) * (p1x - q1x)) * _
           ((q2x - q1x) * (p2y - q1y) - (q2y - q1y) * (p2x - q1x)) < 0
End Function

' Segment s of the leader line of p across rectangle r (Graf.UseckaPres with zmenseni 0), quick bounding box test first.
Private Function Pres(p As Poloha, ByVal s As Long, r As Obd) As Boolean
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    If p.Seg(s).R < r.L Or p.Seg(s).L > r.R Or p.Seg(s).B < r.T Or p.Seg(s).T > r.B Or r.L >= r.R Or r.T >= r.B Then
        Exit Function
    End If
    x1 = p.C.X(s)
    y1 = p.C.Y(s)
    x2 = p.C.X(s + 1)
    y2 = p.C.Y(s + 1)
    If r.L <= x1 And x1 <= r.R And r.T <= y1 And y1 <= r.B Then
        Pres = True
    ElseIf r.L <= x2 And x2 <= r.R And r.T <= y2 And y2 <= r.B Then
        Pres = True
    ElseIf Kriz(x1, y1, x2, y2, r.L, r.T, r.R, r.T) Then
        Pres = True
    ElseIf Kriz(x1, y1, x2, y2, r.R, r.T, r.R, r.B) Then
        Pres = True
    ElseIf Kriz(x1, y1, x2, y2, r.R, r.B, r.L, r.B) Then
        Pres = True
    Else
        Pres = Kriz(x1, y1, x2, y2, r.L, r.B, r.L, r.T)
    End If
End Function

' Leader segments s of a and u of b cross, quick bounding box test first.
Private Function Krizi(a As Poloha, ByVal s As Long, b As Poloha, ByVal u As Long) As Boolean
    If a.Seg(s).R < b.Seg(u).L Or b.Seg(u).R < a.Seg(s).L Or a.Seg(s).B < b.Seg(u).T Or b.Seg(u).B < a.Seg(s).T Then
        Exit Function
    End If
    Krizi = Kriz(a.C.X(s), a.C.Y(s), a.C.X(s + 1), a.C.Y(s + 1), b.C.X(u), b.C.Y(u), b.C.X(u + 1), b.C.Y(u + 1))
End Function

' Conflicts of two positions: labels closer than the gap, leader line across the other label, crossing leader lines.
' Two original positions (fallback) never conflict: their conflicts are the user's.
Private Function Konflikty(a As Poloha, b As Poloha) As Long
    Dim s As Long, u As Long, n As Long
    If a.Puvodni And b.Puvodni Then Exit Function
    If a.Obal.L >= b.Obal.R Then Exit Function
    If b.Obal.L >= a.Obal.R Then Exit Function
    If a.Obal.T >= b.Obal.B Then Exit Function
    If b.Obal.T >= a.Obal.B Then Exit Function
    If a.R.L < b.R.R + mMez Then
        If b.R.L - mMez < a.R.R Then
            If a.R.T < b.R.B + mMez Then
                If b.R.T - mMez < a.R.B Then n = 1
            End If
        End If
    End If
    For s = 1 To a.C.N - 1
        If Pres(a, s, b.R) Then n = n + 1
        For u = 1 To b.C.N - 1
            If Krizi(a, s, b, u) Then n = n + 1
        Next u
    Next s
    For u = 1 To b.C.N - 1
        If Pres(b, u, a.R) Then n = n + 1
    Next u
    Konflikty = n
End Function

' Conflicts of position p of label i with the other labels (except i and krome) in their chosen positions.
Private Function KonfliktyS(ByVal i As Long, p As Poloha, Optional ByVal krome As Long = 0) As Long
    Dim j As Long
    For j = 1 To mN
        If j <> i And j <> krome Then KonfliktyS = KonfliktyS + Konflikty(p, mK(j).P(mVolba(j)))
    Next j
End Function

' MINSTD (Park-Miller): integers below 2^47, exact in Double.
Private Function Nahoda() As Double
    mSeed = mSeed * 48271#
    mSeed = mSeed - Int(mSeed / 2147483647#) * 2147483647#
    Nahoda = mSeed / 2147483647#
End Function

' ---------- search ----------
' Threshold accepting from the start (columns); leaves the best choice without conflicts (or the start) in mVolba.
Private Sub Hledani()
    Dim kroku As Long, s As Long, i As Long, j As Long, c As Long, ci As Long, cj As Long
    Dim si As Long, sj As Long, ki As Long, kj As Long, dk As Long, konfliktu As Long
    Dim prah As Double, pokuta As Double, v As Double, dc As Double, cena As Double, nejCena As Double
    Dim start() As Long, nej() As Long
    kroku = KROKU_NA_POPISEK * mN
    mSeed = 12345#
    For i = 1 To mN
        cena = cena + mK(i).P(mVolba(i)).Cena
    Next i
    For i = 1 To mN
        For j = i + 1 To mN
            konfliktu = konfliktu + Konflikty(mK(i).P(mVolba(i)), mK(j).P(mVolba(j)))
        Next j
    Next i
    start = mVolba
    nej = mVolba
    nejCena = cena
    For s = 0 To kroku - 1
        prah = Int((PRAH0 * (kroku - s) + PRAH1 * s) * MRIZ / kroku)  ' integers in 1/MRIZ, exact
        pokuta = Int((POKUTA0 * (kroku - s) + POKUTA1 * s) * MRIZ / kroku)
        i = Int(Nahoda() * mN) + 1
        If Nahoda() < PROHOZENI Then  ' swap the rows of two labels (reorders a stack without an intermediate step)
            j = Int(Nahoda() * mN) + 1
            If j = i Then GoTo Dalsi
            si = mK(i).P(mVolba(i)).Strana
            sj = mK(j).P(mVolba(j)).Strana
            If si = 0 Or sj = 0 Then GoTo Dalsi
            ki = mK(i).P(mVolba(i)).K
            kj = mK(j).P(mVolba(j)).K
            If ki = kj Then GoTo Dalsi
            ci = Najdi(i, si, mK(i).P(mVolba(i)).A, kj)
            cj = Najdi(j, sj, mK(j).P(mVolba(j)).A, ki)
            If ci = 0 Or cj = 0 Then GoTo Dalsi
            dk = KonfliktyS(i, mK(i).P(ci), j) + KonfliktyS(j, mK(j).P(cj), i) + Konflikty(mK(i).P(ci), mK(j).P(cj)) _
                 - KonfliktyS(i, mK(i).P(mVolba(i)), j) - KonfliktyS(j, mK(j).P(mVolba(j)), i) _
                 - Konflikty(mK(i).P(mVolba(i)), mK(j).P(mVolba(j)))
            dc = mK(i).P(ci).Cena + mK(j).P(cj).Cena - mK(i).P(mVolba(i)).Cena - mK(j).P(mVolba(j)).Cena
            If dc * MRIZ + pokuta * dk < prah Then
                mVolba(i) = ci
                mVolba(j) = cj
                cena = cena + dc
                konfliktu = konfliktu + dk
            End If
        Else
            v = Nahoda()
            c = Int(v * v * mK(i).Pocet) + 1  ' cheap positions more often (candidates are sorted by cost)
            If c = mVolba(i) Then GoTo Dalsi
            dk = KonfliktyS(i, mK(i).P(c)) - KonfliktyS(i, mK(i).P(mVolba(i)))
            dc = mK(i).P(c).Cena - mK(i).P(mVolba(i)).Cena
            If dc * MRIZ + pokuta * dk < prah Then
                mVolba(i) = c
                cena = cena + dc
                konfliktu = konfliktu + dk
            End If
        End If
        If konfliktu = 0 And cena < nejCena - 0.000000001 Then
            nej = mVolba
            nejCena = cena
        End If
Dalsi:
    Next s
    If konfliktu > 0 Then
        If Oprava(start) Then
            cena = 0#
            For i = 1 To mN
                cena = cena + mK(i).P(mVolba(i)).Cena
            Next i
            If cena < nejCena - 0.000000001 Then nej = mVolba
        End If
    End If
    mVolba = nej
End Sub

Private Function VolnaPoloha(ByVal i As Long) As Long
    Dim c As Long
    For c = 1 To mK(i).Pocet
        If KonfliktyS(i, mK(i).P(c)) = 0 Then
            VolnaPoloha = c
            Exit Function
        End If
    Next c
End Function

' Removes the remaining conflicts. The label with the most conflicts (not at its start) moves to its cheapest
' position without a conflict, otherwise to a position blocked by a single label that moves to its cheapest free
' position, otherwise back to its start (the columns have no conflicts among themselves). True = done.
Private Function Oprava(start() As Long) As Boolean
    Dim kolo As Long, i As Long, j As Long, c As Long, cj As Long, kdo As Long, prekazi As Long, puvodni As Long
    Dim mez As Long, pocty() As Long, nic As Boolean, hotovo As Boolean
    ReDim pocty(1 To mN)
    For kolo = 1 To 4 * mN
        nic = True
        For j = 1 To mN
            pocty(j) = KonfliktyS(j, mK(j).P(mVolba(j)))
            If pocty(j) > 0 Then nic = False
        Next j
        If nic Then
            Oprava = True
            Exit Function
        End If
        i = 0
        For j = 1 To mN
            If pocty(j) > 0 And mVolba(j) <> start(j) Then
                If i = 0 Then
                    i = j
                ElseIf pocty(j) > pocty(i) Then
                    i = j
                End If
            End If
        Next j
        If i = 0 Then Exit Function
        c = VolnaPoloha(i)
        If c > 0 Then
            mVolba(i) = c
        Else
            hotovo = False
            mez = OPRAVA_KANDIDATU
            If mK(i).Pocet < mez Then mez = mK(i).Pocet
            For c = 1 To mez
                prekazi = 0
                For j = 1 To mN
                    If j <> i Then
                        If Konflikty(mK(i).P(c), mK(j).P(mVolba(j))) > 0 Then
                            prekazi = prekazi + 1
                            kdo = j
                        End If
                    End If
                Next j
                If prekazi = 1 Then
                    puvodni = mVolba(i)
                    mVolba(i) = c
                    cj = VolnaPoloha(kdo)
                    If cj > 0 Then
                        mVolba(kdo) = cj
                        hotovo = True
                        Exit For
                    End If
                    mVolba(i) = puvodni
                End If
            Next c
            If Not hotovo Then mVolba(i) = start(i)
        End If
    Next kolo
End Function

' Fallback: a placed label in conflict (mostly with the original position of a label not placed) moves to its
' cheapest free position, otherwise to a position blocked by a single placed label that moves to its free
' position; otherwise it stays. Original positions do not move.
Private Sub Uvolni()
    Dim kolo As Long, i As Long, j As Long, c As Long, cj As Long, kdo As Long, prekazi As Long, puvodni As Long
    Dim mez As Long, pocty() As Long, vzdat() As Boolean, hotovo As Boolean
    ReDim pocty(1 To mN)
    ReDim vzdat(1 To mN)
    For j = 1 To mN
        vzdat(j) = mK(j).P(mVolba(j)).Puvodni
    Next j
    For kolo = 1 To 4 * mN
        For j = 1 To mN
            pocty(j) = KonfliktyS(j, mK(j).P(mVolba(j)))
        Next j
        i = 0
        For j = 1 To mN
            If pocty(j) > 0 And Not vzdat(j) Then
                If i = 0 Then
                    i = j
                ElseIf pocty(j) > pocty(i) Then
                    i = j
                End If
            End If
        Next j
        If i = 0 Then Exit Sub
        c = VolnaPoloha(i)
        If c > 0 Then
            mVolba(i) = c
        Else
            hotovo = False
            mez = OPRAVA_KANDIDATU
            If mK(i).Pocet < mez Then mez = mK(i).Pocet
            For c = 1 To mez
                prekazi = 0
                For j = 1 To mN
                    If j <> i Then
                        If Konflikty(mK(i).P(c), mK(j).P(mVolba(j))) > 0 Then
                            prekazi = prekazi + 1
                            kdo = j
                        End If
                    End If
                Next j
                If prekazi = 1 Then
                    If Not vzdat(kdo) Then
                        puvodni = mVolba(i)
                        mVolba(i) = c
                        cj = VolnaPoloha(kdo)
                        If cj > 0 Then
                            mVolba(kdo) = cj
                            hotovo = True
                            Exit For
                        End If
                        mVolba(i) = puvodni
                    End If
                End If
            Next c
            If Not hotovo Then vzdat(i) = True
        End If
    Next kolo
End Sub

' Labels from the most expensive one move to the cheapest position that creates no conflict, while anything improves.
Private Sub Doladeni()
    Dim kolo As Long, m As Long, i As Long, c As Long, zmena As Boolean, poradi() As Long, klic() As Double
    ReDim poradi(0 To mN)
    ReDim klic(0 To mN)
    For kolo = 1 To DOLADENI_KOL
        zmena = False
        For i = 1 To mN
            poradi(i) = i
            klic(i) = -mK(i).P(mVolba(i)).Cena
        Next i
        SeradIndexy poradi, klic
        For m = 1 To mN
            i = poradi(m)
            For c = 1 To mK(i).Pocet
                If mK(i).P(c).Cena >= mK(i).P(mVolba(i)).Cena - 0.000000001 Then Exit For
                If KonfliktyS(i, mK(i).P(c)) = 0 Then
                    mVolba(i) = c
                    zmena = True
                    Exit For
                End If
            Next c
        Next m
        If Not zmena Then Exit Sub
    Next kolo
End Sub
