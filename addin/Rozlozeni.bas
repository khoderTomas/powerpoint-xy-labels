Option Explicit

' Label placement (port of popisky/rozlozeni.py; expressions in the same order so that rounding matches):
' 1. Direct label (no leader line) where there is free space next to the point and it does not cross
'    an Index line or the dividing line.
' 2. The rest go into columns: every left/right split by x is tried and the cheapest one wins
'    (leader line length + labels across lines + does not fit). A column is split by the quadrant dividing
'    line, ordered without crossing leader lines (2-opt over slots), placed as close to the points as possible
'    (least squares), going around obstacles.
' 3. A direct label crossed by a leader line or overlapped by a column moves to a column and the placement
'    is computed again.
' Label lists are arrays of indexes into gP starting at 1 (element 0 unused); sorting is stable as in Python.

Private Const MEZERA_SLOUPCE As Double = 14#   ' outermost point -> column edge
Public Const OKRAJ As Double = 2#              ' distance from the edge of the inner plot area (also Blizko)
Private Const EPS As Double = 0.000001
Private Const NEVEJDE_SE As Double = 100000#

' Finds the placement; returns the gap used, or -1 when the labels do not fit (the chart is then left unchanged).
Public Function Rozmisti() As Double
    Dim i As Long, w As Double, h As Double, mezera As Variant
    For i = 1 To UBound(gP)  ' long labels on a single line (custom width) so that columns stay aligned
        Jednoradkovy i, w, h
        gP(i).W = w
        gP(i).H = h
        gP(i).VlastniSirka = gP(i).W > ZALOMENI * gW - 1#  ' 1 pt margin at the wrapping limit
        If gP(i).VlastniSirka Then gP(i).W = gP(i).W + 0.5  ' margin against wrapping caused by rounding
    Next i
    For Each mezera In Array(1.5, 0.5, 0#)  ' vertical gap between boxes; smaller only when the labels would not fit otherwise
        If Jednou(CDbl(mezera)) Then
            Rozmisti = mezera
            Exit Function
        End If
    Next mezera
    Rozmisti = -1#
End Function

' Start of the fallback placement (Blizko.Nouzove): the cheapest column split for the gap, even when it does not fit.
Public Sub StartSloupce(ByVal mezera As Double)
    Dim ok As Boolean
    ok = Jednou(mezera)
End Sub

Private Function Jednou(ByVal mezera As Double) As Boolean
    Dim n As Long, i As Long, j As Long, m As Long, opak As Long, l As Double, t As Double, dx As Double, dy As Double
    n = UBound(gP)
    Dim poradi() As Long, klic() As Double, zakazany() As Boolean, spatny() As Boolean, uml() As Long, zbytek() As Long
    ReDim poradi(0 To n)
    ReDim klic(0 To n)
    ReDim zakazany(0 To n)
    For i = 1 To n  ' labels with the fewest neighbours first
        poradi(i) = i
        For j = 1 To n
            If j <> i Then
                dx = gP(i).X - gP(j).X
                dy = gP(i).Y - gP(j).Y
                If Sqr(dx * dx + dy * dy) < 40 Then klic(i) = klic(i) + 1
            End If
        Next j
    Next i
    SeradIndexy poradi, klic
    For opak = 1 To 15
        ReDim uml(0 To 0)
        For m = 1 To n
            i = poradi(m)
            gP(i).Rezim = ""
            If Not zakazany(i) Then
                If Primo(i, uml, l, t) Then
                    gP(i).L = l
                    gP(i).T = t
                    gP(i).Rezim = "D"
                    PridejIndex uml, i
                End If
            End If
        Next m
        ReDim zbytek(0 To 0)
        For i = 1 To n
            If gP(i).Rezim <> "D" Then PridejIndex zbytek, i
        Next i
        Jednou = Sloupce(zbytek, uml, mezera)
        ReDim spatny(0 To n)
        Dim nalez As Boolean, rd As Obd, c As Cesta
        nalez = False
        For m = 1 To UBound(uml)
            rd = Obdelnik(uml(m))
            For j = 1 To UBound(zbytek)
                c = VodiciCara(zbytek(j))
                If Prekryv(rd, Obdelnik(zbytek(j)), 1#) Or CestaPres(c, rd, 0#) Then
                    spatny(uml(m)) = True
                    nalez = True
                End If
            Next j
        Next m
        If Not nalez Then Exit Function
        For i = 1 To n
            If spatny(i) Then zakazany(i) = True
        Next i
    Next opak
End Function

Public Function VPlose(r As Obd) As Boolean
    VPlose = r.L >= gPL + OKRAJ And r.R <= gPR - OKRAJ And r.T >= gPT + Max2(OKRAJ, gRezerva) And r.B <= gPB - OKRAJ
End Function

Private Function Primo(ByVal i As Long, uml() As Long, ByRef l As Double, ByRef t As Double) As Boolean
    Dim x As Double, y As Double, w As Double, h As Double, g As Double, k As Long, j As Long, r As Obd, volno As Boolean
    Dim kl(1 To 8) As Double, kt(1 To 8) As Double
    x = gP(i).X
    y = gP(i).Y
    w = gP(i).W
    h = gP(i).H
    g = VYCHOZI_ODSAZENI
    kl(1) = x + g: kt(1) = y - h / 2  ' order = preference
    kl(2) = x - g - w: kt(2) = y - h / 2
    kl(3) = x - w / 2: kt(3) = y - g - h
    kl(4) = x - w / 2: kt(4) = y + g
    kl(5) = x + 3: kt(5) = y - 3 - h
    kl(6) = x + 3: kt(6) = y + 3
    kl(7) = x - 3 - w: kt(7) = y - 3 - h
    kl(8) = x - 3 - w: kt(8) = y + 3
    For k = 1 To 8
        r = NovyObd(kl(k), kt(k), kl(k) + w, kt(k) + h)
        volno = VPlose(r)
        If volno Then volno = (PresCary(r) = 0)
        If volno Then
            For j = 1 To UBound(gPrek)
                If Prekryv(r, gPrek(j), 2#) Then
                    volno = False
                    Exit For
                End If
            Next j
        End If
        If volno Then
            For j = 1 To UBound(gP)
                If j <> i Then
                    If Prekryv(r, NovyObd(gP(j).X - 3 - 2, gP(j).Y - 3 - 2, gP(j).X + 3 + 2, gP(j).Y + 3 + 2)) Then
                        volno = False
                        Exit For
                    End If
                End If
            Next j
        End If
        If volno Then
            For j = 1 To UBound(uml)
                If Prekryv(r, Obdelnik(uml(j)), 1.5) Then
                    volno = False
                    Exit For
                End If
            Next j
        End If
        If volno Then
            l = kl(k)
            t = kt(k)
            Primo = True
            Exit Function
        End If
    Next k
End Function

Private Function Sloupce(zbytek() As Long, uml() As Long, ByVal mezera As Double) As Boolean
    Dim n As Long, k As Long, m As Long, prek() As Obd, poradi() As Long, klic() As Double
    n = UBound(zbytek)
    If n = 0 Then
        Sloupce = True
        Exit Function
    End If
    prek = gPrek
    For m = 1 To UBound(uml)
        PridejObd prek, Obdelnik(uml(m))
    Next m
    poradi = zbytek
    ReDim klic(0 To n)
    For m = 1 To n
        klic(m) = gP(poradi(m)).X
    Next m
    SeradIndexy poradi, klic

    Dim celkem As Double, okVse As Boolean, cena As Double, ok As Boolean, prvni As Boolean, nejCena As Double, nejOk As Boolean
    Dim skupina() As Long, oL() As Double, oT() As Double
    Dim kandL() As Double, kandT() As Double, kandS() As String, nejL() As Double, nejT() As Double, nejS() As String
    ReDim kandL(0 To n)
    ReDim kandT(0 To n)
    ReDim kandS(0 To n)
    prvni = True
    For k = 0 To n
        celkem = 0#
        okVse = True
        If k > 0 Then
            ReDim skupina(0 To k)
            For m = 1 To k
                skupina(m) = poradi(m)
            Next m
            NejlepsiSloupec "L", skupina, prek, mezera, cena, oL, oT, ok
            celkem = celkem + cena
            okVse = okVse And ok
            For m = 1 To k
                kandL(m) = oL(m)
                kandT(m) = oT(m)
                kandS(m) = "L"
            Next m
        End If
        If k < n Then
            ReDim skupina(0 To n - k)
            For m = 1 To n - k
                skupina(m) = poradi(k + m)
            Next m
            NejlepsiSloupec "R", skupina, prek, mezera, cena, oL, oT, ok
            celkem = celkem + cena
            okVse = okVse And ok
            For m = 1 To n - k
                kandL(k + m) = oL(m)
                kandT(k + m) = oT(m)
                kandS(k + m) = "R"
            Next m
        End If
        If prvni Or celkem < nejCena Then
            prvni = False
            nejCena = celkem
            nejOk = okVse
            nejL = kandL
            nejT = kandT
            nejS = kandS
        End If
    Next k
    For m = 1 To n
        gP(poradi(m)).L = nejL(m)
        gP(poradi(m)).T = nejT(m)
        gP(poradi(m)).Rezim = nejS(m)
    Next m
    Sloupce = nejOk
End Function

Private Sub NejlepsiSloupec(ByVal strana As String, skupina() As Long, prek() As Obd, ByVal mezera As Double, _
                            ByRef nejCena As Double, nejL() As Double, nejT() As Double, ByRef nejOk As Boolean)
    Dim n As Long, m As Long, i As Long, loVse As Double, hiVse As Double, zaklad As Double, maxW As Double
    Dim kand() As Double, hrana As Variant, ok As Boolean, cena As Double, prvni As Boolean
    n = UBound(skupina)
    loVse = gPT + Max2(OKRAJ, gRezerva)
    hiVse = gPB - OKRAJ
    zaklad = gP(skupina(1)).X
    maxW = gP(skupina(1)).W
    For m = 2 To n
        If strana = "L" Then zaklad = Min2(zaklad, gP(skupina(m)).X) Else zaklad = Max2(zaklad, gP(skupina(m)).X)
        maxW = Max2(maxW, gP(skupina(m)).W)
    Next m
    If strana = "L" Then zaklad = zaklad - MEZERA_SLOUPCE Else zaklad = zaklad + MEZERA_SLOUPCE
    ReDim kand(0 To 1)
    kand(1) = zaklad
    For m = 1 To UBound(gCary)  ' alternative: column beyond a vertical line it would otherwise cross
        If strana = "R" And zaklad < gCary(m) And gCary(m) < zaklad + maxW Then
            ReDim Preserve kand(0 To UBound(kand) + 1)
            kand(UBound(kand)) = gCary(m) + 4
        End If
        If strana = "L" And zaklad - maxW < gCary(m) And gCary(m) < zaklad Then
            ReDim Preserve kand(0 To UBound(kand) + 1)
            kand(UBound(kand)) = gCary(m) - 4
        End If
    Next m

    Dim horni() As Long, dolni() As Long, kopie() As Long, prekD() As Obd
    ReDim nejL(0 To n)
    ReDim nejT(0 To n)
    prvni = True
    For m = 1 To UBound(kand)
        hrana = kand(m)
        ok = True
        If gMaDelici Then
            ReDim horni(0 To 0)
            ReDim dolni(0 To 0)
            For i = 1 To n
                If gP(skupina(i)).Y < gDelici Then PridejIndex horni, skupina(i) Else PridejIndex dolni, skupina(i)
            Next i
            If UBound(horni) > 0 Then
                If Not Sloupec(horni, strana, CDbl(hrana), loVse, gDelici - 2, prek, mezera) Then ok = False
            End If
            If UBound(dolni) > 0 Then
                If Not Sloupec(dolni, strana, CDbl(hrana), gDelici + 2, hiVse, prek, mezera) Then ok = False
            End If
        End If
        If Not gMaDelici Or Not ok Then  ' one column over the full height, but not across the dividing line
            prekD = prek
            If gMaDelici Then PridejObd prekD, NovyObd(gPL, gDelici - 1, gPR, gDelici + 1)
            kopie = skupina
            ok = Sloupec(kopie, strana, CDbl(hrana), loVse, hiVse, prekD, mezera)
        End If
        If ok Then
            For i = 1 To n
                If Not (gP(skupina(i)).L >= gPL + OKRAJ And gP(skupina(i)).L + gP(skupina(i)).W <= gPR - OKRAJ) Then
                    ok = False
                    Exit For
                End If
            Next i
        End If
        cena = CenaSkupiny(skupina) + IIf(ok, 0#, NEVEJDE_SE)
        If prvni Or cena < nejCena Then
            prvni = False
            nejCena = cena
            nejOk = ok
            For i = 1 To n
                nejL(i) = gP(skupina(i)).L
                nejT(i) = gP(skupina(i)).T
            Next i
        End If
    Next m
End Sub

Private Function CenaSkupiny(skupina() As Long) As Double
    Dim i As Long, c As Cesta
    For i = 1 To UBound(skupina)
        c = VodiciCara(skupina(i))
        CenaSkupiny = CenaSkupiny + DelkaCesty(c)
        CenaSkupiny = CenaSkupiny + 150 * PresCary(Obdelnik(skupina(i)))
    Next i
End Function

' Column at an edge: order by y, then swap two labels whenever that shortens the total leader line length
' (for straight lines to one vertical line, the assignment with the minimum length has no crossings).
Private Function Sloupec(s() As Long, ByVal strana As String, ByVal hrana As Double, ByVal lo As Double, ByVal hi As Double, _
                         prek() As Obd, ByVal mezera As Double) As Boolean
    Dim n As Long, k As Long, a As Long, b As Long, ia As Long, ib As Long, opak As Long, ax As Double
    Dim klic() As Double, stredy() As Double, zlepseni As Boolean, krizi As Boolean, cesty() As Cesta
    n = UBound(s)
    ReDim klic(0 To n)
    For k = 1 To n
        If strana = "L" Then gP(s(k)).L = hrana - gP(s(k)).W Else gP(s(k)).L = hrana
        klic(k) = gP(s(k)).Y
    Next k
    SeradIndexy s, klic
    For k = 1 To n
        gP(s(k)).Chce = gP(s(k)).Y - gP(s(k)).H / 2
    Next k
    Sloupec = Zabal(s, lo, hi, prek, mezera)
    If strana = "L" Then ax = hrana + ZALOMENI_CARY Else ax = hrana - ZALOMENI_CARY
    ReDim stredy(0 To n)
    ReDim cesty(0 To n)
    For opak = 1 To 6
        For k = 1 To n
            stredy(k) = gP(s(k)).T + gP(s(k)).H / 2
        Next k
        Do
            zlepseni = False
            For a = 1 To n
                For b = a + 1 To n
                    ia = s(a)
                    ib = s(b)
                    If Delka(ia, ax, stredy(b)) + Delka(ib, ax, stredy(a)) < _
                       Delka(ia, ax, stredy(a)) + Delka(ib, ax, stredy(b)) - EPS Then
                        s(a) = ib
                        s(b) = ia
                        zlepseni = True
                    End If
                Next b
            Next a
        Loop While zlepseni
        For k = 1 To n
            gP(s(k)).Chce = stredy(k) - gP(s(k)).H / 2
        Next k
        Sloupec = Zabal(s, lo, hi, prek, mezera)
        For k = 1 To n
            cesty(k) = VodiciCara(s(k))
        Next k
        krizi = False
        For a = 1 To n
            For b = a + 1 To n
                If CestyKrizi(cesty(a), cesty(b)) Then
                    krizi = True
                    Exit For
                End If
            Next b
            If krizi Then Exit For
        Next a
        If Not krizi Then Exit For
    Next opak
End Function

Private Function Delka(ByVal i As Long, ByVal ax As Double, ByVal cy As Double) As Double
    Dim dx As Double, dy As Double
    dx = gP(i).X - ax
    dy = gP(i).Y - cy
    Delka = Sqr(dx * dx + dy * dy)
End Function

' Order preserved; least squares (PAV), then going around obstacles (vertical intervals within the label width).
Private Function Zabal(s() As Long, ByVal lo As Double, ByVal hi As Double, prek() As Obd, ByVal mezera As Double) As Boolean
    Dim n As Long, k As Long, i As Long, nb As Long, s1 As Long, e1 As Long, e2 As Long
    Dim t1 As Double, t2 As Double, acc As Double, soucet As Double, horni As Double, cur As Double
    Dim bs() As Long, be() As Long, bt() As Double
    n = UBound(s)
    If n = 0 Then
        Zabal = True
        Exit Function
    End If
    ReDim bs(0 To n)
    ReDim be(0 To n)
    ReDim bt(0 To n)
    For k = 1 To n
        nb = nb + 1
        bs(nb) = k
        be(nb) = k
        bt(nb) = gP(s(k)).Chce
        Do While nb >= 2
            s1 = bs(nb - 1)
            e1 = be(nb - 1)
            t1 = bt(nb - 1)
            e2 = be(nb)
            t2 = bt(nb)
            soucet = 0#
            For i = s1 To e1
                soucet = soucet + (gP(s(i)).H + mezera)
            Next i
            If t1 + soucet <= t2 Then Exit Do
            acc = 0#
            soucet = 0#
            For i = s1 To e2
                soucet = soucet + (gP(s(i)).Chce - acc)
                acc = acc + (gP(s(i)).H + mezera)
            Next i
            nb = nb - 1
            be(nb) = e2
            bt(nb) = soucet / (e2 - s1 + 1)
        Loop
    Next k
    For k = 1 To nb
        horni = bt(k)
        For i = bs(k) To be(k)
            gP(s(i)).T = horni
            horni = horni + (gP(s(i)).H + mezera)
        Next i
    Next k

    Dim za() As Double, zb() As Double, nz As Long
    cur = lo
    For k = 1 To n
        i = s(k)
        nz = Zakazane(i, prek, za, zb)
        gP(i).T = Dolu(i, Max2(gP(i).T, cur), za, zb, nz)
        cur = gP(i).T + gP(i).H + mezera
    Next k
    If gP(s(n)).T + gP(s(n)).H > hi Then
        cur = hi
        For k = n To 1 Step -1
            i = s(k)
            nz = Zakazane(i, prek, za, zb)
            gP(i).T = Nahoru(i, Min2(gP(i).T + gP(i).H, cur) - gP(i).H, za, zb, nz)
            cur = gP(i).T - mezera
        Next k
    End If
    Zabal = True
    For k = 1 To n
        i = s(k)
        If Not (lo - 0.5 <= gP(i).T And gP(i).T + gP(i).H <= hi + 0.5) Then Zabal = False
    Next k
    If Not Zabal Then Exit Function
    For k = 1 To n
        i = s(k)
        nz = Zakazane(i, prek, za, zb)
        For e1 = 1 To nz
            If gP(i).T < zb(e1) - EPS And za(e1) < gP(i).T + gP(i).H - EPS Then
                Zabal = False
                Exit Function
            End If
        Next e1
    Next k
End Function

' Forbidden vertical intervals of a label (obstacles that overlap its width horizontally), sorted by (from, to).
Private Function Zakazane(ByVal i As Long, prek() As Obd, za() As Double, zb() As Double) As Long
    Dim k As Long, j As Long, n As Long, a As Double, b As Double
    ReDim za(0 To UBound(prek) + 1)
    ReDim zb(0 To UBound(prek) + 1)
    For k = 1 To UBound(prek)
        If gP(i).L < prek(k).R + 3 And prek(k).L - 3 < gP(i).L + gP(i).W Then
            a = prek(k).T - 3
            b = prek(k).B + 3
            j = n
            Do While j >= 1
                If za(j) > a Or (za(j) = a And zb(j) > b) Then
                    za(j + 1) = za(j)
                    zb(j + 1) = zb(j)
                    j = j - 1
                Else
                    Exit Do
                End If
            Loop
            za(j + 1) = a
            zb(j + 1) = b
            n = n + 1
        End If
    Next k
    Zakazane = n
End Function

Private Function Dolu(ByVal i As Long, ByVal t As Double, za() As Double, zb() As Double, ByVal nz As Long) As Double
    Dim posun As Boolean, k As Long
    Do
        posun = False
        For k = 1 To nz
            If t < zb(k) - EPS And za(k) < t + gP(i).H - EPS Then
                t = zb(k)
                posun = True
            End If
        Next k
    Loop While posun
    Dolu = t
End Function

Private Function Nahoru(ByVal i As Long, ByVal t As Double, za() As Double, zb() As Double, ByVal nz As Long) As Double
    Dim posun As Boolean, k As Long
    Do
        posun = False
        For k = 1 To nz
            If t < zb(k) - EPS And za(k) < t + gP(i).H - EPS Then
                t = za(k) - gP(i).H
                posun = True
            End If
        Next k
    Loop While posun
    Nahoru = t
End Function

' Stable sort of indexes by key (klic(k) belongs to s(k)); both arrays are permuted together.
Public Sub SeradIndexy(s() As Long, klic() As Double)
    Dim i As Long, j As Long, vs As Long, ks As Double
    For i = 2 To UBound(s)
        vs = s(i)
        ks = klic(i)
        j = i - 1
        Do While j >= 1
            If klic(j) > ks Then
                s(j + 1) = s(j)
                klic(j + 1) = klic(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        s(j + 1) = vs
        klic(j + 1) = ks
    Next i
End Sub
