Option Explicit

' XY chart model in chart frame coordinates (pt), the same as the Python reference popisky/graf.py.
' Reads and writes through the PowerPoint object model; text is measured by module Pisma.
'   chart frame = chart object model coordinates + ChartArea.Left/Top (chart margin, typically 4 pt)
'   point       = inner plot area (PlotArea.Inside*) + linear axis (MinimumScale/MaximumScale; an automatic
'                 bound returns the value PowerPoint draws with, verified by a probe on 9 Oct 2026)
'   label       = text width + 6 pt, lines * line spacing * size + 3 pt; PowerPoint wraps a label wider
'                 than 20 % of the chart width
' Object model pitfalls (verified by a probe on 8 Oct 2026):
'   - TextFrame2.AutoSize = 0 on a data label crashes PowerPoint; a custom width can be set directly
'     through Width and Height.
'   - Width moves Left, so the position is written last. Writing Left also forces the size to be recalculated
'     (after a font change Width/Height return the old values, Chart.Refresh does not help).
'   - A label PowerPoint does not draw returns Left = NaN; Point.Left/Top return nonsense.

Public Const VYCHOZI_ODSAZENI As Double = 6#
Public Const ZALOMENI As Double = 0.2
Public Const ZALOMENI_CARY As Double = 4#
Public Const VODICI_OD As Double = 7#
Private Const REZERVA_INDEX As Double = 34#   ' "Index NN" labels at the top of the chart
Private Const XL_HODNOTY As Long = 2
Private Const XL_KATEGORIE As Long = 1
Private Const XL_LOG As Long = -4133
Private Const XL_BEZ_ZNACKY As Long = -4142
Private Const XL_BEZ_POPISKU_OSY As Long = -4142  ' xlTickLabelPositionNone
Private Const AUTO_VELIKOST As Long = 1        ' msoAutoSizeShapeToFitText
Private Const TOLERANCE_ZAPISU As Double = 0.5

Public Type Obd
    L As Double
    T As Double
    R As Double
    B As Double
End Type

Public Type Cesta
    N As Long
    X(1 To 3) As Double
    Y(1 To 3) As Double
End Type

Public Type Popisek
    Txt As String
    Fs As Double
    FsPuvodni As Double
    Pismo As Long
    X As Double
    Y As Double
    W As Double
    H As Double
    L As Double
    T As Double
    Rezim As String          ' "" | "D" next to the point | "L" / "R" column | "B" with a leader line (Blizko)
                             ' | "N" not placed by the fallback (original position)
    Chce As Double
    VlastniSirka As Boolean
    Serie As Long
    Bod As Long
End Type

Public Type Vady
    PrekryvPopisku As Long
    PopisekPresBod As Long
    KrizeniCar As Long
    CaraPresPopisek As Long
    PopisekPresCaru As Long
    PopisekPresBox As Long
    MimoGraf As Long
End Type

Private Type TvarSlidu
    X As Double
    Y As Double
    W As Double
    H As Double
    Cara As Boolean
    Otoceny As Boolean
End Type

' state of one chart; arrays start at 1, element 0 is unused (empty list = UBound 0)
Public gW As Double, gH As Double
Public gPL As Double, gPT As Double, gPR As Double, gPB As Double
Public gCaL As Double, gCaT As Double
Public gP() As Popisek
Public gPrek() As Obd
Public gCary() As Double
Public gMaDelici As Boolean
Public gDelici As Double
Public gRezerva As Double
' for the pre-adjustment (module Predupravy): the axis ranges the chart was read with, whether a range may be
' widened (no tick labels) and how many of the first gCary come from slide shapes (the rest move with the data)
Public gOsaX0 As Double, gOsaX1 As Double, gOsaY0 As Double, gOsaY1 As Double
Public gOsaXVolna As Boolean, gOsaYVolna As Boolean
Public gCarTvaru As Long
' set by Predupravy.VratOsy before NactiGraf: the axes were just returned to their original ranges and the labels
' still report positions drawn with the ranges gStX0..gStY1 (the object model does not recalculate them)
Public gStaraOsa As Boolean
Public gStX0 As Double, gStX1 As Double, gStY0 As Double, gStY1 As Double

Private mTvary() As TvarSlidu

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)

' ---------- helpers ----------
Public Function Max2(ByVal a As Double, ByVal b As Double) As Double
    If b > a Then Max2 = b Else Max2 = a
End Function

Public Function Min2(ByVal a As Double, ByVal b As Double) As Double
    If b < a Then Min2 = b Else Min2 = a
End Function

Public Function NovyObd(ByVal l As Double, ByVal t As Double, ByVal r As Double, ByVal b As Double) As Obd
    NovyObd.L = l
    NovyObd.T = t
    NovyObd.R = r
    NovyObd.B = b
End Function

Public Sub PridejObd(seznam() As Obd, o As Obd)
    ReDim Preserve seznam(0 To UBound(seznam) + 1)
    seznam(UBound(seznam)) = o
End Sub

Public Sub PridejIndex(seznam() As Long, ByVal i As Long)
    ReDim Preserve seznam(0 To UBound(seznam) + 1)
    seznam(UBound(seznam)) = i
End Sub

Private Sub PridejCislo(seznam() As Double, ByVal v As Double)
    ReDim Preserve seznam(0 To UBound(seznam) + 1)
    seznam(UBound(seznam)) = v
End Sub

Public Function JeNaN(ByVal d As Double) As Boolean
    ' NaN and infinity: all exponent bits set (neither comparisons nor CStr are reliable with NaN in VBA)
    Dim b(0 To 7) As Byte
    CopyMemory b(0), d, 8
    JeNaN = ((b(7) And &H7F) = &H7F) And ((b(6) And &HF0) = &HF0)
End Function

' Series values as an array indexed 1..n (COM may return a different lower bound).
Private Function Hodnoty(v As Variant) As Variant
    Dim out() As Variant, i As Long, n As Long
    If Not IsArray(v) Then
        ReDim out(0 To 0)
    Else
        n = UBound(v) - LBound(v) + 1
        ReDim out(0 To n)
        For i = 1 To n
            out(i) = v(LBound(v) + i - 1)
        Next i
    End If
    Hodnoty = out
End Function

Private Function JeCislo(v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            JeCislo = True
    End Select
End Function

' ---------- geometry ----------
Public Function Obdelnik(ByVal i As Long) As Obd
    Obdelnik.L = gP(i).L
    Obdelnik.T = gP(i).T
    Obdelnik.R = gP(i).L + gP(i).W
    Obdelnik.B = gP(i).T + gP(i).H
End Function

Public Function Prekryv(a As Obd, b As Obd, Optional ByVal pad As Double = 0#) As Boolean
    Prekryv = a.L < b.R + pad And b.L - pad < a.R And a.T < b.B + pad And b.T - pad < a.B
End Function

Private Function Orientace(ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double, _
                           ByVal cx As Double, ByVal cy As Double) As Double
    Orientace = (bx - ax) * (cy - ay) - (by - ay) * (cx - ax)
End Function

Public Function Krizeni(ByVal p1x As Double, ByVal p1y As Double, ByVal p2x As Double, ByVal p2y As Double, _
                        ByVal q1x As Double, ByVal q1y As Double, ByVal q2x As Double, ByVal q2y As Double) As Boolean
    If Orientace(p1x, p1y, p2x, p2y, q1x, q1y) * Orientace(p1x, p1y, p2x, p2y, q2x, q2y) < 0 Then
        Krizeni = Orientace(q1x, q1y, q2x, q2y, p1x, p1y) * Orientace(q1x, q1y, q2x, q2y, p2x, p2y) < 0
    End If
End Function

Public Function UseckaPres(ByVal px As Double, ByVal py As Double, ByVal qx As Double, ByVal qy As Double, _
                           r As Obd, Optional ByVal zmenseni As Double = 1#) As Boolean
    Dim l As Double, t As Double, rr As Double, b As Double
    l = r.L + zmenseni
    t = r.T + zmenseni
    rr = r.R - zmenseni
    b = r.B - zmenseni
    If l >= rr Or t >= b Then Exit Function
    If (l <= px And px <= rr And t <= py And py <= b) Or (l <= qx And qx <= rr And t <= qy And qy <= b) Then
        UseckaPres = True
        Exit Function
    End If
    UseckaPres = Krizeni(px, py, qx, qy, l, t, rr, t) Or Krizeni(px, py, qx, qy, rr, t, rr, b) _
        Or Krizeni(px, py, qx, qy, rr, b, l, b) Or Krizeni(px, py, qx, qy, l, b, l, t)
End Function

' Leader line as PowerPoint draws it: N = 0 (none), or a polyline [point, elbow, label edge].
Public Function VodiciCara(ByVal i As Long) As Cesta
    Dim l As Double, t As Double, r As Double, b As Double, x As Double, y As Double
    Dim dx As Double, dy As Double, sy As Double
    l = gP(i).L
    t = gP(i).T
    r = gP(i).L + gP(i).W
    b = gP(i).T + gP(i).H
    x = gP(i).X
    y = gP(i).Y
    dx = Max2(Max2(l - x, 0#), x - r)
    dy = Max2(Max2(t - y, 0#), y - b)
    If Sqr(dx * dx + dy * dy) <= VODICI_OD Then Exit Function
    VodiciCara.X(1) = x
    VodiciCara.Y(1) = y
    If l <= x And x <= r Then  ' point below or above the label: middle of the top or bottom edge
        VodiciCara.N = 2
        VodiciCara.X(2) = (l + r) / 2
        If y > b Then VodiciCara.Y(2) = b Else VodiciCara.Y(2) = t
        Exit Function
    End If
    sy = (t + b) / 2
    VodiciCara.N = 3
    If x < l Then
        VodiciCara.X(2) = l - ZALOMENI_CARY
        VodiciCara.X(3) = l
    Else
        VodiciCara.X(2) = r + ZALOMENI_CARY
        VodiciCara.X(3) = r
    End If
    VodiciCara.Y(2) = sy
    VodiciCara.Y(3) = sy
End Function

Public Function CestyKrizi(a As Cesta, b As Cesta) As Boolean
    Dim i As Long, j As Long
    For i = 1 To a.N - 1
        For j = 1 To b.N - 1
            If Krizeni(a.X(i), a.Y(i), a.X(i + 1), a.Y(i + 1), b.X(j), b.Y(j), b.X(j + 1), b.Y(j + 1)) Then
                CestyKrizi = True
                Exit Function
            End If
        Next j
    Next i
End Function

Public Function CestaPres(c As Cesta, r As Obd, ByVal zmenseni As Double) As Boolean
    Dim i As Long
    For i = 1 To c.N - 1
        If UseckaPres(c.X(i), c.Y(i), c.X(i + 1), c.Y(i + 1), r, zmenseni) Then
            CestaPres = True
            Exit Function
        End If
    Next i
End Function

Public Function DelkaCesty(c As Cesta) As Double
    Dim i As Long, dx As Double, dy As Double
    For i = 1 To c.N - 1
        dx = c.X(i + 1) - c.X(i)
        dy = c.Y(i + 1) - c.Y(i)
        DelkaCesty = DelkaCesty + Sqr(dx * dx + dy * dy)
    Next i
End Function

Public Function PresCary(r As Obd) As Long
    Dim k As Long
    For k = 1 To UBound(gCary)
        If r.L < gCary(k) And gCary(k) < r.R Then PresCary = PresCary + 1
    Next k
    If gMaDelici Then
        If r.T < gDelici And gDelici < r.B Then PresCary = PresCary + 1
    End If
End Function

Public Sub Jednoradkovy(ByVal i As Long, ByRef w As Double, ByRef h As Double)
    w = SirkaTextu(gP(i).Pismo, gP(i).Txt, gP(i).Fs) + 6#
    h = RadekPisma(gP(i).Pismo) * gP(i).Fs + 3#
End Sub

' ---------- reading from the object model ----------
Public Function MaXYSerii(ch As Object) As Boolean
    Dim i As Long
    For i = 1 To ch.SeriesCollection.Count
        If JeXY(ch.SeriesCollection(i).ChartType) Then
            MaXYSerii = True
            Exit Function
        End If
    Next i
End Function

Private Function JeXY(ByVal typ As Long) As Boolean
    Select Case typ
        Case -4169, 72, 73, 74, 75  ' xlXYScatter, ...Smooth, ...SmoothNoMarkers, ...Lines, ...LinesNoMarkers
            JeXY = True
    End Select
End Function

' Slide shapes (slide coordinates); charts are not obstacles, a group is an obstacle with its whole rectangle.
Public Sub NactiTvary(sl As Slide)
    Dim s As Shape, n As Long, rot As Double
    ReDim mTvary(0 To sl.Shapes.Count)
    For Each s In sl.Shapes
        If Not s.HasChart Then
            n = n + 1
            mTvary(n).X = s.Left
            mTvary(n).Y = s.Top
            mTvary(n).W = s.Width
            mTvary(n).H = s.Height
            mTvary(n).Cara = (s.Connector = msoTrue)
            rot = s.Rotation
            mTvary(n).Otoceny = Abs(rot - 180# * Int(rot / 180# + 0.5)) > 0.001
        End If
    Next s
    ReDim Preserve mTvary(0 To n)
End Sub

' Loads the chart into the module state; returns "" or the reason why the chart is skipped.
Public Function NactiGraf(sl As Slide, shp As Shape) As String
    Dim ch As Object, pa As Object, osaX As Object, osaY As Object
    Dim x0 As Double, x1 As Double, y0 As Double, y1 As Double
    Set ch = shp.Chart
    gW = shp.Width
    gH = shp.Height
    gCaL = ch.ChartArea.Left
    gCaT = ch.ChartArea.Top
    ReDim gP(0 To 0)
    ReDim gPrek(0 To 0)
    ReDim gCary(0 To 0)
    gMaDelici = False
    gRezerva = 0#

    If Not (ch.HasAxis(XL_KATEGORIE, 1) And ch.HasAxis(XL_HODNOTY, 1)) Then
        NactiGraf = "horizontal or vertical axis is missing"
        Exit Function
    End If
    Set osaX = ch.Axes(XL_KATEGORIE, 1)
    Set osaY = ch.Axes(XL_HODNOTY, 1)
    If osaX.ReversePlotOrder Or osaY.ReversePlotOrder Or osaX.ScaleType = XL_LOG Or osaY.ScaleType = XL_LOG Then
        NactiGraf = "reversed or logarithmic axis"
        Exit Function
    End If
    x0 = osaX.MinimumScale  ' also for an automatic bound; the axis stays automatic
    x1 = osaX.MaximumScale
    y0 = osaY.MinimumScale
    y1 = osaY.MaximumScale
    gOsaX0 = x0
    gOsaX1 = x1
    gOsaY0 = y0
    gOsaY1 = y1
    gOsaXVolna = (osaX.TickLabelPosition = XL_BEZ_POPISKU_OSY)
    gOsaYVolna = (osaY.TickLabelPosition = XL_BEZ_POPISKU_OSY)
    Set pa = ch.PlotArea
    gPL = pa.InsideLeft + gCaL
    gPT = pa.InsideTop + gCaT
    gPR = gPL + pa.InsideWidth
    gPB = gPT + pa.InsideHeight

    Dim k As Long, l As Double, t As Double, r As Double, b As Double
    For k = 1 To UBound(mTvary)  ' to chart coordinates
        l = mTvary(k).X - shp.Left
        t = mTvary(k).Y - shp.Top
        r = l + mTvary(k).W
        b = t + mTvary(k).H
        If mTvary(k).Cara Then
            If mTvary(k).H < 1 And mTvary(k).W > 0.5 * (gPR - gPL) Then
                gMaDelici = True
                gDelici = t
            ElseIf mTvary(k).W < 1 And mTvary(k).H > 0.5 * (gPB - gPT) Then
                PridejCislo gCary, l
            End If
        ElseIf Not mTvary(k).Otoceny Then
            If Prekryv(NovyObd(l, t, r, b), NovyObd(gPL, gPT, gPR, gPB)) And _
               Not (l <= gPL And r >= gPR And t <= gPT And b >= gPB) Then
                PridejObd gPrek, NovyObd(l, t, r, b)
            End If
        End If
    Next k
    gCarTvaru = UBound(gCary)  ' vertical lines from the slide come first, lines of helper series follow

    Dim i As Long, j As Long, s As Object, xs As Variant, ys As Variant, dl As Object, pocet As Long
    Dim fs As Double, nazev As String, tucne As Boolean, n As Long, motiv As String
    motiv = PismoMotivu(sl, False)
    For i = 1 To ch.SeriesCollection.Count
        Set s = ch.SeriesCollection(i)
        If JeXY(s.ChartType) Then
            xs = Hodnoty(s.XValues)
            ys = Hodnoty(s.Values)
            pocet = s.Points.Count
            If s.MarkerStyle = XL_BEZ_ZNACKY Then
                PomocnaSerie s, xs, pocet, x0, x1
            Else
                For j = 1 To pocet
                    If s.Points(j).HasDataLabel Then
                        Set dl = s.Points(j).DataLabel
                        If PopisekZBunek(dl, xs, ys, j) Then
                            VlastnostiPisma dl, sl, motiv, fs, nazev, tucne
                            n = UBound(gP) + 1
                            ReDim Preserve gP(0 To n)
                            With gP(n)
                                .Txt = dl.Text
                                .Fs = fs
                                .FsPuvodni = fs
                                .Pismo = PismoIndex(nazev, tucne)
                                .X = gPL + (gPR - gPL) * (CDbl(xs(j)) - x0) / (x1 - x0)
                                .Y = gPT + (gPB - gPT) * (1 - (CDbl(ys(j)) - y0) / (y1 - y0))
                                .L = dl.Left + gCaL
                                .T = dl.Top + gCaT
                                If gStaraOsa Then  ' a label moves with its point: add the shift of the point
                                    .L = .L + (.X - (gPL + (gPR - gPL) * (CDbl(xs(j)) - gStX0) / (gStX1 - gStX0)))
                                    .T = .T + (.Y - (gPT + (gPB - gPT) * (1 - (CDbl(ys(j)) - gStY0) / (gStY1 - gStY0))))
                                End If
                                .W = dl.Width
                                .H = dl.Height
                                .Serie = i
                                .Bod = j
                            End With
                        End If
                    End If
                Next j
            End If
        End If
    Next i
End Function

' A series without markers and with all x equal = vertical line (Index 80/100/120); with labels it
' reserves space at the top of the chart.
Private Sub PomocnaSerie(s As Object, xs As Variant, ByVal pocet As Long, ByVal x0 As Double, ByVal x1 As Double)
    Dim j As Long, prvni As Boolean, gx As Double, stejne As Boolean, popisky As Boolean
    stejne = True
    For j = 1 To pocet
        If j > UBound(xs) Then Exit For
        If JeCislo(xs(j)) Then
            If Not prvni Then
                prvni = True
                gx = xs(j)
            ElseIf xs(j) <> gx Then
                stejne = False
            End If
        End If
        If s.Points(j).HasDataLabel Then popisky = True
    Next j
    If prvni And stejne Then
        PridejCislo gCary, gPL + (gPR - gPL) * (gx - x0) / (x1 - x0)
        If popisky Then gRezerva = REZERVA_INDEX
    End If
End Sub

Private Function PopisekZBunek(dl As Object, xs As Variant, ys As Variant, ByVal j As Long) As Boolean
    Dim zRozsahu As Boolean
    If j > UBound(xs) Or j > UBound(ys) Then Exit Function
    If Not (JeCislo(xs(j)) And JeCislo(ys(j))) Then Exit Function
    If Len(dl.Text) = 0 Then Exit Function
    On Error Resume Next
    zRozsahu = True
    zRozsahu = dl.ShowRange  ' Office 2013+; without the property every label counts
    On Error GoTo 0
    If Not zRozsahu Then Exit Function
    PopisekZBunek = Not JeNaN(dl.Left)  ' NaN = PowerPoint does not draw the label
End Function

Private Sub VlastnostiPisma(dl As Object, sl As Slide, ByVal motiv As String, ByRef fs As Double, _
                            ByRef nazev As String, ByRef tucne As Boolean)
    Dim f As Object, b As Long
    On Error Resume Next
    Set f = dl.Format.TextFrame2.TextRange.Font
    fs = f.Size
    nazev = f.Name
    b = f.Bold
    If fs < 1 Or fs > 4000 Or nazev = "" Or b = -2 Then  ' mixed formatting: first run of the text
        Set f = dl.Format.TextFrame2.TextRange.Runs(1).Font
        If fs < 1 Or fs > 4000 Then fs = f.Size
        If nazev = "" Then nazev = f.Name
        If b = -2 Then b = f.Bold
    End If
    On Error GoTo 0
    If fs < 1 Or fs > 4000 Then fs = 10#
    tucne = (b = msoTrue)
    If nazev = "" Or nazev = "+mn-lt" Then
        nazev = motiv
    ElseIf nazev = "+mj-lt" Then
        nazev = PismoMotivu(sl, True)
    End If
End Sub

Private Function PismoMotivu(sl As Slide, ByVal nadpisove As Boolean) As String
    On Error Resume Next
    If nadpisove Then
        PismoMotivu = sl.Design.SlideMaster.Theme.ThemeFontScheme.MajorFont(1).Name  ' msoThemeLatin
    Else
        PismoMotivu = sl.Design.SlideMaster.Theme.ThemeFontScheme.MinorFont(1).Name
    End If
    On Error GoTo 0
    If PismoMotivu = "" Then PismoMotivu = IIf(nadpisove, "Calibri Light", "Calibri")
End Function

' ---------- writing ----------
' Writes the placement to the chart; returns the number of labels that differ from the model after
' writing (by more than 0.5 pt).
Public Function ZapisGraf(shp As Shape) As Long
    Dim ch As Object, dl As Object, tf As Object, i As Long
    Set ch = shp.Chart
    For i = 1 To UBound(gP)
        Set dl = ch.SeriesCollection(gP(i).Serie).Points(gP(i).Bod).DataLabel
        Set tf = dl.Format.TextFrame2
        If gP(i).Fs <> gP(i).FsPuvodni Then tf.TextRange.Font.Size = gP(i).Fs
        If gP(i).Rezim <> "N" Then  ' "N" = not placed by the fallback: original position and size, only the font
            If gP(i).VlastniSirka Then  ' without a custom width PowerPoint would wrap a long label
                dl.Width = gP(i).W
                dl.Height = gP(i).H
            ElseIf tf.AutoSize <> AUTO_VELIKOST Then  ' manual size from an earlier run: back to automatic
                tf.AutoSize = AUTO_VELIKOST
            End If
            dl.Left = gP(i).L - gCaL  ' last: Width moves Left; writing Left forces the size to be recalculated
            dl.Top = gP(i).T - gCaT
        End If
    Next i
    For i = 1 To UBound(gP)
        Set dl = ch.SeriesCollection(gP(i).Serie).Points(gP(i).Bod).DataLabel
        If gP(i).Rezim = "N" Then
            ' left where it was
        ElseIf JeNaN(dl.Left) Or JeNaN(dl.Width) Then
            ZapisGraf = ZapisGraf + 1
        ElseIf Abs(dl.Left + gCaL - gP(i).L) > TOLERANCE_ZAPISU Or Abs(dl.Top + gCaT - gP(i).T) > TOLERANCE_ZAPISU Or _
           Abs(dl.Width - gP(i).W) > TOLERANCE_ZAPISU Or Abs(dl.Height - gP(i).H) > TOLERANCE_ZAPISU Then
            ZapisGraf = ZapisGraf + 1
        End If
    Next i
End Function

' ---------- font size and defects ----------
' Gives all labels the most common font size (on a tie the one found first); returns the size,
' zmeneno = number of labels changed.
Public Function SjednotPismo(ByRef zmeneno As Long) As Double
    Dim n As Long, i As Long, j As Long, velikosti() As Double, pocty() As Long, m As Long, nej As Long
    n = UBound(gP)
    zmeneno = 0
    If n = 0 Then Exit Function
    ReDim velikosti(1 To n)
    ReDim pocty(1 To n)
    For i = 1 To n
        For j = 1 To m
            If velikosti(j) = gP(i).Fs Then Exit For
        Next j
        If j > m Then
            m = m + 1
            velikosti(m) = gP(i).Fs
        End If
        pocty(j) = pocty(j) + 1
    Next i
    nej = 1
    For j = 2 To m
        If pocty(j) > pocty(nej) Then nej = j
    Next j
    For i = 1 To n
        If gP(i).Fs <> velikosti(nej) Then
            gP(i).Fs = velikosti(nej)
            zmeneno = zmeneno + 1
        End If
    Next i
    SjednotPismo = velikosti(nej)
End Function

Public Function SpocitejVady() As Vady
    Dim n As Long, i As Long, j As Long, k As Long, ra As Obd, v As Vady
    n = UBound(gP)
    If n = 0 Then Exit Function
    Dim cesty() As Cesta
    ReDim cesty(1 To n)
    For i = 1 To n
        cesty(i) = VodiciCara(i)
    Next i
    For i = 1 To n
        ra = Obdelnik(i)
        If Not (ra.L >= gPL - 0.5 And ra.R <= gPR + 0.5 And ra.T >= gPT - 0.5 And ra.B <= gPB + 0.5) Then
            v.MimoGraf = v.MimoGraf + 1
        End If
        For k = 1 To UBound(gPrek)
            If Prekryv(ra, gPrek(k), -0.5) Then v.PopisekPresBox = v.PopisekPresBox + 1
        Next k
        v.PopisekPresCaru = v.PopisekPresCaru + PresCary(NovyObd(ra.L + 1, ra.T + 1, ra.R - 1, ra.B - 1))
        For j = 1 To n
            If j <> i Then
                If j > i Then
                    If Prekryv(ra, Obdelnik(j), -0.5) Then v.PrekryvPopisku = v.PrekryvPopisku + 1
                End If
                If Prekryv(ra, NovyObd(gP(j).X - 3, gP(j).Y - 3, gP(j).X + 3, gP(j).Y + 3), -0.5) Then
                    v.PopisekPresBod = v.PopisekPresBod + 1
                End If
                If CestaPres(cesty(j), ra, 1#) Then v.CaraPresPopisek = v.CaraPresPopisek + 1
                If j > i Then
                    If CestyKrizi(cesty(i), cesty(j)) Then v.KrizeniCar = v.KrizeniCar + 1
                End If
            End If
        Next j
    Next i
    SpocitejVady = v
End Function

Public Function SoucetVad(v As Vady) As Long
    SoucetVad = v.PrekryvPopisku + v.PopisekPresBod + v.KrizeniCar + v.CaraPresPopisek + v.PopisekPresCaru + _
                v.PopisekPresBox + v.MimoGraf
End Function
