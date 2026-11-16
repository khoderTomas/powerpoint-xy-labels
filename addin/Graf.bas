Option Explicit

' XY chart model in chart frame coordinates (pt), the same as the Python reference popisky/graf.py.
' Reads and writes through the PowerPoint object model; text is measured by module Pisma.
'   chart frame = chart object model coordinates + ChartArea.Left/Top (chart margin, typically 4 pt)
'   point       = inner plot area (PlotArea.Inside*) + linear axis (MinimumScale/MaximumScale; an automatic
'                 bound returns the value PowerPoint draws with, verified by a probe on 9 Oct 2026); a date X axis
'                 (combination chart with areas) counts in days with the point in the middle of its day when the
'                 axis is between categories (probe of 15 Nov 2026)
'   radius      = half the marker size, or the bubble radius (Point.Width); as an obstacle a point is a square
'                 with half side max(3, radius) (a 6 pt marker = the former 3 pt); a bubble is an obstacle for its
'                 own label and for the columns too, never for leader lines (PowerPoint draws a bubble's leader
'                 line from its centre even when the label touches its edge, probe of 15 Nov 2026)
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
Public Const BOD_MIN As Double = 3#             ' smallest half side of a point as an obstacle (6 pt marker)
Private Const UVNITR As Double = 1#             ' a slide line or an axis inside the plot area: this far from its edge
Private Const XL_CASOVA As Long = 3             ' xlTimeScale
Private Const XL_DNY As Long = 0                ' xlDays
Private Const XL_KRIZENI_VLASTNI As Long = -4114  ' xlAxisCrossesCustom
Private Const XL_KRIZENI_MAX As Long = 2        ' xlAxisCrossesMaximum
Private Const XL_KRIZENI_MIN As Long = 4        ' xlAxisCrossesMinimum
Private Const XL_PLOCHA As Long = 1             ' xlArea
Private Const XL_CARA As Long = 4               ' xlLine
Private Const XL_CARA_ZNACKY As Long = 65       ' xlLineMarkers
Private Const XL_POLOHA_VLASTNI As Long = 7     ' xlLabelPositionCustom
Private Const XL_POLOHA_VPRAVO As Long = -4152  ' xlLabelPositionRight

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
    R As Double              ' radius of the point: half the marker size or the bubble radius
    Bublina As Boolean
End Type

' a bubble without a label of its own (an obstacle for labels)
Public Type BodBubliny
    X As Double
    Y As Double
    R As Double
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

' The state of one chart put aside (UlozStav, ObnovStav): module Popisky reads every chart before it places the
' first one, because the font size is common to all of them.
Public Type StavGrafu
    W As Double
    H As Double
    PL As Double
    PT As Double
    PR As Double
    PB As Double
    CaL As Double
    CaT As Double
    P() As Popisek
    Prek() As Obd
    Cary() As Double
    MaDelici As Boolean
    Delici As Double
    Rezerva As Double
    OsaX0 As Double
    OsaX1 As Double
    OsaY0 As Double
    OsaY1 As Double
    OsaXVolna As Boolean
    OsaYVolna As Boolean
    CarTvaru As Long
    Bub() As BodBubliny
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
Public gBub() As BodBubliny
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

' Half side of point i as an obstacle: max(3, radius).
Public Function Polomer(ByVal i As Long) As Double
    Polomer = Max2(BOD_MIN, gP(i).R)
End Function

' How much point i is larger than a 6 pt marker: positions next to the point move away from it by that much.
Public Function Presah(ByVal i As Long) As Double
    Presah = Polomer(i) - BOD_MIN
End Function

' Point i as an obstacle (square), pad = extra margin.
Public Function BodObd(ByVal i As Long, ByVal pad As Double) As Obd
    Dim p As Double
    p = Polomer(i) + pad
    BodObd = NovyObd(gP(i).X - p, gP(i).Y - p, gP(i).X + p, gP(i).Y + p)
End Function

' Bubble k without a label as an obstacle (square).
Public Function BublinaObd(ByVal k As Long, ByVal pad As Double) As Obd
    Dim p As Double
    p = Max2(BOD_MIN, gBub(k).R) + pad
    BublinaObd = NovyObd(gBub(k).X - p, gBub(k).Y - p, gBub(k).X + p, gBub(k).Y + p)
End Function

' Bubbles as obstacles of the labels in columns (labelled bubbles and bubbles without a label) appended to prek.
Public Sub PridejBubliny(prek() As Obd)
    Dim i As Long
    For i = 1 To UBound(gP)
        If gP(i).Bublina Then PridejObd prek, BodObd(i, 0#)
    Next i
    For i = 1 To UBound(gBub)
        PridejObd prek, BublinaObd(i, 0#)
    Next i
End Sub

' Rectangle r is clear of the points of the other labels, the bubbles without a label and its own bubble.
Public Function VolnoUBodu(ByVal i As Long, r As Obd) As Boolean
    Dim j As Long
    For j = 1 To UBound(gP)
        If j <> i Then
            If Prekryv(r, BodObd(j, 2#)) Then Exit Function
        End If
    Next j
    For j = 1 To UBound(gBub)
        If Prekryv(r, BublinaObd(j, 2#)) Then Exit Function
    Next j
    If gP(i).Bublina Then
        If Prekryv(r, BodObd(i, 0#)) Then Exit Function
    End If
    VolnoUBodu = True
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
    Dim i As Long, typ As Long
    For i = 1 To ch.SeriesCollection.Count
        typ = ch.SeriesCollection(i).ChartType
        If JeXY(typ) Or JeBublina(typ) Then
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

Private Function JeBublina(ByVal typ As Long) As Boolean
    JeBublina = (typ = 15 Or typ = 87)  ' xlBubble, xlBubble3DEffect
End Function

' CategoryType of a category axis, -1 for a value axis (the property fails there).
Private Function TypKategorie(osa As Object) As Long
    TypKategorie = -1
    On Error Resume Next
    TypKategorie = osa.CategoryType
End Function

' The axis line is drawn.
Private Function CaraOsy(osa As Object) As Boolean
    On Error Resume Next
    CaraOsy = (osa.Format.Line.Visible = msoTrue)
End Function

' The value on axis osa where the perpendicular axis crosses it (object model: Crosses / CrossesAt of osa).
Private Function KrizeniOsy(osa As Object, ByVal mn As Double, ByVal mx As Double) As Double
    Select Case osa.Crosses
        Case XL_KRIZENI_VLASTNI
            KrizeniOsy = osa.CrossesAt
        Case XL_KRIZENI_MAX
            KrizeniOsy = mx
        Case XL_KRIZENI_MIN
            KrizeniOsy = mn
        Case Else  ' automatic: at zero, or at the nearer end
            KrizeniOsy = Min2(Max2(0#, mn), mx)
    End Select
End Function

Private Function PrumerBubliny(p As Object) As Double
    Dim d As Double
    On Error Resume Next  ' Point.Width: Office 2013+
    d = p.Width
    On Error GoTo 0
    If Not JeNaN(d) And d > 0 Then PrumerBubliny = d
End Function

Private Function VelikostZnacky(s As Object, ByVal j As Long) As Double
    On Error Resume Next
    VelikostZnacky = s.Points(j).MarkerSize
    If VelikostZnacky <= 0 Then VelikostZnacky = s.MarkerSize
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
    Dim ch As Object, pa As Object, osaX As Object, osaY As Object, kategorie As Long
    Dim x0 As Double, x1 As Double, y0 As Double, y1 As Double, v As Double
    Set ch = shp.Chart
    gW = shp.Width
    gH = shp.Height
    gCaL = ch.ChartArea.Left
    gCaT = ch.ChartArea.Top
    ReDim gP(0 To 0)
    ReDim gPrek(0 To 0)
    ReDim gCary(0 To 0)
    ReDim gBub(0 To 0)
    gMaDelici = False
    gRezerva = 0#

    If Not (ch.HasAxis(XL_KATEGORIE, 1) And ch.HasAxis(XL_HODNOTY, 1)) Then
        NactiGraf = "horizontal or vertical axis is missing"
        Exit Function
    End If
    Set osaX = ch.Axes(XL_KATEGORIE, 1)
    Set osaY = ch.Axes(XL_HODNOTY, 1)
    kategorie = TypKategorie(osaX)  ' -1 = value axis (XY chart), otherwise a combination chart
    If kategorie <> -1 Then
        If kategorie <> XL_CASOVA Then
            NactiGraf = "category axis without dates"
            Exit Function
        End If
        If osaX.BaseUnit <> XL_DNY Then
            NactiGraf = "date axis in months or years"
            Exit Function
        End If
        If osaX.ReversePlotOrder Or osaY.ReversePlotOrder Or osaY.ScaleType = XL_LOG Then
            NactiGraf = "reversed or logarithmic axis"
            Exit Function
        End If
    ElseIf osaX.ReversePlotOrder Or osaY.ReversePlotOrder Or osaX.ScaleType = XL_LOG Or osaY.ScaleType = XL_LOG Then
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
    ' a date axis is never changed (the areas of the combination chart would move with it)
    gOsaXVolna = (kategorie = -1) And (osaX.TickLabelPosition = XL_BEZ_POPISKU_OSY)
    gOsaYVolna = (osaY.TickLabelPosition = XL_BEZ_POPISKU_OSY)
    Set pa = ch.PlotArea
    gPL = pa.InsideLeft + gCaL
    gPT = pa.InsideTop + gCaT
    gPR = gPL + pa.InsideWidth
    gPB = gPT + pa.InsideHeight
    Dim mx0 As Double, mx1 As Double  ' the X range the points are mapped with
    mx0 = x0
    mx1 = x1
    If kategorie = XL_CASOVA Then
        If osaX.AxisBetweenCategories Then  ' a day is a cell, the point stands in its middle
            mx0 = x0 - 0.5
            mx1 = x1 + 0.5
        End If
    End If

    Dim k As Long, l As Double, t As Double, r As Double, b As Double
    For k = 1 To UBound(mTvary)  ' to chart coordinates; lines only inside the plot area (axis arrows beside it divide nothing)
        l = mTvary(k).X - shp.Left
        t = mTvary(k).Y - shp.Top
        r = l + mTvary(k).W
        b = t + mTvary(k).H
        If mTvary(k).Cara Then
            If mTvary(k).H < 1 And mTvary(k).W > 0.5 * (gPR - gPL) Then
                If gPT + UVNITR < t And t < gPB - UVNITR Then
                    gMaDelici = True
                    gDelici = t
                End If
            ElseIf mTvary(k).W < 1 And mTvary(k).H > 0.5 * (gPB - gPT) Then
                If gPL + UVNITR < l And l < gPR - UVNITR Then PridejCislo gCary, l
            End If
        ElseIf Not mTvary(k).Otoceny Then
            If Prekryv(NovyObd(l, t, r, b), NovyObd(gPL, gPT, gPR, gPB)) And _
               Not (l <= gPL And r >= gPR And t <= gPT And b >= gPB) Then
                PridejObd gPrek, NovyObd(l, t, r, b)
            End If
        End If
    Next k
    gCarTvaru = UBound(gCary)  ' vertical lines from the slide come first, lines from the data (axis, series) follow

    ' an axis with a visible line inside the plot area (quadrants): horizontal = dividing line, vertical = like Index
    If Not gMaDelici And CaraOsy(osaX) Then
        v = gPT + (gPB - gPT) * (1 - (KrizeniOsy(osaY, y0, y1) - y0) / (y1 - y0))
        If gPT + UVNITR < v And v < gPB - UVNITR Then
            gMaDelici = True
            gDelici = v
        End If
    End If
    If kategorie = -1 And CaraOsy(osaY) Then
        v = gPL + (gPR - gPL) * (KrizeniOsy(osaX, x0, x1) - mx0) / (mx1 - mx0)
        If gPL + UVNITR < v And v < gPR - UVNITR Then PridejCislo gCary, v
    End If

    Dim i As Long, j As Long, s As Object, xs As Variant, ys As Variant, dl As Object, pocet As Long, typ As Long
    Dim fs As Double, nazev As String, tucne As Boolean, n As Long, motiv As String, bub As Boolean, polo As Double
    Dim pridano As Boolean, kandidati() As BodBubliny, nk As Long
    ReDim kandidati(0 To 0)
    motiv = PismoMotivu(sl, False)
    For i = 1 To ch.SeriesCollection.Count
        Set s = ch.SeriesCollection(i)
        typ = s.ChartType
        If typ = XL_PLOCHA Or typ = XL_CARA Or typ = XL_CARA_ZNACKY Then
            If Not gMaDelici Then KonstantniSerie s, y0, y1
        ElseIf JeXY(typ) Or JeBublina(typ) Then
            bub = JeBublina(typ)
            xs = Hodnoty(s.XValues)
            ys = Hodnoty(s.Values)
            pocet = s.Points.Count
            If Not bub And s.MarkerStyle = XL_BEZ_ZNACKY Then
                PomocnaSerie s, xs, ys, pocet, mx0, mx1, y0, y1
            Else
                For j = 1 To pocet
                    If bub Then polo = PrumerBubliny(s.Points(j)) / 2 Else polo = VelikostZnacky(s, j) / 2
                    pridano = False
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
                                .X = gPL + (gPR - gPL) * (CDbl(xs(j)) - mx0) / (mx1 - mx0)
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
                                .R = polo
                                .Bublina = bub
                            End With
                            pridano = True
                        End If
                    End If
                    If bub And Not pridano And polo > 0 And j <= UBound(xs) And j <= UBound(ys) Then
                        If JeCislo(xs(j)) And JeCislo(ys(j)) Then
                            nk = nk + 1
                            ReDim Preserve kandidati(0 To nk)
                            kandidati(nk).X = gPL + (gPR - gPL) * (CDbl(xs(j)) - mx0) / (mx1 - mx0)
                            kandidati(nk).Y = gPT + (gPB - gPT) * (1 - (CDbl(ys(j)) - y0) / (y1 - y0))
                            kandidati(nk).R = polo
                        End If
                    End If
                Next j
            End If
        End If
    Next i
    Dim drzet As Boolean  ' a bubble without a label is an obstacle unless an equal or larger one covers it
    For k = 1 To nk
        drzet = True
        For j = 1 To UBound(gP)
            If gP(j).Bublina Then
                If StejnaBublina(gP(j).X, gP(j).Y, gP(j).R, kandidati(k)) Then drzet = False
            End If
        Next j
        For j = 1 To UBound(gBub)
            If StejnaBublina(gBub(j).X, gBub(j).Y, gBub(j).R, kandidati(k)) Then drzet = False
        Next j
        If drzet Then
            ReDim Preserve gBub(0 To UBound(gBub) + 1)
            gBub(UBound(gBub)) = kandidati(k)
        End If
    Next k
End Function

Private Function StejnaBublina(ByVal x As Double, ByVal y As Double, ByVal r As Double, b As BodBubliny) As Boolean
    StejnaBublina = Abs(x - b.X) < 0.001 And Abs(y - b.Y) < 0.001 And r >= b.R
End Function

' A series without markers: all x equal = vertical line (Index 80/100/120, a quadrant border), with labels it
' reserves space at the top of the chart; all y equal = horizontal line (dividing line). Only points with both
' values count.
Private Sub PomocnaSerie(s As Object, xs As Variant, ys As Variant, ByVal pocet As Long, ByVal x0 As Double, _
                         ByVal x1 As Double, ByVal y0 As Double, ByVal y1 As Double)
    Dim j As Long, n As Long, gx As Double, gy As Double, stejneX As Boolean, stejneY As Boolean, popisky As Boolean
    Dim v As Double
    stejneX = True
    stejneY = True
    For j = 1 To pocet
        If j <= UBound(xs) And j <= UBound(ys) Then
            If JeCislo(xs(j)) And JeCislo(ys(j)) Then
                n = n + 1
                If n = 1 Then
                    gx = xs(j)
                    gy = ys(j)
                Else
                    If xs(j) <> gx Then stejneX = False
                    If ys(j) <> gy Then stejneY = False
                End If
            End If
        End If
        If s.Points(j).HasDataLabel Then popisky = True
    Next j
    If n >= 1 And stejneX Then
        PridejCislo gCary, gPL + (gPR - gPL) * (gx - x0) / (x1 - x0)
        If popisky Then gRezerva = REZERVA_INDEX
    ElseIf n >= 2 And stejneY And Not gMaDelici Then
        v = gPT + (gPB - gPT) * (1 - (gy - y0) / (y1 - y0))
        If gPT + UVNITR < v And v < gPB - UVNITR Then
            gMaDelici = True
            gDelici = v
        End If
    End If
End Sub

' An area or line series with one value everywhere (an average) = dividing line.
Private Sub KonstantniSerie(s As Object, ByVal y0 As Double, ByVal y1 As Double)
    Dim hod As Variant, j As Long, n As Long, prvni As Double, stejne As Boolean, v As Double
    On Error GoTo Konec  ' a series without values
    hod = Hodnoty(s.Values)
    On Error GoTo 0
    stejne = True
    For j = 1 To UBound(hod)
        If JeCislo(hod(j)) Then
            n = n + 1
            If n = 1 Then
                prvni = hod(j)
            ElseIf hod(j) <> prvni Then
                stejne = False
            End If
        End If
    Next j
    If n >= 2 And stejne Then
        v = gPT + (gPB - gPT) * (1 - (prvni - y0) / (y1 - y0))
        If gPT + UVNITR < v And v < gPB - UVNITR Then
            gMaDelici = True
            gDelici = v
        End If
    End If
Konec:
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
    If PrazdnyText(dl) Then  ' the chart font of the label (ChartFont)
        fs = dl.Font.Size
        nazev = dl.Font.Name
        If dl.Font.Bold Then b = msoTrue Else b = msoFalse
    Else
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
    For i = 1 To UBound(gP)  ' leader lines of the series: a label in a column would hang without a link to its point
        If gP(i).Rezim <> "N" Then ZapniVodiciCary ch.SeriesCollection(gP(i).Serie)
    Next i
    VratPrazdne ch
    For i = 1 To UBound(gP)
        Set dl = ch.SeriesCollection(gP(i).Serie).Points(gP(i).Bod).DataLabel
        Set tf = dl.Format.TextFrame2
        If gP(i).Fs <> gP(i).FsPuvodni Then
            If PrazdnyText(dl) Then dl.Font.Size = gP(i).Fs Else tf.TextRange.Font.Size = gP(i).Fs
        End If
        If gP(i).Rezim <> "N" Then  ' "N" = not placed by the fallback: original position and size, only the font
            If gP(i).VlastniSirka Then  ' without a custom width PowerPoint would wrap a long label
                dl.Width = gP(i).W
                dl.Height = gP(i).H
            ElseIf tf.AutoSize <> AUTO_VELIKOST Or Abs(dl.Width - gP(i).W) > TOLERANCE_ZAPISU Or _
                   Abs(dl.Height - gP(i).H) > TOLERANCE_ZAPISU Then
                ' manual size (an earlier run, or the author): back to automatic; AutoSize reports 1 even for a
                ' size set by hand and only assigning it again resets the size (probe of 15 Nov 2026)
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

Private Sub ZapniVodiciCary(s As Object)
    On Error Resume Next  ' Series.HasLeaderLines: Office 2013+
    If Not s.HasLeaderLines Then s.HasLeaderLines = True
End Sub

' Empty labels (no text) the author moved away from their points: once the labels around them move, their
' leader lines hang in the air (bubble charts with a grey series under the coloured ones, sample of 15 Nov 2026).
' They go back to their default position, where PowerPoint draws no leader line.
Private Sub VratPrazdne(ch As Object)
    Dim i As Long, j As Long, s As Object, dl As Object, typ As Long, pozice As Long, txt As String, l As Double
    For i = 1 To ch.SeriesCollection.Count
        Set s = ch.SeriesCollection(i)
        typ = s.ChartType
        If JeBublina(typ) Or JeXY(typ) Then
            If JeBublina(typ) Or s.MarkerStyle <> XL_BEZ_ZNACKY Then
                For j = 1 To s.Points.Count
                    If s.Points(j).HasDataLabel Then
                        Set dl = s.Points(j).DataLabel
                        txt = "x"
                        l = 0#
                        pozice = 0
                        On Error Resume Next  ' read into variables: an error inside an If would enter its branch
                        txt = dl.Text
                        l = dl.Left
                        pozice = dl.Position
                        On Error GoTo 0
                        If Len(txt) = 0 And Not JeNaN(l) And pozice = XL_POLOHA_VLASTNI Then
                            On Error Resume Next
                            dl.Position = XL_POLOHA_VPRAVO
                            On Error GoTo 0
                        End If
                    End If
                Next j
            End If
        End If
    Next i
End Sub

' ---------- the state put aside ----------
Public Sub UlozStav(s As StavGrafu)
    s.W = gW
    s.H = gH
    s.PL = gPL
    s.PT = gPT
    s.PR = gPR
    s.PB = gPB
    s.CaL = gCaL
    s.CaT = gCaT
    s.P = gP
    s.Prek = gPrek
    s.Cary = gCary
    s.MaDelici = gMaDelici
    s.Delici = gDelici
    s.Rezerva = gRezerva
    s.OsaX0 = gOsaX0
    s.OsaX1 = gOsaX1
    s.OsaY0 = gOsaY0
    s.OsaY1 = gOsaY1
    s.OsaXVolna = gOsaXVolna
    s.OsaYVolna = gOsaYVolna
    s.CarTvaru = gCarTvaru
    s.Bub = gBub
End Sub

Public Sub ObnovStav(s As StavGrafu)
    gW = s.W
    gH = s.H
    gPL = s.PL
    gPT = s.PT
    gPR = s.PR
    gPB = s.PB
    gCaL = s.CaL
    gCaT = s.CaT
    gP = s.P
    gPrek = s.Prek
    gCary = s.Cary
    gMaDelici = s.MaDelici
    gDelici = s.Delici
    gRezerva = s.Rezerva
    gOsaX0 = s.OsaX0
    gOsaX1 = s.OsaX1
    gOsaY0 = s.OsaY0
    gOsaY1 = s.OsaY1
    gOsaXVolna = s.OsaXVolna
    gOsaYVolna = s.OsaYVolna
    gCarTvaru = s.CarTvaru
    gBub = s.Bub
End Sub

' ---------- font size and defects ----------
' The most common value (on a tie the one found first); pocet = how many of the values are used, 0 gives 0.
Private Function Nejcastejsi(cisla() As Double, ByVal pocet As Long) As Double
    Dim i As Long, j As Long, ruzne() As Double, pocty() As Long, m As Long, nej As Long
    If pocet = 0 Then Exit Function
    ReDim ruzne(1 To pocet)
    ReDim pocty(1 To pocet)
    For i = 1 To pocet
        For j = 1 To m
            If ruzne(j) = cisla(i) Then Exit For
        Next j
        If j > m Then
            m = m + 1
            ruzne(m) = cisla(i)
        End If
        pocty(j) = pocty(j) + 1
    Next i
    nej = 1
    For j = 2 To m
        If pocty(j) > pocty(nej) Then nej = j
    Next j
    Nejcastejsi = ruzne(nej)
End Function

' The most common font size the labels of the model have in the chart (before any change), 0 = no label.
Public Function PismoVGrafu() As Double
    Dim i As Long, n As Long, fs() As Double
    n = UBound(gP)
    If n = 0 Then Exit Function
    ReDim fs(1 To n)
    For i = 1 To n
        fs(i) = gP(i).FsPuvodni
    Next i
    PismoVGrafu = Nejcastejsi(fs, n)
End Function

' Gives all labels the most common font size (on a tie the one found first); returns the size,
' zmeneno = number of labels changed.
Public Function SjednotPismo(ByRef zmeneno As Long) As Double
    Dim n As Long, i As Long, fs() As Double
    n = UBound(gP)
    zmeneno = 0
    If n = 0 Then Exit Function
    ReDim fs(1 To n)
    For i = 1 To n
        fs(i) = gP(i).Fs
    Next i
    SjednotPismo = Nejcastejsi(fs, n)
    For i = 1 To n
        If gP(i).Fs <> SjednotPismo Then
            gP(i).Fs = SjednotPismo
            zmeneno = zmeneno + 1
        End If
    Next i
End Function

' The most common font size of the data labels of a chart that is not laid out (XY series with markers and bubble
' series, labels PowerPoint draws), 0 = no label or a chart that cannot be read. A light read: no geometry,
' only the size.
Public Function PismoPopisku(ch As Object) As Double
    Dim i As Long, j As Long, s As Object, dl As Object, n As Long, fs() As Double, v As Double
    On Error GoTo Konec
    ReDim fs(1 To 16)
    For i = 1 To ch.SeriesCollection.Count
        Set s = ch.SeriesCollection(i)
        If JeXY(s.ChartType) Or JeBublina(s.ChartType) Then
            If JeBublina(s.ChartType) Or s.MarkerStyle <> XL_BEZ_ZNACKY Then
                For j = 1 To s.Points.Count
                    If s.Points(j).HasDataLabel Then
                        Set dl = s.Points(j).DataLabel
                        If Len(dl.Text) > 0 And Not JeNaN(dl.Left) Then
                            v = VelikostPopisku(dl)
                            If v > 0 Then
                                n = n + 1
                                If n > UBound(fs) Then ReDim Preserve fs(1 To 2 * n)
                                fs(n) = v
                            End If
                        End If
                    End If
                Next j
            End If
        End If
    Next i
    PismoPopisku = Nejcastejsi(fs, n)
Konec:
End Function

Private Function VelikostPopisku(dl As Object) As Double
    Dim v As Double
    On Error Resume Next
    If PrazdnyText(dl) Then
        v = dl.Font.Size
    Else
        v = dl.Format.TextFrame2.TextRange.Font.Size
        If v < 1 Or v > 4000 Then v = dl.Format.TextFrame2.TextRange.Runs(1).Font.Size  ' mixed formatting: first run
    End If
    If v >= 1 And v <= 4000 Then VelikostPopisku = v
End Function

' The text frame of a label with text can be empty (length 0, no runs); its font then reports the default 18 pt
' while the label is drawn with the chart font of the label, DataLabel.Font (probe of 15 Nov 2026).
Private Function PrazdnyText(dl As Object) As Boolean
    On Error Resume Next
    PrazdnyText = (dl.Format.TextFrame2.TextRange.Length = 0)
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
        ' a point is a square max(3, radius); its own point counts only for a bubble, bubbles without a label too
        For k = 1 To UBound(gBub)
            If Prekryv(ra, BublinaObd(k, 0#), -0.5) Then v.PopisekPresBod = v.PopisekPresBod + 1
        Next k
        If gP(i).Bublina Then
            If Prekryv(ra, BodObd(i, 0#), -0.5) Then v.PopisekPresBod = v.PopisekPresBod + 1
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
                If Prekryv(ra, BodObd(j, 0#), -0.5) Then
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
