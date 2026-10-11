Option Explicit

' Pre-adjustment of a chart whose placement does not come out without defects (approved on 10 Oct 2026; port of
' popisky/predupravy.py, expressions in the same order): a wider axis range and a smaller label font, with one
' font size for all charts that are laid out (approved on 11 Oct 2026).
' The font size is decided for the whole run, not chart by chart. Module Popisky first asks every chart for the
' largest size it comes out with: its own, or smaller by PISMO_KROK up to PISMO_KROK * PISMO_KROKU. At a given
' size the chart is tried as it is and then with the axis-range rungs from the mildest (PrickyOs), so a wider
' range goes before a smaller font. The common size is the smallest one a chart asks for or has, and every chart
' is placed with it. An axis range grows by OSA_KROK up to 1 + OSA_KROK * OSA_KROKU times. A range is widened
' around the quadrant dividing lines (slide shapes, they do not move with the axis), so every point stays in its
' quadrant; without a dividing line around the middle of the plot area. An axis with tick labels is not changed
' (its numbers would stop being round), neither is the X axis with two or more vertical lines from the slide.
' Data points are never changed.
' The chart remembers its state before the pre-adjustment in shape tags: the same file gets new data and is
' processed again, so the next run first returns the axis ranges and the font size (VratOsy, PismoPredUpravou)
' and only then decides again. Without tags the chart is not pre-adjusted at all and keeps its own font size.
' Object model (probe of 10 Oct 2026): labels written after an axis change land where they are written; labels
' not written keep reporting their old position (they are drawn shifted with their points); an axis returned
' to automatic reports the automatic range at once; shape tags survive saving.

Public Const PISMO_KROK As Double = 0.5
Public Const PISMO_KROKU As Long = 4
Public Const OSA_KROK As Double = 0.1
Public Const OSA_KROKU As Long = 3
Private Const ZNACKA_OSY As String = "POPISKY_OSY"      ' two axes, each: changed|min auto|min|max auto|max|set min|set max
Private Const ZNACKA_PISMO As String = "POPISKY_PISMO"  ' original size|set size
Private Const ZNACKA_ZKOUSKA As String = "POPISKY_ZKOUSKA"
Private Const XL_HODNOTY As Long = 2
Private Const XL_KATEGORIE As Long = 1
Private Const SHODA As Double = 0.000000001             ' relative tolerance of a number stored in a tag

Private mZakX() As Double, mZakY() As Double, mZakFs() As Double, mZakCary() As Double

' ---------- the model at a font size and an axis-range rung ----------
' Remembers the model before the pre-adjustment: points, font sizes and vertical lines.
Public Sub ZapamatujZaklad()
    Dim i As Long, n As Long
    n = UBound(gP)
    ReDim mZakX(0 To n)
    ReDim mZakY(0 To n)
    ReDim mZakFs(0 To n)
    For i = 1 To n
        mZakX(i) = gP(i).X
        mZakY(i) = gP(i).Y
        mZakFs(i) = gP(i).Fs
    Next i
    mZakCary = gCary
End Sub

' Axis-range rungs of the chart in the model from the mildest: kroky(1..n) = steps of OSA_KROK, obeOsy(1..n) = both
' axes (on the same step the X axis alone goes first). Returns n; an axis that may not be changed takes its
' rungs away.
Public Function PrickyOs(kroky() As Long, obeOsy() As Boolean) As Long
    Dim a As Long, o As Long, n As Long
    ReDim kroky(0 To 2 * OSA_KROKU)
    ReDim obeOsy(0 To 2 * OSA_KROKU)
    If Not (gOsaXVolna And gCarTvaru <= 1) Then Exit Function
    For a = 1 To OSA_KROKU
        For o = 0 To IIf(gOsaYVolna, 1, 0)
            n = n + 1
            kroky(n) = a
            obeOsy(n) = (o = 1)
        Next o
    Next a
    PrickyOs = n
End Function

' The point the axis ranges are widened around (chart frame coordinates).
Private Sub Stred(ByRef cx As Double, ByRef cy As Double)
    If gCarTvaru = 1 Then cx = mZakCary(1) Else cx = (gPL + gPR) / 2
    If gMaDelici Then cy = gDelici Else cy = (gPT + gPB) / 2
End Sub

' Sets the model: the font size of the labels (velikost <= 0 = their own) and an axis-range rung (kroky = 0 = the
' ranges of the chart), that is the points and the helper-series lines.
Public Sub PouzijUpravu(ByVal velikost As Double, ByVal kroky As Long, ByVal obeOsy As Boolean)
    Dim i As Long, fx As Double, fy As Double, cx As Double, cy As Double
    fx = 1# + OSA_KROK * kroky
    If obeOsy Then fy = fx Else fy = 1#
    Stred cx, cy
    For i = 1 To UBound(gP)
        If velikost > 0 Then gP(i).Fs = velikost Else gP(i).Fs = mZakFs(i)
        If kroky > 0 Then
            gP(i).X = cx + (mZakX(i) - cx) / fx
            gP(i).Y = cy + (mZakY(i) - cy) / fy
        Else
            gP(i).X = mZakX(i)
            gP(i).Y = mZakY(i)
        End If
    Next i
    gCary = mZakCary
    If kroky > 0 Then
        For i = gCarTvaru + 1 To UBound(gCary)
            gCary(i) = cx + (mZakCary(i) - cx) / fx
        Next i
    End If
End Sub

' Axis ranges of a rung in axis units; menitX / menitY = the rung changes that axis.
Private Sub RozsahUpravy(ByVal kroky As Long, ByVal obeOsy As Boolean, ByRef menitX As Boolean, ByRef x0 As Double, _
                         ByRef x1 As Double, ByRef menitY As Boolean, ByRef y0 As Double, ByRef y1 As Double)
    Dim f As Double, cx As Double, cy As Double, c As Double
    menitX = False
    menitY = False
    If kroky = 0 Then Exit Sub
    f = 1# + OSA_KROK * kroky
    Stred cx, cy
    c = gOsaX0 + (cx - gPL) / (gPR - gPL) * (gOsaX1 - gOsaX0)
    x0 = c - (c - gOsaX0) * f
    x1 = c + (gOsaX1 - c) * f
    menitX = True
    If obeOsy Then
        c = gOsaY1 - (cy - gPT) / (gPB - gPT) * (gOsaY1 - gOsaY0)
        y0 = c - (c - gOsaY0) * f
        y1 = c + (gOsaY1 - c) * f
        menitY = True
    End If
End Sub

' Note for the check report, "" = the chart was not pre-adjusted. pismo = font size before the pre-adjustment,
' velikost = after it, kroky / obeOsy = the axis-range rung, urcilSnimek = number of the slide whose chart set
' the common font size (0 = this chart).
Public Function PopisUpravy(ByVal pismo As Double, ByVal velikost As Double, ByVal kroky As Long, ByVal obeOsy As Boolean, _
                            ByVal urcilSnimek As Long) As String
    Dim s As String
    If velikost <> pismo Then
        s = "font " & CisloPrehledu(pismo) & " → " & CisloPrehledu(velikost) & " pt"
        If urcilSnimek > 0 Then s = s & " (common size, set by slide " & urcilSnimek & ")"
    End If
    If kroky > 0 Then
        If s <> "" Then s = s & ", "
        s = s & IIf(obeOsy, "range of both axes", "X axis range") & " " & CisloPrehledu(1# + OSA_KROK * kroky) & "×"
    End If
    If s <> "" Then PopisUpravy = "pre-adjustment: " & s
End Function

' ---------- memory of the chart (shape tags) ----------
' True when the shape accepts tags; without them the next run could not return the pre-adjustment.
Public Function MuzeZnacit(shp As Shape) As Boolean
    On Error GoTo Konec
    shp.Tags.Add ZNACKA_ZKOUSKA, "1"
    shp.Tags.Delete ZNACKA_ZKOUSKA
    MuzeZnacit = True
Konec:
End Function

' Returns the axis ranges an earlier run changed (tag ZNACKA_OSY) to the state before it, so that the chart is
' read with the ranges it has from its author (an automatic bound follows the new data). An axis somebody has
' changed since is left as it is. Call before Graf.NactiGraf: sets gStaraOsa and the ranges the labels are
' still drawn with. Returns True when a range was returned.
Public Function VratOsy(shp As Shape) As Boolean
    Dim z As String, c() As String, ch As Object, osa As Object, t As Long, od As Long
    gStaraOsa = False
    z = Znacka(shp, ZNACKA_OSY)
    If z = "" Then Exit Function
    SmazZnacku shp, ZNACKA_OSY
    c = Split(z, "|")
    If UBound(c) <> 13 Then Exit Function
    Set ch = shp.Chart
    If Not (ch.HasAxis(XL_KATEGORIE, 1) And ch.HasAxis(XL_HODNOTY, 1)) Then Exit Function
    gStX0 = ch.Axes(XL_KATEGORIE, 1).MinimumScale
    gStX1 = ch.Axes(XL_KATEGORIE, 1).MaximumScale
    gStY0 = ch.Axes(XL_HODNOTY, 1).MinimumScale
    gStY1 = ch.Axes(XL_HODNOTY, 1).MaximumScale
    For t = 0 To 1
        od = 7 * t
        If c(od) = "1" Then
            Set osa = ch.Axes(IIf(t = 0, XL_KATEGORIE, XL_HODNOTY), 1)
            If Not osa.MinimumScaleIsAuto And Not osa.MaximumScaleIsAuto Then
                If Shodne(osa.MinimumScale, Val(c(od + 5))) And Shodne(osa.MaximumScale, Val(c(od + 6))) Then
                    If c(od + 1) = "1" Then osa.MinimumScaleIsAuto = True Else osa.MinimumScale = Val(c(od + 2))
                    If c(od + 3) = "1" Then osa.MaximumScaleIsAuto = True Else osa.MaximumScale = Val(c(od + 4))
                    VratOsy = True
                End If
            End If
        End If
    Next t
    gStaraOsa = VratOsy
End Function

' Font size before an earlier pre-adjustment (tag ZNACKA_PISMO), when the labels still have the size that run
' set; otherwise 0 and the tag is dropped (somebody has changed the font since). velikost = size in the chart.
Public Function PismoPredUpravou(shp As Shape, ByVal velikost As Double) As Double
    Dim z As String, c() As String
    z = Znacka(shp, ZNACKA_PISMO)
    If z = "" Then Exit Function
    c = Split(z, "|")
    If UBound(c) = 1 Then
        If Abs(Val(c(1)) - velikost) < 0.01 And Val(c(0)) > velikost Then PismoPredUpravou = Val(c(0))
    End If
    If PismoPredUpravou = 0 Then SmazZnacku shp, ZNACKA_PISMO
End Function

' Writes the pre-adjustment to the chart: the axis ranges of the rung (kroky = 0 = none) and the tags for the next
' run. The font size itself is written with the labels by Graf.ZapisGraf. Call before ZapisGraf, the labels are
' then placed against the new ranges. pismo = font size before the pre-adjustment, velikost = after it.
Public Sub ZapisPredupravu(shp As Shape, ByVal pismo As Double, ByVal velikost As Double, ByVal kroky As Long, _
                           ByVal obeOsy As Boolean)
    Dim ch As Object, osaX As Object, osaY As Object, menitX As Boolean, menitY As Boolean
    Dim x0 As Double, x1 As Double, y0 As Double, y1 As Double, z As String
    SmazZnacku shp, ZNACKA_PISMO
    If velikost < pismo Then shp.Tags.Add ZNACKA_PISMO, CisloText(pismo) & "|" & CisloText(velikost)
    RozsahUpravy kroky, obeOsy, menitX, x0, x1, menitY, y0, y1
    If Not menitX Then Exit Sub
    Set ch = shp.Chart
    Set osaX = ch.Axes(XL_KATEGORIE, 1)
    Set osaY = ch.Axes(XL_HODNOTY, 1)
    z = ZnackaOsy(osaX, True, x0, x1) & "|" & ZnackaOsy(osaY, menitY, y0, y1)
    shp.Tags.Add ZNACKA_OSY, z  ' the tag first: a changed range without it could never be returned
    osaX.MinimumScale = x0
    osaX.MaximumScale = x1
    If menitY Then
        osaY.MinimumScale = y0
        osaY.MaximumScale = y1
    End If
End Sub

Private Function ZnackaOsy(osa As Object, ByVal menit As Boolean, ByVal mn As Double, ByVal mx As Double) As String
    If Not menit Then
        ZnackaOsy = "0|0|0|0|0|0|0"
        Exit Function
    End If
    ZnackaOsy = "1|" & IIf(osa.MinimumScaleIsAuto, "1", "0") & "|" & CisloText(osa.MinimumScale) & "|" & _
        IIf(osa.MaximumScaleIsAuto, "1", "0") & "|" & CisloText(osa.MaximumScale) & "|" & CisloText(mn) & "|" & CisloText(mx)
End Function

Private Function Znacka(shp As Shape, ByVal nazev As String) As String
    On Error Resume Next  ' a shape without tags
    Znacka = shp.Tags(nazev)
End Function

Private Sub SmazZnacku(shp As Shape, ByVal nazev As String)
    On Error Resume Next  ' a missing tag
    shp.Tags.Delete nazev
End Sub

Private Function Shodne(ByVal a As Double, ByVal b As Double) As Boolean
    Shodne = Abs(a - b) <= SHODA * Max2(1#, Max2(Abs(a), Abs(b)))
End Function

' Number for a tag: Str$ always uses a decimal point and Val reads it back in any locale.
Private Function CisloText(ByVal d As Double) As String
    CisloText = Trim$(Str$(d))
End Function

' Number for the report: 9.5 or 1.3, with a decimal point in any locale.
Public Function CisloPrehledu(ByVal d As Double) As String
    CisloPrehledu = Trim$(Str$(d))
    If Left$(CisloPrehledu, 1) = "." Then CisloPrehledu = "0" & CisloPrehledu
End Function
