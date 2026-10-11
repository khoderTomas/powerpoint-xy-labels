Option Explicit

' Ribbon button "Lay out labels": asks which slides, which placement and where the result goes (dialog
' Nastaveni). By default it saves a copy of the open presentation (including unsaved changes) next to the
' original, places the data labels in the XY charts of the chosen slides of the copy and opens it; the original
' is not changed. With "Edit the open presentation" it places the labels in the open presentation itself and
' leaves saving to the user.
' Everything runs inside PowerPoint (modules Graf, Rozlozeni, Blizko, Predupravy, Pisma), no external program.
' The label font size is one for all charts laid out in a run (approved 11 Oct 2026): see Zpracuj.
' ZpracujSoubor and ZpracujOtevrenou are the entry points for tests (tests/test_vba.py): they process a file
' without a window, or an open presentation, and return the results as JSON.
' The texts the user sees are English (since 0.7.0, 11 Oct 2026); identifiers stay Czech.
' Non-ASCII text in string literals is converted to ChrW by sestav_doplnek.py (VBA source is in the system code page).

Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, _
    ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)

Private Const VERZE As String = "0.8.0"
Private Const NAZEV_DOPLNKU As String = "RozmistiPopisky"  ' file name, TEMP folder and settings key: not translated
Private Const TITULEK As String = "Lay out labels"
Private Const PRIPONA As String = "-labels"
Private Const PRIPONA_STARA As String = "-popisky"         ' output suffix of versions up to 0.6
Private Const PRIPONA_PREHLEDU As String = "-check.html"
Private Const PRIPONA_ZALOHY As String = "-backup-"
Private Const PP_PPTX As Long = 24
Private Const PP_PPTM As Long = 25                 ' ppSaveAsOpenXMLPresentationMacroEnabled
Private Const PP_PPT As Long = 1                   ' ppSaveAsPresentation
Private Const POCET_ZALOH As Long = 3
Private Const ZNACKA_PRESKOCIT As String = "#labels-off"
Private Const ZNACKA_PRESKOCIT_STARA As String = "#popisky-ne"  ' marker of versions up to 0.6, still accepted
Private Const STAV_UPRAVENO As String = "laid out"
Private Const STAV_BEZE_ZMENY As String = "unchanged"
Private Const STAV_PRESKOCENO As String = "skipped"
Private Const STAV_KONTROLA As String = "check only"
Private Const ZALOHA_TEXT As String = "Backup of the state before the edit: "
Private Const BEZ_GRAFU As String = "The presentation has no XY chart with labels from cells, nothing was changed."
Public Const STYL_SLOUPCE As String = "sloupce"   ' aligned columns (approved 8 Oct 2026)
Public Const STYL_KRATSI As String = "kratsi"     ' shorter leader lines (module Blizko)
Public Const KAM_KOPIE As String = "kopie"        ' result: a new presentation next to the original (default)
Public Const KAM_OTEVRENA As String = "otevrena"  ' result: the open presentation itself (approved 10 Oct 2026)
Private Const PP_VYBER_SNIMKU As Long = 1          ' ppSelectionSlides
Private Const MB_SETFOREGROUND As Long = &H10000

Private Type Vysledek
    CisloSlidu As Long
    Jmeno As String
    Stav As String
    Poznamka As String
    Varovani As Boolean
    Popisku As Long
    MaPred As Boolean
    Pred As Vady
    MaPo As Boolean
    Po As Vady
    Primo As Long
    Geometrie As String
    Cas As Double            ' seconds spent by the placement
    Model As String          ' JSON model for the parity test of the shorter-lines placement
    Predupraveno As Boolean  ' the font size or the axis ranges were changed (module Predupravy)
End Type

Private mV() As Vysledek

' A chart whose labels are to be placed: read in the first pass and placed in the last one, because the font
' size is common to all charts and every chart has to be read before the first one is written.
Private Type Prace
    Tvar As Shape
    Vysl As Long             ' index of its row in mV
    Model As StavGrafu       ' the chart as read, with the font size unified
    PuvL() As Double         ' original positions and sizes: labels the fallback cannot place stay there
    PuvT() As Double
    PuvW() As Double
    PuvH() As Double
    Poznamky As String
    Nahradni As String       ' fonts measured with a substitute
    Vraceno As Boolean       ' a pre-adjustment of an earlier run was returned
    Pismo As Double          ' font size before the pre-adjustment
    Chce As Double           ' the largest font size the chart comes out with
    Spolecne As Boolean      ' takes part in the common font size (pre-adjustment on, the shape accepts tags)
    Cas As Single            ' seconds spent on the question for the font size
    ZkVelikost() As Double   ' font sizes tried so far (ZkusVelikost), from index 1
    ZkBezVad() As Boolean    ' the size comes out without defects...
    ZkKroky() As Long        ' ...with this axis-range rung
    ZkObeOsy() As Boolean
End Type

Private mPrace() As Prace
Private mJineSnimek() As Long     ' charts left with another font size than the common one: slide...
Private mJinePismo() As Double    ' ...and the label font size
Private mVetaVelikosti As String  ' the sentence about them for the check report, "" = none

' ---------- ribbon button ----------
Public Sub RozmistiPopiskyTlacitko(control As IRibbonControl)
    RozmistiPopisky
End Sub

Public Sub RozmistiPopisky()
    Dim pres As Presentation, kopie As Presentation, styl As String, snimky As String, kam As String
    Dim prac As String, vystup As String, prehled As String, kod As Long, popis As String, zaloha As String
    Dim hotovo As String
    On Error GoTo Chyba
    If Presentations.Count = 0 Then
        Zprava "First open the presentation whose labels you want to lay out.", vbExclamation
        Exit Sub
    End If
    Set pres = ActivePresentation
    If Not ZeptejSe(pres, styl, snimky, kam) Then Exit Sub
    prac = PracovniSlozka()

    If kam = KAM_OTEVRENA Then
        kod = UpravOtevrenou(pres, styl, snimky, prac, prehled, zaloha)
    Else
        vystup = CilovaCesta(pres, ZakladNazvu(pres.Name), prac)
        prehled = Left$(vystup, Len(vystup) - 5) & PRIPONA_PREHLEDU
        pres.SaveCopyAs vystup, PP_PPTX
        Set kopie = Presentations.Open(vystup, msoFalse, msoFalse, msoFalse)
        kod = Zpracuj(kopie, True, styl, snimky, True)
        If kod = 2 Then
            kopie.Saved = msoTrue
            kopie.Close
            Set kopie = Nothing
            CreateObject("Scripting.FileSystemObject").DeleteFile vystup
        Else
            kopie.Save
            kopie.Close
            Set kopie = Nothing
            NapisPrehled prehled, pres.Name, vystup, "", styl, snimky
        End If
    End If
    If kod = 2 Then
        If snimky = "" Then
            Zprava BEZ_GRAFU, vbInformation
        Else
            Zprava "The selected slides have no XY chart with labels from cells, nothing was changed.", vbInformation
        End If
        Exit Sub
    End If
    If kam <> KAM_OTEVRENA Then Presentations.Open vystup
    If kod = 3 Then
        hotovo = "Some charts were left unchanged or need a check. The report opens in the browser."
        If zaloha <> "" Then hotovo = hotovo & vbLf & vbLf & ZALOHA_TEXT & zaloha
        Zprava hotovo, vbInformation
        OtevriSoubor prehled
    ElseIf kam = KAM_OTEVRENA Then  ' no new window opens, so say that it is done
        hotovo = "The labels are laid out in the open presentation (charts laid out: " & PocetUpravenych(False)
        If PocetUpravenych(True) > 0 Then hotovo = hotovo & ", of which with a pre-adjusted font or axes: " & PocetUpravenych(True)
        Zprava hotovo & "). The add-in does not save the presentation." & vbLf & vbLf & ZALOHA_TEXT & zaloha, vbInformation
    End If
    Exit Sub
Chyba:
    popis = Err.Description
    On Error Resume Next
    If Not kopie Is Nothing Then
        kopie.Saved = msoTrue
        kopie.Close
    End If
    If prehled <> "" Then
        NapisPrehled prehled, pres.Name, vystup, popis, styl, snimky, zaloha
        OtevriSoubor prehled
        If kam = KAM_OTEVRENA Then
            popis = popis & vbLf & vbLf & "Some charts may already be changed, the report lists them. " & ZALOHA_TEXT & zaloha
        End If
    End If
    Zprava "The layout failed: " & popis, vbExclamation
End Sub

' Dialog before placement; False = cancelled. The placement style and the result are remembered for the next run.
Private Function ZeptejSe(pres As Presentation, ByRef styl As String, ByRef snimky As String, ByRef kam As String) As Boolean
    Dim f As Nastaveni, vybrane As String, zMiniatur As Boolean
    vybrane = VybraneSnimky(pres, zMiniatur)
    Set f = New Nastaveni
    f.Priprav pres.Slides.Count, vybrane, zMiniatur, GetSetting(NAZEV_DOPLNKU, "Nastaveni", "Styl", STYL_SLOUPCE), _
        GetSetting(NAZEV_DOPLNKU, "Nastaveni", "Vysledek", KAM_KOPIE)
    f.Show
    If f.Ok Then
        styl = f.Styl
        snimky = f.Snimky
        kam = f.Kam
        SaveSetting NAZEV_DOPLNKU, "Nastaveni", "Styl", styl
        SaveSetting NAZEV_DOPLNKU, "Nastaveni", "Vysledek", kam
        ZeptejSe = True
    End If
    Unload f
End Function

' Places the labels in the open presentation itself: no copy, and the presentation is not saved. First it saves a
' backup to the work folder (zaloha; an error there means nothing was touched). The check report goes next to
' the presentation (name-check.html), or to the work folder when the presentation has no local folder or the
' folder cannot be written to; prehled is set before the charts are touched. Returns the code of Zpracuj.
Private Function UpravOtevrenou(pres As Presentation, ByVal styl As String, ByVal snimky As String, ByVal prac As String, _
                                ByRef prehled As String, ByRef zaloha As String, _
                                Optional ByVal predupravy As Boolean = True) As Long
    Dim jmeno As String
    jmeno = pres.Name
    If InStrRev(jmeno, ".") > 0 Then jmeno = Left$(jmeno, InStrRev(jmeno, ".") - 1)
    zaloha = ZalohujOtevrenou(pres, prac, jmeno)
    jmeno = jmeno & PRIPONA_PREHLEDU
    prehled = SlozkaVystupu(pres, prac) & "\" & jmeno
    UpravOtevrenou = Zpracuj(pres, True, styl, snimky, predupravy)
    If UpravOtevrenou = 2 Then  ' nothing was changed, so there is nothing to go back to
        On Error Resume Next
        CreateObject("Scripting.FileSystemObject").DeleteFile zaloha
        zaloha = ""
        Exit Function
    End If
    On Error Resume Next
    NapisPrehled prehled, pres.Name, "", "", styl, snimky, zaloha
    If Err.Number = 0 Then Exit Function
    On Error GoTo 0
    prehled = prac & "\" & jmeno
    NapisPrehled prehled, pres.Name, "", "", styl, snimky, zaloha
End Function

' Backup of the open presentation before it is edited in place: name-backup-yyyymmdd-hhnnss in the work folder,
' in the format of the presentation. The POCET_ZALOH newest backups of a presentation are kept.
Private Function ZalohujOtevrenou(pres As Presentation, ByVal prac As String, ByVal jmeno As String) As String
    Dim konc As String, typ As Long, zaklad As String, fso As Object, f As Object, stare() As String
    Dim n As Long, i As Long, j As Long, s As String
    typ = PP_PPTX
    konc = "pptx"
    If InStrRev(pres.Name, ".") > 0 Then
        Select Case LCase$(Mid$(pres.Name, InStrRev(pres.Name, ".") + 1))
            Case "pptm"
                typ = PP_PPTM
                konc = "pptm"
            Case "ppt"
                typ = PP_PPT
                konc = "ppt"
        End Select
    End If
    zaklad = jmeno & PRIPONA_ZALOHY
    ZalohujOtevrenou = prac & "\" & zaklad & Format$(Now, "yyyymmdd-hhnnss") & "." & konc
    pres.SaveCopyAs ZalohujOtevrenou, typ

    Set fso = CreateObject("Scripting.FileSystemObject")
    ReDim stare(0 To 0)  ' backups of this presentation, newest first (the time stamp sorts as text)
    For Each f In fso.GetFolder(prac).Files
        If StrComp(Left$(f.Name, Len(zaklad)), zaklad, vbTextCompare) = 0 Then
            n = n + 1
            ReDim Preserve stare(0 To n)
            s = f.Name
            j = n - 1
            Do While j >= 1
                If StrComp(stare(j), s, vbTextCompare) < 0 Then
                    stare(j + 1) = stare(j)
                    j = j - 1
                Else
                    Exit Do
                End If
            Loop
            stare(j + 1) = s
        End If
    Next f
    On Error Resume Next  ' a backup that is open stays
    For i = POCET_ZALOH + 1 To n
        fso.DeleteFile prac & "\" & stare(i)
    Next i
End Function

Private Function PocetUpravenych(ByVal jenPredupravene As Boolean) As Long
    Dim i As Long
    For i = 1 To PocetVysledku()
        If mV(i).Stav = STAV_UPRAVENO And (mV(i).Predupraveno Or Not jenPredupravene) Then PocetUpravenych = PocetUpravenych + 1
    Next i
End Function

' Slides selected in the active window of the presentation as a normalized list ("" = none);
' zMiniatur = the selection is in the slide thumbnails (not a shape or text on the current slide).
Private Function VybraneSnimky(pres As Presentation, ByRef zMiniatur As Boolean) As String
    Dim w As DocumentWindow, sr As SlideRange, i As Long, s As String, chyba As String
    On Error GoTo Konec  ' nothing selected: Selection.SlideRange raises an error
    Set w = Application.ActiveWindow
    If StrComp(w.Presentation.FullName, pres.FullName, vbTextCompare) <> 0 Then Exit Function
    zMiniatur = (w.Selection.Type = PP_VYBER_SNIMKU)
    Set sr = w.Selection.SlideRange
    For i = 1 To sr.Count
        If i > 1 Then s = s & ","
        s = s & sr(i).SlideIndex
    Next i
    VybraneSnimky = RozsahSnimku(s, pres.Slides.Count, chyba)
Konec:
End Function

' ---------- entry point for tests ----------
' Processes a file without a window: upravit = place the labels and save in place, otherwise only count
' the defects. styl = STYL_SLOUPCE or STYL_KRATSI, snimky = slide numbers as typed in the dialog ("" = all),
' predupravy = False switches the pre-adjustment off (the button always has it on). Returns JSON.
Public Function ZpracujSoubor(ByVal cesta As String, ByVal upravit As Boolean, Optional ByVal styl As String = STYL_SLOUPCE, _
                              Optional ByVal snimky As String = "", Optional ByVal predupravy As Boolean = True) As String
    Dim pres As Presentation, popis As String, chyba As String
    On Error GoTo Chyba
    If styl <> STYL_SLOUPCE And styl <> STYL_KRATSI Then Err.Raise vbObjectError + 1, , "unknown placement style: " & styl
    cesta = Replace(cesta, "/", "\")  ' with mixed slashes PowerPoint opens the file but cannot save it (file in use)
    Set pres = Presentations.Open(cesta, IIf(upravit, msoFalse, msoTrue), msoFalse, msoFalse)
    If snimky <> "" Then
        snimky = RozsahSnimku(snimky, pres.Slides.Count, chyba)
        If chyba <> "" Then Err.Raise vbObjectError + 2, , chyba
    End If
    Zpracuj pres, upravit, styl, snimky, predupravy
    If upravit Then pres.Save
    pres.Close
    ZpracujSoubor = VysledkyJson()
    Exit Function
Chyba:
    popis = Err.Description
    On Error Resume Next
    If Not pres Is Nothing Then
        pres.Saved = msoTrue
        pres.Close
    End If
    ZpracujSoubor = "{""chyba"":" & TextJson(popis) & "}"
End Function

' The "Edit the open presentation" path of the button without the dialog and the messages. cesta = full name
' of a presentation that is already open; styl, snimky and predupravy as in ZpracujSoubor.
' Returns JSON {"kod": code of Zpracuj, "prehled": path of the check report, "zaloha": path of the backup,
' "vysledky": [...]}.
Public Function ZpracujOtevrenou(ByVal cesta As String, Optional ByVal styl As String = STYL_SLOUPCE, _
                                 Optional ByVal snimky As String = "", Optional ByVal predupravy As Boolean = True) As String
    Dim pres As Presentation, p As Presentation, prehled As String, kod As Long, chyba As String, zaloha As String
    On Error GoTo Chyba
    If styl <> STYL_SLOUPCE And styl <> STYL_KRATSI Then Err.Raise vbObjectError + 1, , "unknown placement style: " & styl
    cesta = Replace(cesta, "/", "\")
    For Each p In Presentations
        If StrComp(p.FullName, cesta, vbTextCompare) = 0 Then Set pres = p
    Next p
    If pres Is Nothing Then Err.Raise vbObjectError + 3, , "the presentation is not open: " & cesta
    If snimky <> "" Then
        snimky = RozsahSnimku(snimky, pres.Slides.Count, chyba)
        If chyba <> "" Then Err.Raise vbObjectError + 2, , chyba
    End If
    kod = UpravOtevrenou(pres, styl, snimky, PracovniSlozka(), prehled, zaloha, predupravy)
    ZpracujOtevrenou = "{""kod"":" & kod & ",""prehled"":" & TextJson(prehled) & ",""zaloha"":" & TextJson(zaloha) & _
        ",""vysledky"":" & VysledkyJson() & "}"
    Exit Function
Chyba:
    ZpracujOtevrenou = "{""chyba"":" & TextJson(Err.Description) & "}"
End Function

' ---------- processing ----------
' Only the slides in snimky (normalized list, "" = all). Three passes: every chart is read (PrectiGraf), the
' common font size is decided (SpolecnaVelikost) and only then the charts are placed and written (UmistiGraf).
' predupravy = a chart whose placement has defects may get wider axis ranges and a smaller font, and the font
' size is then the same in all charts that are laid out. Returns 0 all charts processed, 2 no suitable chart,
' 3 some charts unchanged, skipped, with a warning or left with another font size.
Private Function Zpracuj(pres As Presentation, ByVal upravit As Boolean, ByVal styl As String, ByVal snimky As String, _
                         ByVal predupravy As Boolean) As Long
    Dim sl As Slide, grafy As Collection, shp As Shape, znacka As String, i As Long, cast As Variant
    Dim cislo As Long, popis As String, vybrano() As Boolean, spolecna As Double, urcil As Long
    ReDim mV(0 To 0)
    ReDim mPrace(0 To 0)
    ReDim mJineSnimek(0 To 0)
    ReDim mJinePismo(0 To 0)
    mVetaVelikosti = ""
    ReDim vybrano(0 To pres.Slides.Count)
    For i = 1 To pres.Slides.Count
        vybrano(i) = (snimky = "")
    Next i
    If snimky <> "" Then
        For Each cast In Split(snimky, ",")
            vybrano(CLng(cast)) = True
        Next cast
    End If
    On Error GoTo Chyba
    For Each sl In pres.Slides
        If vybrano(sl.SlideIndex) Then
            Set grafy = GrafySlidu(sl)
            If grafy.Count > 0 Then
                NactiTvary sl
                znacka = ZnackaSlidu(sl)
                For Each shp In grafy
                    If MaXYSerii(shp.Chart) Then PrectiGraf sl, shp, znacka, upravit, predupravy
                Next shp
            End If
        End If
    Next sl
    If UBound(mPrace) > 0 Then
        If predupravy Then spolecna = SpolecnaVelikost(styl, urcil)
        For i = 1 To UBound(mPrace)
            UmistiGraf i, styl, spolecna, urcil
        Next i
        If spolecna > 0 Then
            JineVelikostiMimoVyber pres, vybrano
            mVetaVelikosti = PopisJinychVelikosti(spolecna)
        End If
    End If
    Erase mPrace
    UvolniPisma
    If UBound(mV) = 0 Then
        Zpracuj = 2
        Exit Function
    End If
    If mVetaVelikosti <> "" Then Zpracuj = 3
    For i = 1 To UBound(mV)
        If mV(i).Stav <> STAV_UPRAVENO Or mV(i).Varovani Then Zpracuj = 3
    Next i
    Exit Function
Chyba:
    cislo = Err.Number
    popis = Err.Description
    Erase mPrace
    UvolniPisma
    Err.Raise cislo, , popis
End Function

' First pass: reads the chart and adds its row to the results. A chart whose labels are to be placed is put
' aside in mPrace with the font size unified and a pre-adjustment of an earlier run returned.
' znacka = the skip marker found in the slide notes ("" = none).
Private Sub PrectiGraf(sl As Slide, shp As Shape, ByVal znacka As String, ByVal upravit As Boolean, ByVal predupravy As Boolean)
    Dim v As Vysledek, duvod As String, i As Long, n As Long, zmeneno As Long, velikost As Double
    Dim nahradni As String, vraceno As Boolean, puvPismo As Double, preskocit As Boolean, ceka As Boolean
    preskocit = (znacka <> "")
    v.CisloSlidu = sl.SlideIndex
    v.Jmeno = shp.Name
    gStaraOsa = False
    If upravit And Not preskocit Then vraceno = VratOsy(shp)  ' before reading: the chart with its own axis ranges
    duvod = NactiGraf(sl, shp)
    gStaraOsa = False
    If duvod <> "" Then  ' a chart the tool cannot handle is left as it is
        v.Stav = STAV_PRESKOCENO
        v.Poznamka = duvod
        PridejVysledek v
        Exit Sub
    End If
    If UBound(gP) = 0 Then Exit Sub
    v.Popisku = UBound(gP)
    v.Pred = SpocitejVady()
    v.MaPred = True
    nahradni = NahradniPisma()
    If preskocit Then
        v.Stav = STAV_PRESKOCENO
        v.Poznamka = "marker " & znacka & " in the slide notes"
        If upravit Then PridejJinou sl.SlideIndex, PismoVGrafu()
    ElseIf Not upravit Then
        v.Stav = STAV_KONTROLA
    Else
        ceka = True
        n = UBound(mPrace) + 1
        ReDim Preserve mPrace(0 To n)
        ReDim mPrace(n).PuvL(1 To UBound(gP))  ' original positions: labels the fallback cannot place stay there
        ReDim mPrace(n).PuvT(1 To UBound(gP))
        ReDim mPrace(n).PuvW(1 To UBound(gP))
        ReDim mPrace(n).PuvH(1 To UBound(gP))
        For i = 1 To UBound(gP)
            mPrace(n).PuvL(i) = gP(i).L
            mPrace(n).PuvT(i) = gP(i).T
            mPrace(n).PuvW(i) = gP(i).W
            mPrace(n).PuvH(i) = gP(i).H
        Next i
        velikost = SjednotPismo(zmeneno)
        If zmeneno > 0 Then
            mPrace(n).Poznamky = "font size unified to " & CisloPrehledu(velikost) & " pt (labels changed: " & zmeneno & ")"
        End If
        puvPismo = PismoPredUpravou(shp, velikost)
        If puvPismo > 0 Then  ' an earlier run made the font smaller: start from the size before it
            For i = 1 To UBound(gP)
                gP(i).Fs = puvPismo
            Next i
            velikost = puvPismo
            vraceno = True
        End If
        Set mPrace(n).Tvar = shp
        mPrace(n).Vysl = UBound(mV) + 1
        mPrace(n).Pismo = velikost
        mPrace(n).Chce = velikost
        mPrace(n).Vraceno = vraceno
        mPrace(n).Nahradni = nahradni
        If predupravy Then mPrace(n).Spolecne = MuzeZnacit(shp)
        ReDim mPrace(n).ZkVelikost(0 To 0)
        ReDim mPrace(n).ZkBezVad(0 To 0)
        ReDim mPrace(n).ZkKroky(0 To 0)
        ReDim mPrace(n).ZkObeOsy(0 To 0)
        UlozStav mPrace(n).Model
    End If
    If nahradni <> "" And Not ceka Then
        v.Poznamka = Pripoj(v.Poznamka, "font not installed: " & nahradni & " (measured with a substitute)")
    End If
    PridejVysledek v
End Sub

' Second pass: the common font size of the charts that take part = the smallest one a chart asks for or has
' (0 = no such chart); urcil = index in mPrace of the first chart with that size.
Private Function SpolecnaVelikost(ByVal styl As String, ByRef urcil As Long) As Double
    Dim i As Long, zacatek As Single, nej As Double, chce As Double
    urcil = 0
    For i = 1 To UBound(mPrace)
        If mPrace(i).Spolecne Then
            zacatek = Timer
            ObnovStav mPrace(i).Model
            ZapamatujZaklad
            chce = PotrebnaVelikost(i, styl)
            mPrace(i).Chce = chce
            mPrace(i).Cas = Timer - zacatek
            If urcil = 0 Or chce < nej Then
                nej = chce
                urcil = i
            End If
        End If
    Next i
    SpolecnaVelikost = nej
End Function

' The largest font size the chart in the model comes out with without defects (ZkusVelikost): its own, or smaller
' by PISMO_KROK up to PISMO_KROK * PISMO_KROKU. When none helps, its own: such a chart does not decide.
Private Function PotrebnaVelikost(ByVal i As Long, ByVal styl As String) As Double
    Dim p As Long, velikost As Double, kroky As Long, obeOsy As Boolean, mezera As Double
    Dim startL() As Double, startT() As Double, startW() As Double, startH() As Double
    For p = 0 To PISMO_KROKU
        velikost = mPrace(i).Pismo - PISMO_KROK * p
        If ZkusVelikost(i, styl, velikost, startL, startT, startW, startH, kroky, obeOsy, mezera) Then
            PotrebnaVelikost = velikost
            Exit Function
        End If
    Next p
    PotrebnaVelikost = mPrace(i).Pismo
End Function

' Places the model with the font size velikost: first the chart as it is (every gap: cramped columns are no reason
' to change a chart), then the axis-range rungs from the mildest (the full gap only). True = it comes out without
' defects, kroky / obeOsy = the rung and mezera = the gap; otherwise the model has the font size velikost and the
' ranges of the chart. The outcome for a size is remembered in mPrace(i): a second question only repeats the
' placement found, a known failure is not computed again.
Private Function ZkusVelikost(ByVal i As Long, ByVal styl As String, ByVal velikost As Double, startL() As Double, _
                              startT() As Double, startW() As Double, startH() As Double, ByRef kroky As Long, _
                              ByRef obeOsy As Boolean, ByRef mezera As Double) As Boolean
    Dim k As Long, j As Long, n As Long, odJ As Long, prKroky() As Long, prObe() As Boolean, bezVad As Boolean, znamy As Long
    For k = 1 To UBound(mPrace(i).ZkVelikost)
        If mPrace(i).ZkVelikost(k) = velikost Then znamy = k
    Next k
    n = PrickyOs(prKroky, prObe)
    If znamy > 0 Then
        If mPrace(i).ZkBezVad(znamy) Then  ' only the rung found (0 = the chart as it is)
            prKroky(0) = mPrace(i).ZkKroky(znamy)
            prObe(0) = mPrace(i).ZkObeOsy(znamy)
            n = 0
        Else
            odJ = n + 1
        End If
    Else
        znamy = UBound(mPrace(i).ZkVelikost) + 1
        ReDim Preserve mPrace(i).ZkVelikost(0 To znamy)
        ReDim Preserve mPrace(i).ZkBezVad(0 To znamy)
        ReDim Preserve mPrace(i).ZkKroky(0 To znamy)
        ReDim Preserve mPrace(i).ZkObeOsy(0 To znamy)
        mPrace(i).ZkVelikost(znamy) = velikost
    End If
    For j = odJ To n
        PouzijUpravu velikost, prKroky(j), prObe(j)
        mezera = Pokus(styl, prKroky(j) > 0, startL, startT, startW, startH, bezVad)
        If bezVad Then
            kroky = prKroky(j)
            obeOsy = prObe(j)
            mPrace(i).ZkBezVad(znamy) = True
            mPrace(i).ZkKroky(znamy) = kroky
            mPrace(i).ZkObeOsy(znamy) = obeOsy
            ZkusVelikost = True
            Exit Function
        End If
    Next j
    mPrace(i).ZkBezVad(znamy) = False
    kroky = 0
    obeOsy = False
    PouzijUpravu velikost, 0, False
End Function

' Third pass: places the labels of chart i of mPrace with the common font size (0 = pre-adjustment off) and
' writes them. urcil = index in mPrace of the chart that set the common size.
Private Sub UmistiGraf(ByVal i As Long, ByVal styl As String, ByVal spolecna As Double, ByVal urcil As Long)
    Dim k As Long, iv As Long, velikost As Double, mezera As Double, kroky As Long, obeOsy As Boolean
    Dim startL() As Double, startT() As Double, startW() As Double, startH() As Double, bezVad As Boolean
    Dim puvL() As Double, puvT() As Double, puvW() As Double, puvH() As Double, nouze As Boolean, nevesly As String
    Dim nesedi As Long, poznamky As String, uprava As String, zacatek As Single, modelJs As String, urcilSnimek As Long, shp As Shape
    zacatek = Timer
    iv = mPrace(i).Vysl
    Set shp = mPrace(i).Tvar
    puvL = mPrace(i).PuvL
    puvT = mPrace(i).PuvT
    puvW = mPrace(i).PuvW
    puvH = mPrace(i).PuvH
    poznamky = mPrace(i).Poznamky
    velikost = mPrace(i).Pismo
    ObnovStav mPrace(i).Model
    ZapamatujZaklad
    If spolecna > 0 And mPrace(i).Spolecne Then
        velikost = spolecna
        If Not ZkusVelikost(i, styl, velikost, startL, startT, startW, startH, kroky, obeOsy, mezera) Then
            ' it does not come out without defects with the common size: the chart as it is
            mezera = Pokus(styl, False, startL, startT, startW, startH, bezVad)
        End If
    Else
        mezera = Pokus(styl, False, startL, startT, startW, startH, bezVad)
    End If
    If mezera >= 0 And styl = STYL_KRATSI Then
        modelJs = ModelJson(mezera, startL, startT, startW, startH, False, puvL, puvT, puvW, puvH)
    ElseIf mezera < 0 Then  ' the columns do not fit: fallback placement (both options)
        nouze = True
        mezera = Nouzove(puvL, puvT, puvW, puvH, startL, startT)
        If mezera >= 0 Then modelJs = ModelJson(mezera, startL, startT, startW, startH, True, puvL, puvT, puvW, puvH)
    End If
    mV(iv).Model = modelJs
    If mezera < 0 Then
        mV(iv).Stav = STAV_BEZE_ZMENY
        poznamky = "the labels do not fit into the chart, adjust by hand"
        If spolecna > 0 Then PridejJinou mV(iv).CisloSlidu, PismoVGrafu()
    Else
        ZapisPredupravu shp, mPrace(i).Pismo, velikost, kroky, obeOsy  ' before the labels: against the new axis ranges
        nesedi = ZapisGraf(shp)
        mV(iv).Po = SpocitejVady()
        mV(iv).MaPo = True
        For k = 1 To UBound(gP)
            If gP(k).Rezim = "D" Then mV(iv).Primo = mV(iv).Primo + 1
        Next k
        mV(iv).Stav = STAV_UPRAVENO
        If urcil > 0 And urcil <> i Then urcilSnimek = mV(mPrace(urcil).Vysl).CisloSlidu
        uprava = PopisUpravy(mPrace(i).Pismo, velikost, kroky, obeOsy, urcilSnimek)
        If uprava <> "" Then
            poznamky = Pripoj(poznamky, uprava)
            mV(iv).Predupraveno = True
        ElseIf mPrace(i).Vraceno Then
            poznamky = Pripoj(poznamky, "pre-adjustment from the previous run reverted")
        End If
        If nouze Then poznamky = Pripoj(poznamky, "the columns did not fit, labels placed closer to their points")
        If mezera < 1.5 Then poznamky = Pripoj(poznamky, "little room, labels closer together")
        For k = 1 To UBound(gP)
            If gP(k).Rezim = "N" Then
                If nevesly <> "" Then nevesly = nevesly & ", "
                nevesly = nevesly & gP(k).Txt
            End If
        Next k
        If nevesly <> "" Then
            poznamky = Pripoj(poznamky, "labels that did not fit stayed where they were (adjust by hand): " & nevesly)
            mV(iv).Varovani = True
        End If
        If nesedi > 0 Then
            poznamky = Pripoj(poznamky, "position after writing differs from the computed one by more than 0.5 pt (labels: " & nesedi & "), check them")
            mV(iv).Varovani = True
        End If
        mV(iv).Geometrie = GeometrieJson()
        If spolecna > 0 And velikost <> spolecna Then PridejJinou mV(iv).CisloSlidu, velikost
    End If
    If mPrace(i).Nahradni <> "" Then
        poznamky = Pripoj(poznamky, "font not installed: " & mPrace(i).Nahradni & " (measured with a substitute)")
    End If
    mV(iv).Poznamka = poznamky
    mV(iv).Cas = mPrace(i).Cas + (Timer - zacatek)
End Sub

' ---------- charts left with another font size ----------
' A chart that is not laid out with the common font size (slide not selected, skip marker, labels that do not
' fit, a shape without tags): its slide and its label font size.
Private Sub PridejJinou(ByVal snimek As Long, ByVal velikost As Double)
    Dim n As Long
    If velikost <= 0 Then Exit Sub
    n = UBound(mJineSnimek) + 1
    ReDim Preserve mJineSnimek(0 To n)
    ReDim Preserve mJinePismo(0 To n)
    mJineSnimek(n) = snimek
    mJinePismo(n) = velikost
End Sub

' The XY charts on the slides that are not selected: only their label font size is read, nothing is changed.
Private Sub JineVelikostiMimoVyber(pres As Presentation, vybrano() As Boolean)
    Dim sl As Slide, shp As Shape
    On Error Resume Next  ' the note is only informative, a chart that cannot be read is left out
    For Each sl In pres.Slides
        If Not vybrano(sl.SlideIndex) Then
            For Each shp In GrafySlidu(sl)
                If MaXYSerii(shp.Chart) Then PridejJinou sl.SlideIndex, PismoPopisku(shp.Chart)
            Next shp
        End If
    Next sl
End Sub

' Sentence for the check report about the charts whose labels have another size than the common one, in slide
' order; "" = there is none.
Private Function PopisJinychVelikosti(ByVal spolecna As Double) As String
    Dim i As Long, j As Long, n As Long, snimek As Long, velikost As Double, s As String, polozka As String
    n = UBound(mJineSnimek)
    For i = 2 To n  ' insertion sort by slide number, stable
        snimek = mJineSnimek(i)
        velikost = mJinePismo(i)
        j = i - 1
        Do While j >= 1
            If mJineSnimek(j) <= snimek Then Exit Do
            mJineSnimek(j + 1) = mJineSnimek(j)
            mJinePismo(j + 1) = mJinePismo(j)
            j = j - 1
        Loop
        mJineSnimek(j + 1) = snimek
        mJinePismo(j + 1) = velikost
    Next i
    For i = 1 To n
        If mJinePismo(i) <> spolecna Then
            polozka = "slide " & mJineSnimek(i) & " (" & CisloPrehledu(mJinePismo(i)) & " pt)"
            If InStr(1, ", " & s & ",", ", " & polozka & ",", vbBinaryCompare) = 0 Then
                If s <> "" Then s = s & ", "
                s = s & polozka
            End If
        End If
    Next i
    If s <> "" Then
        PopisJinychVelikosti = "Common label size of the charts laid out: " & CisloPrehledu(spolecna) & " pt. " & _
            "Other charts keep their own size: " & s & "."
    End If
End Function

' One placement of the model as it is: the columns, then the shorter lines for STYL_KRATSI. Returns the gap of the
' columns (-1 = they do not fit; jenPlna = only the full gap counts); bezVad = it fits and has no defects.
' startW/startH get the single-line sizes the placement works with and startL/startT the column start (the model
' for the parity test of the shorter lines).
Private Function Pokus(ByVal styl As String, ByVal jenPlna As Boolean, startL() As Double, startT() As Double, _
                       startW() As Double, startH() As Double, ByRef bezVad As Boolean) As Double
    Dim i As Long, n As Long, m As Double
    n = UBound(gP)
    m = Rozmisti(jenPlna)
    Pokus = m
    ReDim startL(1 To n)
    ReDim startT(1 To n)
    ReDim startW(1 To n)
    ReDim startH(1 To n)
    For i = 1 To n
        startW(i) = gP(i).W
        startH(i) = gP(i).H
    Next i
    bezVad = False
    If m < 0 Then Exit Function
    If styl = STYL_KRATSI Then
        For i = 1 To n
            startL(i) = gP(i).L
            startT(i) = gP(i).T
        Next i
        ZlepsiBlizko m
    End If
    bezVad = (SoucetVad(SpocitejVady()) = 0)
End Function

Private Function GrafySlidu(sl As Slide) As Collection
    Dim c As New Collection, s As Shape, k As Long
    For Each s In sl.Shapes
        If s.Type = msoGroup Then
            If s.Rotation = 0 Then  ' charts inside groups too, only unrotated ones
                For k = 1 To s.GroupItems.Count
                    If s.GroupItems(k).HasChart Then c.Add s.GroupItems(k)
                Next k
            End If
        ElseIf s.HasChart Then
            c.Add s
        End If
    Next s
    Set GrafySlidu = c
End Function

' The skip marker found in the slide notes, "" = none.
Private Function ZnackaSlidu(sl As Slide) As String
    Dim s As Shape, t As String
    On Error GoTo Konec  ' slide without notes
    If Not sl.HasNotesPage Then Exit Function
    For Each s In sl.NotesPage.Shapes
        If s.HasTextFrame Then
            t = s.TextFrame.TextRange.Text
            If InStr(1, t, ZNACKA_PRESKOCIT, vbTextCompare) > 0 Then
                ZnackaSlidu = ZNACKA_PRESKOCIT
            ElseIf ZnackaSlidu = "" And InStr(1, t, ZNACKA_PRESKOCIT_STARA, vbTextCompare) > 0 Then
                ZnackaSlidu = ZNACKA_PRESKOCIT_STARA
            End If
        End If
    Next s
Konec:
End Function

Private Function NahradniPisma() As String
    Dim i As Long, nazev As String, seznam As String
    For i = 1 To UBound(gP)
        If JeNahradni(gP(i).Pismo) Then
            nazev = NazevPisma(gP(i).Pismo)
            If InStr(1, "|" & seznam & "|", "|" & nazev & "|", vbTextCompare) = 0 Then
                If seznam <> "" Then seznam = seznam & "|"
                seznam = seznam & nazev
            End If
        End If
    Next i
    NahradniPisma = Replace(seznam, "|", ", ")
End Function

Private Function Pripoj(ByVal a As String, ByVal b As String) As String
    If a = "" Then Pripoj = b Else Pripoj = a & "; " & b
End Function

Private Sub PridejVysledek(v As Vysledek)
    ReDim Preserve mV(0 To UBound(mV) + 1)
    mV(UBound(mV)) = v
End Sub

Private Function PocetVysledku() As Long
    On Error Resume Next  ' the array is not allocated before the first run
    PocetVysledku = UBound(mV)
End Function

' ---------- JSON for tests ----------
Private Function VysledkyJson() As String
    Dim i As Long, s As String
    s = "["
    For i = 1 To PocetVysledku()
        If i > 1 Then s = s & ","
        With mV(i)
            s = s & "{""slide"":" & .CisloSlidu & ",""graf"":" & TextJson(.Jmeno) & ",""stav"":" & TextJson(.Stav) & _
                ",""poznamka"":" & TextJson(.Poznamka) & ",""popisku"":" & .Popisku & _
                ",""pred"":" & VadyJson(.Pred, .MaPred) & ",""po"":" & VadyJson(.Po, .MaPo) & _
                ",""primo"":" & .Primo & ",""geometrie"":" & IIf(.Geometrie = "", "[]", .Geometrie) & _
                ",""cas"":" & CisloJson(.Cas) & ",""model"":" & IIf(.Model = "", "null", .Model) & "}"
        End With
    Next i
    VysledkyJson = s & "]"
End Function

Private Function VadyJson(v As Vady, ByVal ma As Boolean) As String
    If Not ma Then
        VadyJson = "null"
        Exit Function
    End If
    VadyJson = "{""prekryv_popisku"":" & v.PrekryvPopisku & ",""popisek_pres_bod"":" & v.PopisekPresBod & _
        ",""popisek_pres_box"":" & v.PopisekPresBox & ",""mimo_graf"":" & v.MimoGraf & _
        ",""krizeni_car"":" & v.KrizeniCar & ",""cara_pres_popisek"":" & v.CaraPresPopisek & _
        ",""popisek_pres_caru"":" & v.PopisekPresCaru & "}"
End Function

Private Function GeometrieJson() As String
    Dim i As Long, s As String
    s = "["
    For i = 1 To UBound(gP)
        If i > 1 Then s = s & ","
        s = s & "{""text"":" & TextJson(gP(i).Txt) & ",""left"":" & CisloJson(gP(i).L) & ",""top"":" & CisloJson(gP(i).T) & _
            ",""w"":" & CisloJson(gP(i).W) & ",""h"":" & CisloJson(gP(i).H) & ",""fs"":" & CisloJson(gP(i).Fs) & _
            ",""x"":" & CisloJson(gP(i).X) & ",""y"":" & CisloJson(gP(i).Y) & ",""rezim"":" & TextJson(gP(i).Rezim) & "}"
    Next i
    GeometrieJson = s & "]"
End Function

' Model of the chart for the parity test of the shorter-lines placement: inputs, the column start and the result,
' every number as the 8 bytes of the Double in hex (decimal text would lose the last bits). nouze: fallback
' placement, with the original positions of the labels (pl, pt, pw, ph).
Private Function ModelJson(ByVal mezera As Double, startL() As Double, startT() As Double, startW() As Double, _
                           startH() As Double, ByVal nouze As Boolean, puvL() As Double, puvT() As Double, _
                           puvW() As Double, puvH() As Double) As String
    Dim i As Long, s As String
    s = "{""mezera"":" & HexJson(mezera) & ",""nouze"":" & IIf(nouze, "true", "false") & _
        ",""PL"":" & HexJson(gPL) & ",""PT"":" & HexJson(gPT) & _
        ",""PR"":" & HexJson(gPR) & ",""PB"":" & HexJson(gPB) & ",""rezerva"":" & HexJson(gRezerva) & ",""delici"":"
    If gMaDelici Then s = s & HexJson(gDelici) Else s = s & "null"
    s = s & ",""cary"":["
    For i = 1 To UBound(gCary)
        If i > 1 Then s = s & ","
        s = s & HexJson(gCary(i))
    Next i
    s = s & "],""prekazky"":["
    For i = 1 To UBound(gPrek)
        If i > 1 Then s = s & ","
        s = s & "[" & HexJson(gPrek(i).L) & "," & HexJson(gPrek(i).T) & "," & HexJson(gPrek(i).R) & "," & HexJson(gPrek(i).B) & "]"
    Next i
    s = s & "],""popisky"":["
    For i = 1 To UBound(gP)
        If i > 1 Then s = s & ","
        s = s & "{""text"":" & TextJson(gP(i).Txt) & ",""x"":" & HexJson(gP(i).X) & ",""y"":" & HexJson(gP(i).Y) & _
            ",""w"":" & HexJson(startW(i)) & ",""h"":" & HexJson(startH(i)) & ",""l0"":" & HexJson(startL(i)) & _
            ",""t0"":" & HexJson(startT(i)) & ",""l"":" & HexJson(gP(i).L) & ",""t"":" & HexJson(gP(i).T) & _
            ",""rezim"":" & TextJson(gP(i).Rezim)
        If nouze Then
            s = s & ",""pl"":" & HexJson(puvL(i)) & ",""pt"":" & HexJson(puvT(i)) & ",""pw"":" & HexJson(puvW(i)) & _
                ",""ph"":" & HexJson(puvH(i))
        End If
        s = s & "}"
    Next i
    ModelJson = s & "],""ladeni"":" & Ladeni() & "}"
End Function

Private Function HexJson(ByVal d As Double) As String
    Dim b(0 To 7) As Byte, i As Long, s As String
    CopyMemory b(0), d, 8
    For i = 0 To 7
        s = s & Right$("0" & Hex$(b(i)), 2)
    Next i
    HexJson = """" & s & """"
End Function

Private Function TextJson(ByVal s As String) As String
    Dim i As Long, c As Long, out As String
    out = """"
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        Select Case c
            Case 34
                out = out & "\"""
            Case 92
                out = out & "\\"
            Case 32 To 126
                out = out & ChrW(c)
            Case Else
                out = out & "\u" & Right$("000" & Hex$(c), 4)
        End Select
    Next i
    TextJson = out & """"
End Function

Private Function CisloJson(ByVal d As Double) As String
    CisloJson = Trim$(Str$(d))  ' Str$ always uses a decimal point
    If Left$(CisloJson, 1) = "." Then CisloJson = "0" & CisloJson
    If Left$(CisloJson, 2) = "-." Then CisloJson = "-0" & Mid$(CisloJson, 2)
End Function

' ---------- check report ----------
' vystup = path of the new presentation, "" = the open presentation itself was edited (zaloha = its backup).
Private Sub NapisPrehled(ByVal cesta As String, ByVal vstup As String, ByVal vystup As String, ByVal chyba As String, _
                         ByVal styl As String, ByVal snimky As String, Optional ByVal zaloha As String = "")
    Dim i As Long, radky As String, souhrn As String, nUpr As Long, nBez As Long, nPres As Long, tabulka As String
    Dim volby As String, cil As String
    For i = 1 To PocetVysledku()
        With mV(i)
            Select Case .Stav
                Case STAV_UPRAVENO: nUpr = nUpr + 1
                Case STAV_BEZE_ZMENY: nBez = nBez + 1
                Case STAV_PRESKOCENO: nPres = nPres + 1
            End Select
            radky = radky & "<tr><td class='cislo'>" & .CisloSlidu & "</td><td>" & Html(.Stav) & "</td><td class='cislo'>" & _
                .Popisku & "</td><td>" & Html(VadyText(.Pred, .MaPred)) & "</td><td>" & Html(VadyText(.Po, .MaPo)) & _
                "</td><td>" & Html(.Poznamka) & "</td></tr>"
        End With
    Next i
    If chyba <> "" Then
        souhrn = "Processing failed: " & Html(chyba)
    ElseIf PocetVysledku() = 0 Then
        souhrn = BEZ_GRAFU
    Else
        souhrn = "Charts laid out: " & nUpr & ". Unchanged, adjust by hand: " & nBez & ". Skipped: " & nPres & "."
        If mVetaVelikosti <> "" Then souhrn = souhrn & "</p>" & vbLf & "<p class='souhrn'>" & Html(mVetaVelikosti)
    End If
    If radky <> "" Then
        tabulka = "<table><thead><tr><th>Slide</th><th>Result</th><th>Labels</th><th>Defects before</th><th>Defects after</th>" & _
            "<th>Note</th></tr></thead><tbody>" & radky & "</tbody></table>"
    End If
    If styl = STYL_KRATSI Then volby = "shorter lines" Else volby = "aligned columns"
    volby = "Placement: " & volby & ". Slides: "
    If snimky = "" Then volby = volby & "all." Else volby = volby & PopisSnimku(snimky) & "."
    If vystup = "" Then cil = "Result: the open presentation was edited." Else cil = "Output: " & vystup
    If zaloha <> "" Then cil = cil & " " & ZALOHA_TEXT & zaloha
    ZapisUtf8 cesta, "<!doctype html>" & vbLf & "<html lang='en'><head><meta charset='utf-8'>" & _
        "<meta name='viewport' content='width=device-width, initial-scale=1'><title>Label layout check</title><style>" & Css() & _
        "</style></head>" & vbLf & "<body><main>" & vbLf & "<h1>Label layout: " & Html(vstup) & "</h1>" & vbLf & _
        "<p class='meta'>" & Html(cil) & "<br>" & Html(volby) & "<br>Add-in version " & VERZE & ", " & _
        Format$(Now, "yyyy-mm-dd hh:nn") & "</p>" & vbLf & _
        "<p class='souhrn'>" & souhrn & "</p>" & vbLf & tabulka & vbLf & "</main></body></html>" & vbLf
End Sub

Private Function Css() As String
    Css = ":root { color-scheme: light; }" & _
        "body { font-family: 'Segoe UI', system-ui, sans-serif; margin: 0; padding: 24px 16px 48px; color: #1f2328; background: #fff; }" & _
        "main { max-width: 1400px; margin: 0 auto; }" & _
        "h1 { font-size: 22px; margin: 0 0 4px; }" & _
        "p.meta { color: #59636e; margin: 0 0 20px; font-size: 14px; }" & _
        "p.souhrn { font-size: 16px; margin: 0 0 16px; }" & _
        "table { border-collapse: collapse; width: 100%; font-size: 14px; margin-bottom: 32px; }" & _
        "th, td { text-align: left; padding: 6px 10px; border-bottom: 1px solid #d1d9e0; vertical-align: top; }" & _
        "th { font-weight: 600; color: #59636e; }" & _
        "td.cislo { font-variant-numeric: tabular-nums; white-space: nowrap; }"
End Function

Private Function VadyText(v As Vady, ByVal ma As Boolean) As String
    Dim s As String
    If Not ma Then
        VadyText = "–"
        Exit Function
    End If
    s = PridejVadu(s, "label overlaps", v.PrekryvPopisku)
    s = PridejVadu(s, "label over point", v.PopisekPresBod)
    s = PridejVadu(s, "crossing lines", v.KrizeniCar)
    s = PridejVadu(s, "line over label", v.CaraPresPopisek)
    s = PridejVadu(s, "label over line", v.PopisekPresCaru)
    s = PridejVadu(s, "label over box", v.PopisekPresBox)
    s = PridejVadu(s, "outside chart", v.MimoGraf)
    If s = "" Then s = "no defects"
    VadyText = s
End Function

Private Function PridejVadu(ByVal s As String, ByVal nazev As String, ByVal pocet As Long) As String
    PridejVadu = s
    If pocet = 0 Then Exit Function
    If s <> "" Then PridejVadu = s & ", "
    PridejVadu = PridejVadu & nazev & " " & pocet
End Function

Private Function Html(ByVal s As String) As String
    Html = Replace(Replace(Replace(Replace(s, "&", "&amp;"), "<", "&lt;"), ">", "&gt;"), "'", "&#39;")
End Function

Private Sub ZapisUtf8(ByVal cesta As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2  ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.SaveToFile cesta, 2  ' adSaveCreateOverWrite
    st.Close
End Sub

' ---------- slide numbers ----------
' Slide numbers as typed in the dialog ("2, 5-7": separated by commas, semicolons or spaces, ranges with a hyphen
' or an en dash) -> ascending list without duplicates ("2,5,6,7"); chyba = message for the user, "" = OK.
Public Function RozsahSnimku(ByVal spec As String, ByVal pocet As Long, ByRef chyba As String) As String
    Dim ano() As Boolean, cast As Variant, i As Long, a As Long, b As Long, p As Long, s As String
    chyba = ""
    ReDim ano(0 To pocet)
    spec = Replace(Replace(Replace(spec, ";", ","), " ", ","), ChrW(8211), "-")
    For Each cast In Split(spec, ",")
        cast = Trim$(cast)
        If cast <> "" Then
            p = InStr(2, cast, "-")
            If p > 0 Then
                If Not CeleCislo(Left$(cast, p - 1), a) Or Not CeleCislo(Mid$(cast, p + 1), b) Then GoTo Spatne
            Else
                If Not CeleCislo(cast, a) Then GoTo Spatne
                b = a
            End If
            If a > b Then
                p = a
                a = b
                b = p
            End If
            If a < 1 Then
                chyba = "There is no slide " & a & ", numbering starts at 1."
                Exit Function
            End If
            If b > pocet Then
                chyba = "There is no slide " & b & " (the last one is " & pocet & ")."
                Exit Function
            End If
            For i = a To b
                ano(i) = True
            Next i
        End If
    Next cast
    For i = 1 To pocet
        If ano(i) Then s = s & "," & i
    Next i
    If s = "" Then
        chyba = "Enter slide numbers, for example 2, 5-7."
        Exit Function
    End If
    RozsahSnimku = Mid$(s, 2)
    Exit Function
Spatne:
    chyba = """" & cast & """ is not a slide number. Enter slide numbers, for example 2, 5-7."
End Function

Private Function CeleCislo(ByVal s As String, ByRef n As Long) As Boolean
    Dim i As Long
    s = Trim$(s)
    If s = "" Or Len(s) > 5 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    n = CLng(s)
    CeleCislo = True
End Function

' For tests: the normalized list, or "!" and the message.
Public Function OverRozsahSnimku(ByVal spec As String, ByVal pocet As Long) As String
    Dim chyba As String, s As String
    s = RozsahSnimku(spec, pocet, chyba)
    If chyba <> "" Then OverRozsahSnimku = "!" & chyba Else OverRozsahSnimku = s
End Function

' "2,5,6,7" -> "2, 5–7" (for the dialog and the report).
Public Function PopisSnimku(ByVal seznam As String) As String
    Dim c As Variant, a As Long, b As Long, n As Long, s As String, prvni As Boolean
    prvni = True
    For Each c In Split(seznam, ",")
        n = CLng(c)
        If prvni Then
            a = n
            b = n
            prvni = False
        ElseIf n = b + 1 Then
            b = n
        Else
            s = PridejUsek(s, a, b)
            a = n
            b = n
        End If
    Next c
    If Not prvni Then s = PridejUsek(s, a, b)
    PopisSnimku = s
End Function

Private Function PridejUsek(ByVal s As String, ByVal a As Long, ByVal b As Long) As String
    If s <> "" Then s = s & ", "
    If a = b Then PridejUsek = s & a Else PridejUsek = s & a & "–" & b
End Function

' ---------- files and messages ----------
Private Function ZakladNazvu(nazev As String) As String
    Dim z As String
    z = nazev
    If InStrRev(z, ".") > 0 Then z = Left$(z, InStrRev(z, ".") - 1)
    ZakladNazvu = BezPripony(BezPripony(z, PRIPONA), PRIPONA_STARA)
End Function

' Name without the output suffix (also a numbered one such as -labels-2), so that outputs do not chain.
Private Function BezPripony(ByVal z As String, ByVal konc As String) As String
    Dim p As Long
    BezPripony = z
    p = InStrRev(LCase$(z), konc)
    If p > 0 Then
        If IsNumeric(Mid$(z, p + Len(konc) + 1)) Or p + Len(konc) - 1 = Len(z) Then BezPripony = Left$(z, p - 1)
    End If
End Function

' Work folder in TEMP: outputs of presentations that have no local folder (unsaved, OneDrive, SharePoint).
Private Function PracovniSlozka() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    PracovniSlozka = Environ$("TEMP") & "\" & NAZEV_DOPLNKU
    If Not fso.FolderExists(PracovniSlozka) Then fso.CreateFolder PracovniSlozka
End Function

Private Function SlozkaVystupu(pres As Presentation, prac As String) As String
    SlozkaVystupu = pres.Path
    If SlozkaVystupu = "" Or LCase$(Left$(SlozkaVystupu, 4)) = "http" Then SlozkaVystupu = prac
End Function

Private Function CilovaCesta(pres As Presentation, zaklad As String, prac As String) As String
    Dim slozka As String
    slozka = SlozkaVystupu(pres, prac)
    Dim n As Long, kandidat As String
    For n = 1 To 99
        kandidat = slozka & "\" & zaklad & PRIPONA
        If n > 1 Then kandidat = kandidat & "-" & n
        kandidat = kandidat & ".pptx"
        If Not JeOtevrena(kandidat) Then
            CilovaCesta = kandidat
            Exit Function
        End If
    Next n
    CilovaCesta = prac & "\" & zaklad & PRIPONA & "-" & Format$(Now, "hhnnss") & ".pptx"
End Function

Private Function JeOtevrena(cesta As String) As Boolean
    Dim p As Presentation
    For Each p In Presentations
        If StrComp(p.FullName, cesta, vbTextCompare) = 0 Then
            JeOtevrena = True
            Exit Function
        End If
    Next p
End Function

Private Sub OtevriSoubor(cesta As String)
    If CreateObject("Scripting.FileSystemObject").FileExists(cesta) Then
        CreateObject("Shell.Application").ShellExecute cesta
    End If
End Sub

Public Sub Zprava(ByVal text As String, ByVal styl As VbMsgBoxStyle)
    Dim nadpis As String
    nadpis = TITULEK
    MessageBoxW 0, StrPtr(text), StrPtr(nadpis), styl Or MB_SETFOREGROUND
End Sub
