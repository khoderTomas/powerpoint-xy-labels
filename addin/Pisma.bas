Option Explicit

' Text measurement through GDI: font 1000 px per em, width from GetTextExtentPoint32W, line spacing from
' the head/hhea/OS/2 tables (GetFontData). Matches FreeType in the Python reference (popisky/pisma.py):
' Calibri to 0.0000 pt, Arial, Cambria and Segoe UI within 0.012 pt.
' Measuring in a temporary text box on the slide does NOT match chart data labels (different line spacing,
' widths off by 1 to 11 pt).

Private Type SIZEL
    cx As Long
    cy As Long
End Type

Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal nHeight As Long, ByVal nWidth As Long, _
    ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, _
    ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, _
    ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As LongPtr) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function GetTextExtentPoint32W Lib "gdi32" (ByVal hdc As LongPtr, ByVal lpString As LongPtr, _
    ByVal c As Long, ByRef lpSize As SIZEL) As Long
Private Declare PtrSafe Function GetTextFaceW Lib "gdi32" (ByVal hdc As LongPtr, ByVal nCount As Long, ByVal lpFaceName As LongPtr) As Long
Private Declare PtrSafe Function GetFontData Lib "gdi32" (ByVal hdc As LongPtr, ByVal dwTable As Long, ByVal dwOffset As Long, _
    ByRef lpvBuffer As Any, ByVal cbData As Long) As Long
Private Declare PtrSafe Function GetFontDataDelka Lib "gdi32" Alias "GetFontData" (ByVal hdc As LongPtr, ByVal dwTable As Long, _
    ByVal dwOffset As Long, ByVal lpvBuffer As LongPtr, ByVal cbData As Long) As Long

Private Const EM As Long = 1000
Private Const NAHRADNI As String = "Calibri"
Private Const RADEK_VYCHOZI As Double = 1.2207

Private mDC As LongPtr
Private mPocet As Long
Private mKlic() As String
Private mNazev() As String      ' font name from the presentation (for the substitute font notice)
Private mFont() As LongPtr
Private mRadek() As Double
Private mNahradni() As Boolean

' Font index in the cache; a missing font is measured with the substitute (Calibri) and JeNahradni reports it.
Public Function PismoIndex(ByVal nazev As String, ByVal tucne As Boolean) As Long
    Dim klic As String, i As Long
    klic = LCase$(nazev) & IIf(tucne, "|1", "|0")
    For i = 1 To mPocet
        If mKlic(i) = klic Then
            PismoIndex = i
            Exit Function
        End If
    Next i
    If mDC = 0 Then mDC = CreateCompatibleDC(0)
    Dim f As LongPtr, nahradni As Boolean
    f = Vytvor(nazev, tucne)
    If StrComp(TvarPisma(), nazev, vbTextCompare) <> 0 Then
        DeleteObject f
        f = Vytvor(NAHRADNI, tucne)
        nahradni = StrComp(nazev, NAHRADNI, vbTextCompare) <> 0
    End If
    mPocet = mPocet + 1
    ReDim Preserve mKlic(1 To mPocet)
    ReDim Preserve mNazev(1 To mPocet)
    ReDim Preserve mFont(1 To mPocet)
    ReDim Preserve mRadek(1 To mPocet)
    ReDim Preserve mNahradni(1 To mPocet)
    mKlic(mPocet) = klic
    mNazev(mPocet) = nazev
    mFont(mPocet) = f
    mNahradni(mPocet) = nahradni
    mRadek(mPocet) = Radkovani()
    PismoIndex = mPocet
End Function

' Text width in pt (no wrapping).
Public Function SirkaTextu(ByVal idx As Long, ByVal text As String, ByVal pt As Double) As Double
    Dim sz As SIZEL
    If Len(text) = 0 Then Exit Function
    SelectObject mDC, mFont(idx)
    GetTextExtentPoint32W mDC, StrPtr(text), Len(text), sz
    SirkaTextu = sz.cx * pt / EM
End Function

' Line spacing as in PowerPoint: max(hhea ascender - descender + lineGap, OS/2 winAscent + winDescent) / unitsPerEm.
Public Function RadekPisma(ByVal idx As Long) As Double
    RadekPisma = mRadek(idx)
End Function

Public Function JeNahradni(ByVal idx As Long) As Boolean
    JeNahradni = mNahradni(idx)
End Function

Public Function NazevPisma(ByVal idx As Long) As String
    NazevPisma = mNazev(idx)
End Function

Public Sub UvolniPisma()
    Dim i As Long
    If mDC <> 0 Then DeleteDC mDC  ' DC first: a font selected into a DC cannot be deleted
    For i = 1 To mPocet
        DeleteObject mFont(i)
    Next i
    mDC = 0
    mPocet = 0
End Sub

Private Function Vytvor(ByVal nazev As String, ByVal tucne As Boolean) As LongPtr
    Vytvor = CreateFontW(-EM, 0, 0, 0, IIf(tucne, 700, 400), 0, 0, 0, 1, 0, 0, 0, 0, StrPtr(nazev))
    SelectObject mDC, Vytvor
End Function

Private Function TvarPisma() As String
    Dim buf As String
    buf = String$(64, vbNullChar)
    GetTextFaceW mDC, 64, StrPtr(buf)
    TvarPisma = Left$(buf, InStr(buf, vbNullChar) - 1)
End Function

Private Function Radkovani() As Double
    Dim head() As Byte, hhea() As Byte, os2() As Byte, upm As Long, rad As Long, win As Long
    Radkovani = RADEK_VYCHOZI
    If Not Tabulka("head", head) Or Not Tabulka("hhea", hhea) Then Exit Function
    upm = U16(head, 18)
    If upm = 0 Then Exit Function
    rad = S16(hhea, 4) - S16(hhea, 6) + S16(hhea, 8)
    If Tabulka("OS/2", os2) Then
        If UBound(os2) >= 77 Then win = U16(os2, 74) + U16(os2, 76)
    End If
    If win > rad Then rad = win
    Radkovani = rad / upm
End Function

Private Function Tabulka(ByVal znacka As String, b() As Byte) As Boolean
    Dim t As Long, n As Long
    t = Asc(Mid$(znacka, 1, 1)) + Asc(Mid$(znacka, 2, 1)) * &H100& + Asc(Mid$(znacka, 3, 1)) * &H10000 + Asc(Mid$(znacka, 4, 1)) * &H1000000
    n = GetFontDataDelka(mDC, t, 0, 0, 0)
    If n <= 0 Then Exit Function  ' GDI_ERROR = -1
    ReDim b(0 To n - 1)
    Tabulka = (GetFontData(mDC, t, 0, b(0), n) = n)
End Function

Private Function U16(b() As Byte, ByVal o As Long) As Long
    U16 = CLng(b(o)) * 256 + b(o + 1)
End Function

Private Function S16(b() As Byte, ByVal o As Long) As Long
    S16 = U16(b, o)
    If S16 >= 32768 Then S16 = S16 - 65536
End Function
