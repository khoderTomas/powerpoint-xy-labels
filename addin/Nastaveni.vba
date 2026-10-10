Option Explicit

' Dialog before placement: which slides, which placement (approved on 9 Oct 2026) and where the result goes
' (approved on 10 Oct 2026). The controls are created at run time; the build (doplnek/sestav_doplnek.py) only
' puts this code into an empty UserForm named Nastaveni.
' Czech text in string literals goes through ChrW at build time.
' Use: Priprav, Show (modal), read Ok, Styl, Snimky and Kam, Unload.

Public Ok As Boolean
Public Styl As String          ' STYL_SLOUPCE or STYL_KRATSI
Public Snimky As String        ' "" = all slides, otherwise a normalized list such as "2,5,6,7"
Public Kam As String           ' KAM_KOPIE or KAM_OTEVRENA

Private mPocet As Long
Private mVybrane As String
Private mPripraveno As Boolean

Private mVse As MSForms.OptionButton
Private mVybraneOpt As MSForms.OptionButton
Private WithEvents mCislaOpt As MSForms.OptionButton
Private WithEvents mCisla As MSForms.TextBox
Private mSloupce As MSForms.OptionButton
Private mKratsi As MSForms.OptionButton
Private mKopie As MSForms.OptionButton
Private mOtevrena As MSForms.OptionButton
Private WithEvents mOk As MSForms.CommandButton
Private WithEvents mZrusit As MSForms.CommandButton

' pocet = slides in the presentation, vybrane = slides selected in PowerPoint (normalized list, "" = none),
' zMiniatur = the selection is in the slide thumbnails, styl = the placement chosen last time,
' kamMinule = the result chosen last time.
Public Sub Priprav(ByVal pocet As Long, ByVal vybrane As String, ByVal zMiniatur As Boolean, ByVal styl As String, _
                   ByVal kamMinule As String)
    Dim popis As String, n As Long
    mPocet = pocet
    mVybrane = vybrane
    Me.Caption = "Rozmístit popisky"
    Me.Width = 296
    Me.Height = 300
    Nadpis "Které snímky", 12, 10
    Set mVse = Volba("Všechny snímky", "V", "Snimky", 18, 28, 260)
    Set mVybraneOpt = Volba("Vybrané snímky", "B", "Snimky", 18, 48, 260)
    Set mCislaOpt = Volba("Čísla snímků:", "L", "Snimky", 18, 68, 92)
    Set mCisla = Me.Controls.Add("Forms.TextBox.1")
    Umisti mCisla, 112, 67, 166, 18
    Nadpis "Rozmístění", 12, 98
    Set mSloupce = Volba("Zarovnané sloupce", "Z", "Styl", 18, 116, 260)
    Set mKratsi = Volba("Kratší čáry (zkušební)", "K", "Styl", 18, 136, 260)
    Nadpis "Výsledek", 12, 166
    Set mKopie = Volba("Nová prezentace vedle původní (název-popisky.pptx)", "N", "Kam", 18, 184, 260)
    Set mOtevrena = Volba("Upravit otevřenou prezentaci", "U", "Kam", 18, 204, 260)
    Set mOk = Me.Controls.Add("Forms.CommandButton.1")
    Umisti mOk, 124, 236, 74, 24
    mOk.Caption = "Rozmístit"
    mOk.Default = True
    Set mZrusit = Me.Controls.Add("Forms.CommandButton.1")
    Umisti mZrusit, 206, 236, 74, 24
    mZrusit.Caption = "Zrušit"
    mZrusit.Cancel = True

    If vybrane <> "" Then
        n = UBound(Split(vybrane, ",")) + 1
        popis = PopisSnimku(vybrane)
        If n = 1 Then
            mVybraneOpt.Caption = "Vybraný snímek: " & popis
        ElseIf Len(popis) <= 20 Then
            mVybraneOpt.Caption = "Vybrané snímky: " & popis
        Else
            mVybraneOpt.Caption = "Vybrané snímky (" & n & ")"
        End If
    Else
        mVybraneOpt.Enabled = False
    End If
    If zMiniatur And n >= 2 Then mVybraneOpt.Value = True Else mVse.Value = True
    If styl = STYL_KRATSI Then mKratsi.Value = True Else mSloupce.Value = True
    If kamMinule = KAM_OTEVRENA Then mOtevrena.Value = True Else mKopie.Value = True
    mPripraveno = True
End Sub

Private Sub Umisti(c As Object, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    c.Left = l
    c.Top = t
    c.Width = w
    c.Height = h
    c.Font.Name = "Segoe UI"
    c.Font.Size = 9
End Sub

Private Sub Nadpis(ByVal text As String, ByVal l As Single, ByVal t As Single)
    Dim c As MSForms.Label
    Set c = Me.Controls.Add("Forms.Label.1")
    Umisti c, l, t, 268, 15
    c.Caption = text
    c.Font.Bold = True
End Sub

Private Function Volba(ByVal text As String, ByVal klavesa As String, ByVal skupina As String, ByVal l As Single, _
                       ByVal t As Single, ByVal w As Single) As MSForms.OptionButton
    Dim c As MSForms.OptionButton
    Set c = Me.Controls.Add("Forms.OptionButton.1")
    Umisti c, l, t, w, 18
    c.Caption = text
    c.Accelerator = klavesa
    c.GroupName = skupina
    Set Volba = c
End Function

Private Sub mCisla_Change()
    If mPripraveno Then mCislaOpt.Value = True
End Sub

Private Sub mCislaOpt_Click()
    If Not mPripraveno Then Exit Sub
    On Error Resume Next  ' focus fails while the form is not shown yet
    mCisla.SetFocus
End Sub

Private Sub mOk_Click()
    Dim chyba As String, s As String
    If mCislaOpt.Value Then
        s = RozsahSnimku(mCisla.Text, mPocet, chyba)
        If chyba <> "" Then
            Zprava chyba, vbExclamation
            mCisla.SetFocus
            Exit Sub
        End If
        Snimky = s
    ElseIf mVybraneOpt.Value Then
        Snimky = mVybrane
    Else
        Snimky = ""
    End If
    If mKratsi.Value Then Styl = STYL_KRATSI Else Styl = STYL_SLOUPCE
    If mOtevrena.Value Then Kam = KAM_OTEVRENA Else Kam = KAM_KOPIE
    Ok = True
    Me.Hide
End Sub

Private Sub mZrusit_Click()
    Ok = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then  ' the X button: hide instead of unloading, the caller still reads Ok
        Cancel = True
        Ok = False
        Me.Hide
    End If
End Sub
