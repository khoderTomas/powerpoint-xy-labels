"""Builds RozmistiPopisky.ppam: VBA modules (MODULY), the dialog (FORMULARE) and the ribbon button on the Home tab.

Needs PowerPoint (builds through COM). Inserting macro code programmatically works only with the option
"Trust access to the VBA project object model" (HKCU AccessVBOM); the script turns it on only for the build
and then restores it. PowerPoint must be closed (it reads the option at startup).
The VBA project is compiled before saving; a compile error stops the build with the module and line.
The dialog is an empty UserForm with the code from <name>.vba next to this script; it creates its controls
at run time.

Usage:  uv run <this script> [path.ppam]      (default dist/RozmistiPopisky.ppam)
"""
import contextlib
import os
import re
import subprocess
import sys
import time
import zipfile
import winreg

import pythoncom
import win32com.client
import win32con
import win32gui
import win32process

ZDE = os.path.dirname(os.path.abspath(__file__))
KOREN = os.path.dirname(ZDE)
KLIC = r'Software\Microsoft\Office\16.0\PowerPoint\Security'
PP_ADDIN = 30  # ppSaveAsOpenXMLAddin
OFFICE_GUID = '{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}'
IKONY = ['ChartDataLabel', 'ChartDataLabelsGallery', 'ChartTypeXYScatterInsertGallery', 'ChartInsert', 'ChartTypeOtherInsertGallery']
REL_UI = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
MODULY = ['Graf', 'Rozlozeni', 'Blizko', 'Pisma', 'Popisky']  # <module>.bas next to this script, module name = file name
FORMULARE = ['Nastaveni']  # <form>.vba next to this script: code of the UserForm


def vba_ascii(kod):
    """Strings with non-ASCII characters -> "..." & ChrW(n) & "...", the VBA source is then plain ASCII (independent of the code page)."""
    def literal(m):
        obsah = m.group(0)[1:-1]
        if all(ord(c) < 128 for c in obsah):
            return m.group(0)
        casti, buf = [], ''
        for c in obsah:
            if ord(c) < 128:
                buf += c
            else:
                if buf:
                    casti.append('"%s"' % buf)
                    buf = ''
                casti.append('ChrW(%d)' % ord(c))
        if buf:
            casti.append('"%s"' % buf)
        return ' & '.join(casti)
    out = re.sub(r'"(?:[^"\n]|"")*"', literal, kod)
    zbytek = sorted({c for c in out if ord(c) > 127})
    if zbytek:
        raise ValueError('non-ASCII characters outside strings (VBA comments must be ASCII): %s' % ''.join(zbytek))
    return out


@contextlib.contextmanager
def pristup_k_vba():
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, KLIC) as k:
        try:
            puvodni = winreg.QueryValueEx(k, 'AccessVBOM')
        except FileNotFoundError:
            puvodni = None
        winreg.SetValueEx(k, 'AccessVBOM', 0, winreg.REG_DWORD, 1)
    try:
        yield
    finally:
        # PowerPoint writes the security settings back when it exits; restore only after it ends (also after an error)
        pythoncom.CoFreeUnusedLibraries()
        for _ in range(120):
            if not _powerpoint_bezi():
                break
            time.sleep(0.5)
        else:
            print('WARNING: PowerPoint did not exit, check AccessVBOM manually.')
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, KLIC) as k:
            if puvodni is None:
                winreg.DeleteValue(k, 'AccessVBOM')
            else:
                winreg.SetValueEx(k, 'AccessVBOM', 0, puvodni[1], puvodni[0])
        time.sleep(1.0)
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, KLIC) as k:
            try:
                ted = winreg.QueryValueEx(k, 'AccessVBOM')[0]
            except FileNotFoundError:
                ted = None
        if ted != (None if puvodni is None else puvodni[0]):
            raise RuntimeError('AccessVBOM could not be restored (now %s), fix it in the Trust Center' % ted)
        print('AccessVBOM restored and verified (%s)' % ('not set' if ted is None else ted))


def pridej_ribbon(ppam, xml):
    with zipfile.ZipFile(ppam) as z:
        polozky = [(i, z.read(i.filename)) for i in z.infolist()]
    vystup = []
    for info, data in polozky:
        if info.filename == '_rels/.rels':
            text = data.decode('utf-8')
            if REL_UI not in text:
                text = text.replace('</Relationships>', '<Relationship Id="rIdPopiskyUI" Type="%s" Target="customUI/customUI14.xml"/></Relationships>' % REL_UI)
            data = text.encode('utf-8')
        elif info.filename == '[Content_Types].xml':
            text = data.decode('utf-8')
            if 'Extension="xml"' not in text:
                text = text.replace('<Default ', '<Default Extension="xml" ContentType="application/xml"/><Default ', 1)
            data = text.encode('utf-8')
        vystup.append((info, data))
    with zipfile.ZipFile(ppam, 'w', zipfile.ZIP_DEFLATED) as z:
        for info, data in vystup:
            z.writestr(info, data)
        z.writestr('customUI/customUI14.xml', xml.encode('utf-8'))


# GetImageMso through COM from another process does not work (the image cannot be passed), checked by a macro inside PowerPoint
KONTROLA_IKON = '''Public Function IkonaExistuje(nazev As String) As Boolean
    On Error Resume Next
    Dim obrazek As Object
    Set obrazek = Application.CommandBars.GetImageMso(nazev, 32, 32)
    IkonaExistuje = (Err.Number = 0) And Not (obrazek Is Nothing)
End Function
'''


def _novy_modul(projekt, nazev, kod, typ=1):
    modul = projekt.VBComponents.Add(typ)  # 1 standard module, 3 UserForm (also adds the MSForms reference)
    modul.Name = nazev
    if modul.CodeModule.CountOfLines:  # "Option Explicit" from the editor settings would be there twice
        modul.CodeModule.DeleteLines(1, modul.CodeModule.CountOfLines)
    modul.CodeModule.AddFromString(kod)
    return modul


def _pid_powerpointu():
    vypis = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq POWERPNT.EXE', '/FO', 'CSV', '/NH'],
                           capture_output=True, text=True).stdout
    return {int(r.split('","')[1]) for r in vypis.splitlines() if r.upper().startswith('"POWERPNT.EXE')}


def _dialogy(pidy):
    """Dialog windows of the processes (class #32770) and their texts."""
    out = []

    def okno(hwnd, _):
        if win32gui.GetClassName(hwnd) == '#32770' and win32process.GetWindowThreadProcessId(hwnd)[1] in pidy:
            texty = []
            win32gui.EnumChildWindows(hwnd, lambda h, _: texty.append(win32gui.GetWindowText(h)) or True, None)
            out.append((hwnd, ' | '.join(t for t in texty if t.strip())))
        return True
    win32gui.EnumWindows(okno, None)
    return out


def _preloz(app, projekt):
    """Compiles the VBA project (Debug > Compile VBAProject). Compiling runs asynchronously and an error opens a
    modal dialog that would block PowerPoint and the tests: waits for a dialog (error) or for the command to be
    greyed out (compiled). The dialog is read and closed, then the place of the error in the code is reported."""
    hlasky, pidy = [], _pid_powerpointu()
    app.VBE.ActiveVBProject = projekt  # otherwise e.g. the installed add-in would be compiled
    prikaz = app.VBE.CommandBars.FindControl(1, 578)  # msoControlButton, Compile VBAProject
    if prikaz is None:
        raise RuntimeError('the VBA editor has no Compile command (578)')
    if prikaz.Enabled:
        prikaz.Execute()
    for _ in range(120):
        time.sleep(0.5)
        dialogy = _dialogy(pidy)
        if dialogy:
            for hwnd, text in dialogy:
                hlasky.append(text)
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            time.sleep(1.0)
            break
        try:
            if not app.VBE.CommandBars.FindControl(1, 578).Enabled:
                return
        except pythoncom.com_error:  # PowerPoint is not accepting calls right now (compiling or showing a dialog)
            pass
    else:
        raise RuntimeError('VBA compilation did not finish within 60 s')
    if hlasky:
        pane = app.VBE.ActiveCodePane
        radek = pane.GetSelection()[0]
        modul = pane.CodeModule
        raise RuntimeError('VBA compile error in %s.bas, line %d: %s\n    %s' % (
            modul.Name, radek, '; '.join(hlasky), modul.Lines(radek, 1).strip()))


def _sestav(cil, moduly, formulare):
    """Saves the add-in under a temporary name: an add-in with the same file name (installed from dist) loads
    also when PowerPoint starts through COM, and PowerPoint would refuse to save a second one. main() renames it."""
    app = win32com.client.Dispatch('PowerPoint.Application')
    ikona, chyba, pres, projekt, kontrola = None, None, None, None, None
    try:
        pres = app.Presentations.Add(False)
        projekt = pres.VBProject
        if not any(r.Guid.upper() == OFFICE_GUID for r in projekt.References):
            projekt.References.AddFromGuid(OFFICE_GUID, 2, 0)
        kontrola = _novy_modul(projekt, 'KontrolaIkon', KONTROLA_IKON)
        for kandidat in IKONY:
            if app.Run('%s!KontrolaIkon.IkonaExistuje' % pres.Name, kandidat):
                ikona = kandidat
                break
        projekt.VBComponents.Remove(kontrola)
        for nazev, kod in moduly:
            _novy_modul(projekt, nazev, kod)
        for nazev, kod in formulare:
            _novy_modul(projekt, nazev, kod, 3)
        _preloz(app, projekt)
        pres.SaveAs(cil, PP_ADDIN)
        print('VBA project references: %s' % ', '.join(r.Name for r in projekt.References))
    except Exception as e:  # references held by the exception would keep PowerPoint alive and it would write AccessVBOM back
        chyba = '%s: %s' % (type(e).__name__, e)
    finally:
        try:
            if pres is not None:
                pres.Close()
            if app.Presentations.Count == 0:
                app.Quit()
        except Exception as e:
            chyba = chyba or 'PowerPoint could not be closed: %s' % e
    app = pres = projekt = kontrola = None
    if chyba:
        raise RuntimeError(chyba)
    return ikona


def _powerpoint_bezi():
    vypis = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq POWERPNT.EXE', '/NH'], capture_output=True, text=True).stdout
    return 'POWERPNT.EXE' in vypis.upper()


def main(cil):
    moduly, formulare = [], []
    for nazev in MODULY:
        with open(os.path.join(ZDE, nazev + '.bas'), encoding='utf-8') as f:
            moduly.append((nazev, vba_ascii(f.read())))
    for nazev in FORMULARE:
        with open(os.path.join(ZDE, nazev + '.vba'), encoding='utf-8') as f:
            formulare.append((nazev, vba_ascii(f.read())))
    for _ in range(60):  # after the tests PowerPoint may still be exiting
        if not _powerpoint_bezi():
            break
        time.sleep(0.5)
    else:
        sys.exit('PowerPoint is running: close it and run the build again (it reads the VBA access option at startup).')
    docasny = os.path.join(os.path.dirname(cil), '_sestavuji-%d.ppam' % os.getpid())
    pythoncom.CoInitialize()
    with pristup_k_vba():  # on exit waits for PowerPoint to end and restores AccessVBOM
        ikona = _sestav(docasny, moduly, formulare)
    if ikona is None:
        print('WARNING: none of the icons %s exists, the button will have no icon' % IKONY)
    with open(os.path.join(ZDE, 'customUI14.xml'), encoding='utf-8') as f:
        xml = f.read().replace('{IKONA}', ikona or 'ChartInsert')
    pridej_ribbon(docasny, xml)
    os.replace(docasny, cil)
    print('Built: %s (button icon %s)' % (cil, ikona))


if __name__ == '__main__':
    main(os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.join(KOREN, 'dist', 'RozmistiPopisky.ppam'))
