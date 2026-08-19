Attribute VB_Name = "modBExAktualisierung"
Option Explicit

'====================================================================
'  Automatische Aktualisierung von SAP BEx Analyzer Reports
'--------------------------------------------------------------------
'  Traegt man in der Tabelle "Steuerung" ab Zeile 5 die vollstaendigen
'  Pfade zu den BEx-Dateien ein, oeffnet dieses Makro jede Datei nach-
'  einander, aktualisiert sie (entspricht Rechtsklick -> "Aktualisieren"),
'  speichert und schliesst sie wieder.
'
'  Voraussetzung: SAP BEx Analyzer ist installiert und man ist am SAP-
'  System angemeldet (bzw. wird beim ersten Refresh zur Anmeldung
'  aufgefordert).
'====================================================================

Private Const START_ZEILE As Long = 5
Private Const SP_AKTIV    As String = "A"
Private Const SP_PFAD     As String = "B"
Private Const SP_STATUS   As String = "C"
Private Const SP_ZEIT     As String = "D"
Private Const SP_MELD     As String = "E"
Private Const TITEL       As String = "BEx Reports aktualisieren"

'--- Wird ueber den Button bzw. Alt+F8 gestartet -----------------------
Public Sub AktualisiereAlleReports()
    Dim ws As Worksheet
    Dim r As Long, letzte As Long
    Dim pfad As String, aktiv As String, meldung As String, hinweis As String
    Dim anzOK As Long, anzFehler As Long, anzUeber As Long

    Set ws = ThisWorkbook.Worksheets("Steuerung")
    letzte = ws.Cells(ws.Rows.Count, SP_PFAD).End(xlUp).Row

    If letzte < START_ZEILE Then
        MsgBox "Es sind keine Report-Pfade eingetragen." & vbCrLf & _
               "Bitte ab Zeile " & START_ZEILE & " in Spalte B die Pfade eintragen.", _
               vbExclamation, TITEL
        Exit Sub
    End If

    If MsgBox("Es werden jetzt alle aktiven Reports nacheinander geoeffnet, " & _
              "aktualisiert und wieder geschlossen." & vbCrLf & vbCrLf & _
              "Bitte waehrenddessen nicht in Excel arbeiten. Fortfahren?", _
              vbQuestion + vbOKCancel, TITEL) <> vbOK Then Exit Sub

    ' BEx-Refresh benoetigt sichtbare Fenster -> ScreenUpdating aktiv lassen
    Application.DisplayAlerts = False
    Application.EnableEvents = True

    For r = START_ZEILE To letzte
        pfad = Trim$(CStr(ws.Cells(r, SP_PFAD).Value))
        aktiv = LCase$(Trim$(CStr(ws.Cells(r, SP_AKTIV).Value)))

        If Len(pfad) = 0 Then GoTo NaechsteZeile

        If aktiv = "nein" Or aktiv = "n" Or aktiv = "0" Or aktiv = "aus" Then
            SetzeStatus ws, r, "Uebersprungen", "deaktiviert (Spalte A)"
            anzUeber = anzUeber + 1
            GoTo NaechsteZeile
        End If

        ' Absicherung: nur echte Pfade verarbeiten (enthalten \).
        ' Verhindert, dass versehentlicher Text als Datei interpretiert wird.
        If InStr(pfad, "\") = 0 Then
            SetzeStatus ws, r, "Uebersprungen", "kein gueltiger Pfad"
            anzUeber = anzUeber + 1
            GoTo NaechsteZeile
        End If

        SetzeStatus ws, r, "laeuft ...", ""
        DoEvents

        hinweis = ""
        meldung = AktualisiereEineDatei(pfad, hinweis)

        If Len(meldung) = 0 Then
            SetzeStatus ws, r, "OK", hinweis
            ws.Cells(r, SP_ZEIT).Value = Now
            anzOK = anzOK + 1
        Else
            SetzeStatus ws, r, "FEHLER", meldung
            anzFehler = anzFehler + 1
        End If
        DoEvents
NaechsteZeile:
    Next r

    Application.DisplayAlerts = True

    MsgBox "Aktualisierung abgeschlossen." & vbCrLf & vbCrLf & _
           "Erfolgreich:  " & anzOK & vbCrLf & _
           "Fehler:       " & anzFehler & vbCrLf & _
           "Uebersprungen: " & anzUeber, _
           IIf(anzFehler = 0, vbInformation, vbExclamation), TITEL
End Sub

'--- Oeffnet, aktualisiert, speichert und schliesst eine Datei ---------
'    Rueckgabe: ""  = Erfolg (Datei wurde gespeichert)
'               sonst = harte Fehlermeldung (z. B. Datei/Save-Problem).
'    hinweis (ByRef): weicher Hinweis (Zeile gilt trotzdem als OK).
Private Function AktualisiereEineDatei(ByVal pfad As String, _
                                       ByRef hinweis As String) As String
    Dim wb As Workbook
    On Error GoTo Fehler

    If Dir$(pfad) = "" Then
        AktualisiereEineDatei = "Datei nicht gefunden"
        Exit Function
    End If

    Set wb = Workbooks.Open(Filename:=pfad, UpdateLinks:=False)
    wb.Activate
    DoEvents

    ' Aktualisieren (best effort). Ein Refresh-Problem darf das
    ' anschliessende SPEICHERN NICHT verhindern.
    hinweis = AktualisiereBEx()
    DoEvents

    ' Nach dem BEx-Refresh Excel-Einstellungen wieder sicherstellen -
    ' das Add-In setzt DisplayAlerts u. U. zurueck.
    Application.DisplayAlerts = False

    ' Speichern, danach schliessen. Da bereits gespeichert wurde, ist die
    ' Mappe nicht mehr "dirty" -> es erscheint kein Speichern-Dialog.
    wb.Save
    wb.Close SaveChanges:=False
    Set wb = Nothing

    AktualisiereEineDatei = ""
    Exit Function

Fehler:
    Dim m As String
    m = "Fehler " & Err.Number & ": " & Err.Description
    ' Beim harten Fehler NICHT ungewollt speichern - Datei nur schliessen.
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    AktualisiereEineDatei = m
End Function

'--- Ruft die Aktualisierung des BEx Analyzer Add-Ins auf --------------
'    Entspricht dem Rechtsklick in der Tabelle -> "Aktualisieren".
'    Rueckgabe: "" = Refresh (sehr wahrscheinlich) ausgefuehrt,
'               sonst = weicher Hinweis (Zeile bleibt OK).
'    WICHTIG: Diese Funktion loest KEINEN Laufzeitfehler aus, damit das
'    Speichern der Datei niemals abgebrochen wird.
'
'    BEx 7.x ist ein COM-Add-In und liefert nach erfolgreichem Refresh
'    oft trotzdem Err <> 0. Ein NICHT vorhandenes Makro erkennt man daran,
'    dass der Fehlertext den Makronamen ("SAPBEXrefresh") enthaelt.

Private Function AktualisiereBEx() As String
    Dim ziele As Variant, i As Long
    Dim eNr As Long, eDesc As String
    Dim gelaufen As Boolean

    ziele = Array("BExAnalyzer.xla!SAPBEXrefresh", _
                  "SAPBEXrefresh", _
                  "SAPBEX.xla!SAPBEXrefresh")
    gelaufen = False

    For i = LBound(ziele) To UBound(ziele)
        On Error Resume Next
        Err.Clear
        Application.Run ziele(i), True
        eNr = Err.Number
        eDesc = Err.Description
        On Error GoTo 0

        If eNr = 0 Then
            gelaufen = True
            Exit For
        End If

        ' Fehlertext nennt den Makronamen NICHT -> Makro wurde gefunden
        ' und ausgefuehrt, meldete aber etwas (typisch BEx 7.x).
        If InStr(1, eDesc, "SAPBEXrefresh", vbTextCompare) = 0 Then
            gelaufen = True
            Exit For
        End If
        ' sonst: unter diesem Namen nicht vorhanden -> naechsten Namen testen
    Next i

    If gelaufen Then
        AktualisiereBEx = ""
    ElseIf BExGeladen() Then
        ' Add-In ist geladen -> Aktualisierung als erfolgt werten.
        AktualisiereBEx = ""
    Else
        AktualisiereBEx = "Hinweis: BEx-Refresh evtl. nicht ausgefuehrt " & _
                          "(Add-In nicht ansprechbar). Bitte Diagnose ausfuehren."
    End If
End Function

'--- Prueft, ob der SAP BEx Analyzer geladen ist (COM- oder xla-Add-In) -
Private Function BExGeladen() As Boolean
    Dim c As Object, ad As Object, w As Workbook, s As String
    On Error Resume Next
    For Each c In Application.COMAddIns
        If c.Connect Then
            s = CStr(c.progID) & " " & CStr(c.Description)
            If InStr(1, s, "BEx", vbTextCompare) > 0 _
               Or InStr(1, s, "Business Explorer", vbTextCompare) > 0 _
               Or InStr(1, s, "SAP", vbTextCompare) > 0 Then
                BExGeladen = True: Exit Function
            End If
        End If
    Next
    For Each ad In Application.AddIns
        If ad.Installed Then
            If InStr(1, ad.Name, "BEx", vbTextCompare) > 0 _
               Or InStr(1, ad.Name, "SAPBEX", vbTextCompare) > 0 Then
                BExGeladen = True: Exit Function
            End If
        End If
    Next
    For Each w In Application.Workbooks
        If InStr(1, w.Name, "BExAnalyzer", vbTextCompare) > 0 _
           Or InStr(1, w.Name, "SAPBEX", vbTextCompare) > 0 Then
            BExGeladen = True: Exit Function
        End If
    Next
    On Error GoTo 0
End Function

'--- Diagnose: zeigt geladene Add-Ins (bei Bedarf ausfuehren) -----------
'    Alt+F8 -> ZeigeBExDiagnose. Bitte Ergebnis an den Ersteller senden,
'    falls der BEx-Refresh nicht funktioniert.
Public Sub ZeigeBExDiagnose()
    Dim c As Object, ad As Object, w As Workbook
    Dim t As String
    On Error Resume Next

    t = "=== COM-Add-Ins ===" & vbCrLf
    For Each c In Application.COMAddIns
        t = t & "- " & CStr(c.progID) & "  | verbunden=" & CStr(c.Connect) & _
                "  | " & CStr(c.Description) & vbCrLf
    Next

    t = t & vbCrLf & "=== Add-Ins (installiert) ===" & vbCrLf
    For Each ad In Application.AddIns
        If ad.Installed Then t = t & "- " & ad.Name & vbCrLf
    Next

    t = t & vbCrLf & "=== Offene Mappen/Add-Ins ===" & vbCrLf
    For Each w In Application.Workbooks
        t = t & "- " & w.Name & vbCrLf
    Next

    On Error GoTo 0
    MsgBox t, vbInformation, "BEx-Diagnose"
End Sub

Private Sub SetzeStatus(ByVal ws As Worksheet, ByVal r As Long, _
                        ByVal status As String, ByVal meldung As String)
    ws.Cells(r, SP_STATUS).Value = status
    ws.Cells(r, SP_MELD).Value = meldung
End Sub

