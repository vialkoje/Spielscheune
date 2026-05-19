Attribute VB_Name = "AbrechnungsMakro"
Option Explicit

' Startet den Abrechnungslauf via start.sh (macOS).
' Speichert die Arbeitsmappe zuerst, damit Python die aktuellen Daten liest.
Sub StartAbrechnung()
    If MsgBox("Abrechnungslauf jetzt starten?" & Chr(10) & Chr(10) & _
              "Die Datei wird vorher gespeichert.", _
              vbYesNo + vbQuestion, "Spielscheune Sonnenstunden") = vbNo Then Exit Sub

    ThisWorkbook.Save

    Dim wbPath As String
    wbPath = ThisWorkbook.Path

    Dim ergebnis As String
    On Error GoTo FehlerBehandlung

    ' AppleScript führt start.sh synchron aus und gibt stdout zurück
    ergebnis = MacScript( _
        "do shell script ""cd '" & wbPath & "' && bash start.sh 2>&1""")

    MsgBox "Abrechnung abgeschlossen." & Chr(10) & Chr(10) & ergebnis, _
           vbInformation, "Spielscheune Sonnenstunden"
    Exit Sub

FehlerBehandlung:
    MsgBox "Fehler beim Ausführen:" & Chr(10) & Err.Description & Chr(10) & Chr(10) & _
           "Bitte start.sh im Terminal prüfen.", _
           vbCritical, "Spielscheune Sonnenstunden"
End Sub
