Attribute VB_Name = "modStart"
Option Explicit

'=====================================================================
'  Stoffverteilungsplan - was man von Hand starten darf
'  ------------------------------------------------------------------
'  DIES IST DAS EINZIGE MODUL OHNE "Option Private Module".
'
'  Damit ist es das einzige, dessen Makros in der Liste unter Alt+F8
'  auftauchen. Vorher standen dort 30 Eintraege - Layout-Abfragen,
'  FastOff, ResetKopfzeilen, EnsureRowButtons und aehnliches, alles
'  interne Bausteine, die man von Hand nie aufruft und mit denen man
'  im schlimmsten Fall etwas kaputtmacht. Gemeint waren immer nur die
'  vier hier.
'
'  Die uebrigen Module bleiben untereinander voll erreichbar:
'  Option Private Module sperrt nur den Zugriff von FREMDEN
'  VBA-Projekten und blendet die Makroliste aus - innerhalb dieser
'  Mappe aendert sich nichts. Auch die Schaltflaechen laufen weiter,
'  denn Shape.OnAction loest den Namen im eigenen Projekt auf. Der
'  Selbsttest weist das in Abschnitt 0 ausdruecklich nach, damit man
'  sich nicht auf eine Vermutung verlassen muss.
'
'  Alles, was man im Alltag braucht, sitzt ohnehin als Schaltflaeche
'  im Blatt "Steuerung". Die vier hier sind fuer den Fall, dass genau
'  dieses Blatt nicht mehr da oder kaputt ist.
'=====================================================================


'  Baut Steuerung, Anleitung, Formeln und Blattschutz neu auf.
'  Dasselbe wie die Schaltflaeche "Einrichtung / Reparatur" - hier
'  fuer den Fall, dass das Blatt "Steuerung" selbst fehlt.
Public Sub Einrichtung_und_Reparatur()
    modSteuerung.Einrichtung_Reparatur
End Sub


'  Schreibt das Blatt "Anleitung" komplett neu. Bewusst keine
'  Schaltflaeche: es ueberschreibt alles, was jemand dort von Hand
'  ergaenzt hat.
Public Sub Anleitungsblatt_neu_schreiben()
    modAnleitung.Anleitungsblaetter_Erzeugen
End Sub


'  Schreibt eine Textdatei mit dem erkannten Layout und dem Zustand
'  der Mappe. Die braucht man, wenn eine Fehlermeldung auftaucht,
'  die man nicht einordnen kann.
Public Sub Diagnose()
    modWochenplan.Diag_Wochenplan
End Sub


'  Prueft die Mappe in echtem Excel durch. Nur fuer die Entwicklung -
'  laeuft ausdruecklich nur in einer Kopie, deren Name "Test" enthaelt,
'  weil er die Mappe unterwegs veraendert.
Public Sub Selbsttest()
    modSelbsttest.Selbsttest
End Sub


'  Prueft den Selbsttest selbst: baut definierte Fehler ein und
'  verlangt, dass genau der dafuer zustaendige Check rot wird.
'  Ebenfalls nur in einer Kopie mit "Test" im Dateinamen.
Public Sub Selbsttest_pruefen()
    modSelbsttest.Selbsttest_Pruefen
End Sub


'  Nur fuer den Selbsttest, Abschnitt 0: die Gegenprobe zu der Sonde
'  in modWochenplan. Schlaegt dort etwas fehl, sagt diese hier, ob es
'  an Option Private Module lag oder ob Application.Run ueberhaupt
'  nicht laeuft.
'
'  Das Argument ist der Grund, warum sie NICHT in der Makroliste
'  steht: dort erscheinen nur Public Subs ganz OHNE Argumente. Es ist
'  optional, damit Application.Run sie ohne Parameter aufrufen kann.
Public Sub Probe_modStart(Optional ByVal ohneBedeutung As Long = 0)
End Sub
