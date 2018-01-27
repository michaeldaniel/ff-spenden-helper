# ff-spenden-helper
## Eine Anwendung für Feuerwehren, zur Einmeldung von Spenden

### Einleitung
Freiwillige Feurwehren müssen ihre Spenden bei Finanzonline einmelden.
Dieses Tool soll dabei unterstützen. Es wird kostenlos zur Verfügung gestellt und kann gerne verändert/erweitert werden.

### Systemvoraussetzungen
Dieses Tool wird unter Windows 7 und Windows 10 (empfohlen) unterstützt. Windows Powershell in der akutellen Version für das eingesetzte Betriebssystem wird benötigt (ist normalerweise automatisch Teil von Windows 7/10).

### Anleitung
Die Einmeldung der Spenden verläuft in wenigen Schritten

1. Eintragung aller Spenden in ein Excel-Dokument (ein Beispieldokument liegt bei).
  * Eintragung des SZR-Headers in die Excel-Datei (zu finden bei Finanzonline). Dies ist nur 1x notwendig.
  
2. Aufruf des ff-spenden-helper Tools
  * Auswahl der Excel-Datei
  * Auswahl des Jahres (so muss die Arbeitsmappe im Excel-Dokument für das betreffende Jahr genannt werden)
  * Angabe der Zeile, wo die eigentliche Tabelle startet (inkl. Überschriften)
  * Angabe des Arbeitsmappennamens, wo die SZR-Informationen zu finden sind.
  
3. Eine SZR-Upload Datei generieren (Klick auf "Stammzahlenregister Upload-Datei generieren").

4. Mittels SZR-Upload Datei im Stammzahlenregister die SZR-Informationen anfordern (kann einen Tag dauern) und downloaden (zip-Archiv).

5. Die Excel-Datei erneut im Spendentool angeben (inkl. aller Informationen von Schritt 2.)
  * Zusätzlich das zuvor downgeloadete ZIP-Archiv bei Schritt 3 im Tool angeben.
  
6. Finanzonline XML-Datei generieren und bei Finanzonline Hochladen.

7. Die Spenden wurden eingemeldet.

### Hinweise zur Datenverarbeitung
Dieses Tool arbeitet nur lokal am eigenen PC. Die Spenderdaten werden im Zuge der Verarbeitung durch dieses Tool an niemanden weitergegeben oder eingemeldet. Die Einmeldung findet erst durch die manuelle Übermittlung der XML-Datei statt. Somit hat der Benutzer volle Kontrolle über die Spenderdaten.

### Haftungsausschluss
Die Benutzung des Tools erfolgt auf eigene Gefahr. Es wird keinerlei Haftung übernommen.

### Verbesserungswünsche/Beschwerden/Fehler ;)
Bitte an n08107082@feuerwehr.gv.at senden.
