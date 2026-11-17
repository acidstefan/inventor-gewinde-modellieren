# Changelog

## 2.1.0 – 2026-11-17

- Angeordnete Gewinde: rechteckige und runde Anordnungen werden für die modellierten Gewinde nachgebildet und mit den Original-Parametern verknüpft (Anzahl, Abstand, Winkel); unterdrückte Elemente werden übernommen
- Gespiegelte Gewinde werden direkt als Rechtsgewinde modelliert
- Auswahl einer Anordnung modelliert deren Gewinde
- Lässt sich eine Anordnung nicht verknüpft nachbilden, werden ihre Gewinde einzeln modelliert
- Rohrgewinde: Inventor-Bezeichnung „G47.803x2.309“ (Ø x Steigung in mm) wird erkannt
- Zuordnung Spirale → Bohrung neu gespeichert (Attributsatz-Name), da Inventor 2016 den Attributwert nicht lesbar zurückgibt; Spiralen aus 2.0 bitte neu erzeugen
- Ausführlichere Diagnose im Direktfenster (Gewindeinfo, Anordnungen)

## 2.0.0 – 2026-11-13

- ISO-Rohrgewinde ergänzt: G/Rp (ISO 228-1) und kegelig R/Rc (ISO 7-1)
- Steigung bei Feingewinden robuster ermittelt
- Diagnosezeile im Direktfenster, Konstante `SHOW_DEBUG`

## 1.0.0 – 2026-11-12

- Erste Version: metrische ISO-Gewinde aus Gewinde-Features und Gewindebohrungen als Spirale modellieren
