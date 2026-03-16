# Ablauf

- Datei `SJ 25-26_Gesamtübersicht Einsatz Lehrkräfte EA Halle 2025-05-21_3.xlsx` wird eingelesen
  - wird eingelesen, wobei alle Tabellen mit `variablen` oder `kumuliert` im Namen ignoriert werden
  - daraus werden die Lehrer und Klassen extrahiert
    - bei den Klassen noch die Fächer (Subjects) mit den Stunden für die Lehrer (`teachers_with_hours`)


- Datei `Mappings.xlsx` wird eingelesen
  - die Tabelle `farben` gibt 10 farben für die Pläne der Lehrer an
    - 10 Farben (grün) `kann nur da` gibt an, dass der Lehrer nur da Zeit hat
     10 Farben (rot) `kann da nicht` gibt an, dass der Lehrer da keine Zeit hat
  - die Tabelle `klassen` gibt die Kürzel für die lange Bezeichnung der Klasse an
    - eine Klasse kann bis zu drei Zeilen in der Gesamtübersicht haben
  - die Tabelle `fächer` gibt an das Kürzel der Fächer an
    - die exakten Namen aus der Gesamtübersicht werden genutzt

- dann wird `03_KW 10_02.03.-06.03.2026_gesetzteverpflichtende Stunden drin ohne Reste aus Vorwoche_IN.xlsm` gelesen
  - für jeden Lehrer wird das entsprechende Sheet gelesen (wenn ein Lehrer kein Sheet hat, wird er ignoriert und aus allen Plänen entfernt)
  - an das Lehrer-Objekt wird der Plan als `availability_preference_table` angehängt
  - dann wird für jeden slot in jedem Plan bestimmt, ob er gesetzt werden soll oder nicht (wenn schon was drinne steht, dann nicht)
    - wenn die Farbe `kann nur da` (grün) ist, dann wird der Slot gesetzt
    - wenn die Farbe `kann da nicht` (rot) ist, dann wird der Slot nicht gesetzt
  - dann wird der Gesamtplan (Sheet `Plan`) ausgelesen



## TODOs

- in mappings, füge klassen und fächer hinzu (falls tabelle schon da), welche noch nicht da sind!