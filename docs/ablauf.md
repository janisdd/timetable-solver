# Ablauf

- Datei `SJ 25-26_Gesamtübersicht Einsatz Lehrkräfte EA Halle 2025-05-21_3.xlsx` wird eingelesen
  - wird eingelesen, wobei alle Tabellen mit `variablen` oder `kumuliert` im Namen ignoriert werden
  - daraus werden die Lehrer und Klassen extrahiert
    - bei den Klassen noch die Fächer (Subjects) mit den Stunden für die Lehrer (`teachers_with_hours`)


- Datei `Mappings.xlsx` wird eingelesen
  - die Tabelle `farben` gibt 10 farben für die Pläne der Lehrere an
    - 10 Farben (grün) `kann nur da` gibt an, dass der Lehrere nur da Zeit hat
     10 Farben (rot) `kann da nicht` gibt an, dass der Lehrere da keine Zeit hat
  - die Tabelle `klassen` gibt die Kürzel für die lange Bezeichnung der Klasse an
    - eine Klasse kann bis zu drei Zeilen in der Gesamtübersicht haben
  - die Tabelle `fächer` gibt an das Kürzel der Fächer an
    - die exakten Namen aus der Gesamtübersicht werden genutzt

- dann wird `03_KW 10_02.03.-06.03.2026_gesetzteverpflichtende Stunden drin ohne Reste aus Vorwoche_IN.xlsm` gelesen



## TODOs

- in mappings, füge klassen und fächer hinzu (falls tabelle schon da), welche noch nicht da sind!