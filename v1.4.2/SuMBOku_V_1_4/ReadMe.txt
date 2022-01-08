Attribute VB_Name = "ReadMe"
'Bisher (18.03.2006) Zeilen:
'40+733+28+19+9+490(OFD)+26+521(SFD)+107+219+593+65+164
'2003 ohne OFD SFD
'3014 mit
'
'erl.| in Version 1.2 ist folgendes noch zu tun:
'(   ) * Undo/Redo für einzelne Zellen
'( |/) * Öffnen/Speichern von Spieldateien Endung .smbk
'(   ) * Gesamtiteration, mit einem Klick das S. lösen
'( |/) * Zusammenführen der drei Klassen B,L,C in eine einzige
'( |/) * Anzeige der fehlenden bzw. mögliche Werte pro
'(   )   B,L,C und Zelle
'( |/)   - 9 Image-Komponenten links vom Spielfeld für Zeile
'( |/)   - 9 IK unterhalb erste Blockreihe für jede Spalte
'( |/)   - 9 IK rechts von jeder mittl Zeile f jeden Block
'( |/)   - bewegt sich Maus over, dann Tooltip anzeigen
'( |/)   - Doppelklick kleines EingabeFenster anzeigen,
'( |/) * Spielbarkeit erhöhen, mit Doppelklick kleines
'(   )   Eingabefester für fehlende, bzw mögliche Werte pro
'(   )   B,L,C und Zelle
'(   ) * komplettUmbau des Spielfeldes für 4*4 & 16*16 Felder
'
'
'L(i) = L(i-1)+W(i-1)-1
'für 2*2*2*2
'Width = Height
'Left = Top
'10  74  74   11  74  74  10
'0    9  82  155 165 238 311 321
'Schriftgröße 40 = 0,5454 * 74
'
'
'für 3*3*3*3
'Width = Height
'Left = Top
'9 33  33  33    9  33  33  33   9  33  33  33   9
'0  8  40  72  104 112 144 176 208 216 248 280 312 321
'Schriftgröße 18 = 0,5454 * 33
'
'
'für 4*4*4*4
'Width = Height
'Left = Top
'8 19  19  19  19   7  19   19  19  19   7  19  19  19  19   7  19  19  19  19   8
'0  7  25  43  61  79  85  103 121 139 157 163 181 199 217 235 241 259 277 295 313 321
'Schriftgröße 10 = 0,5454 * 19
'
