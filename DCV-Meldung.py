# Input: Reporter-Abfrage 
#        "Namensliste aller Aktiven für DCV-Meldung (Datenbasis für Skript)"
#        in Zwischenablage, dann dieses Skript starten

import csv
import io
import win32clipboard
from collections import defaultdict

print("Bitte Reporter-Abfrage ")
print("'Namensliste aller Aktiven für DCV-Meldung (Datenbasis für Skript)'")
print("durchführen und Daten in Zwischenablage ablegen.")
input("Bitte ENTER drücken, wenn dies geschehen ist!")

win32clipboard.OpenClipboard()
data = win32clipboard.GetClipboardData()
win32clipboard.CloseClipboard()

if not data.startswith("lfd. Nr.\t"):
    print("Fehler: Unerwarteter Inhalt der Zwischenablage!")
    exit()

with io.StringIO(data) as infile:
    daten = list(csv.DictReader(infile, delimiter="\t"))


jugend = {"Weiblich": defaultdict(int),
          "Männlich": defaultdict(int),
          "Divers": defaultdict(int)
}
erwachsen = {"Weiblich": defaultdict(int),
          "Männlich": defaultdict(int),
          "Divers": defaultdict(int)
}

for eintrag in daten:
    if not eintrag["lfd. Nr."]:
        continue # ignoriere Mitgliedschaft in mehrere Chören
    if eintrag["Bereich"] == "Les Passerelles":
        erwachsen[eintrag["Geschlecht"]][int(eintrag["Alter am Stichtag"])] += 1
    else:
        jugend[eintrag["Geschlecht"]][int(eintrag["Alter am Stichtag"])] += 1

print("DCV-Abrechnungsdaten") 
print("--------------------")

DCV = {"bis 26 Jahre": (0,26),
       "ab 27 Jahre": (27,999)}

for kategorie in DCV:
    summe = 0
    for geschlecht in jugend:
        for alter in jugend[geschlecht]:
            if DCV[kategorie][0] <= alter <= DCV[kategorie][1]:
                summe += jugend[geschlecht][alter]
    print(f"Aktive Mitglieder in Kinder- und Jugendchören, {kategorie}: {summe}")

for kategorie in DCV:
    summe = 0
    for geschlecht in erwachsen:
        for alter in erwachsen[geschlecht]:
            if DCV[kategorie][0] <= alter <= DCV[kategorie][1]:
                summe += erwachsen[geschlecht][alter]

    print(f"Aktive Mitglieder in Erwachsenenchören, {kategorie}: {summe}")

print()
print("DCV-Statistik")
print("-------------")

DCV_Stat = {"bis 13 Jahre": (0,13),
            "14-18 Jahre": (14,18),
            "19-26 Jahre": (19,26),
            "27-40 Jahre": (27,40),
            "41-59 Jahre": (41,59),
            "ab 60 Jahre": (60,999)
            }

for kategorie in DCV_Stat:
    summe = {"Männlich": 0,
              "Weiblich": 0,
              "Divers":0
            }
    for geschlecht in summe:
        for alter in jugend[geschlecht]:
            if DCV_Stat[kategorie][0] <= alter <= DCV_Stat[kategorie][1]:
                summe[geschlecht] += jugend[geschlecht][alter]
        for alter in erwachsen[geschlecht]:
            if DCV_Stat[kategorie][0] <= alter <= DCV_Stat[kategorie][1]:
                summe[geschlecht] += erwachsen[geschlecht][alter]
    print(f"Kategorie {kategorie}: Männlich {summe['Männlich']}, Weiblich {summe['Weiblich']}, Divers {summe['Divers']}")

print()
print("MV-Statistik")
print("-------------")


MV_Stat = {"bis 17 Jahre": (0,17),
            "18-26 Jahre": (18,26),
            "27-59 Jahre": (27,59),
            "ab 60 Jahre": (60,999)
            }

for kategorie in MV_Stat:
    summe = {"Männlich": 0,
              "Weiblich": 0,
              "Divers":0
            }
    for geschlecht in summe:
        for alter in jugend[geschlecht]:
            if MV_Stat[kategorie][0] <= alter <= MV_Stat[kategorie][1]:
                summe[geschlecht] += jugend[geschlecht][alter]
        for alter in erwachsen[geschlecht]:
            if MV_Stat[kategorie][0] <= alter <= MV_Stat[kategorie][1]:
                summe[geschlecht] += erwachsen[geschlecht][alter]
    print(f"Kategorie {kategorie}: Männlich {summe['Männlich']}, Weiblich {summe['Weiblich']}, Divers {summe['Divers']}")
