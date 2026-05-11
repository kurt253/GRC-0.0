# GRC Tool

Excel-gebaseerde tool voor informatieclassificatie en kwetsbaarhedenbeheer voor Belgische overheidsorganisaties. Gegenereerd vanuit Python, gevuld via VBA-macro's die data ophalen uit een MS Access database.

## Inhoud

- [Doel](#doel)
- [Vereisten](#vereisten)
- [Projectstructuur](#projectstructuur)
- [Tool bouwen](#tool-bouwen)
- [Sheets en gebruik](#sheets-en-gebruik)
- [Data importeren via macro's](#data-importeren-via-macros)
- [Kwetsbaarheden sheet](#kwetsbaarheden-sheet)
- [Bronbestanden](#bronbestanden)
- [Bekende valkuilen](#bekende-valkuilen)

---

## Doel

De GRC Tool ondersteunt informatieveiligheidsclassificatie conform NIS2, GDPR, ISO 27001:2022 en BIO. Ze laat toe om:

- Bedrijfsprocessen te registreren met integriteits- en beschikbaarheidsvereisten
- Informatieassets te koppelen aan processen via een interactieve picker
- Kwetsbaarheden in kaart te brengen en te koppelen aan CyFun 2025 controls
- Remediëring bij te houden per kwetsbaarheid per control
- Alles te exporteren naar een .xlsx voor rapportering

---

## Vereisten

| Component | Versie |
|---|---|
| Python | 3.10+ |
| openpyxl | 3.x |
| pywin32 (win32com) | voor VBA-injectie |
| pyodbc | voor Access DB |
| MS Access ACE 16.0 of ACE 12.0 driver | voor ADODB in VBA |
| Microsoft Excel | met macro's ingeschakeld |

Installeer Python-afhankelijkheden:
```
pip install openpyxl pywin32 pyodbc
```

---

## Projectstructuur

```
GRC-0.0\
├── scripts\
│   └── build_grc.py          ← enige build-script; genereert GRC_Tool.xlsm
├── data\
│   ├── template\
│   │   └── GRC_Tool.xlsm     ← OUTPUT (wordt overschreven bij elke build)
│   ├── Import\
│   │   └── MNMTool - SocSec.accdb   ← brondata (processen, assets, kwetsbaarheden)
│   └── repositories\
│       ├── CyFun Self-Assessment tool V2025-08-04.xlsx       ← CyFun 2025 controls
│       ├── CyFun2025_Self-Assessment_tool_ESSENTIAL_v3.1.xlsx
│       └── Mapping_CyFun2023-CyFun2025_v2026-02-25.xlsx     ← ID-mapping 2023→2025
└── README.md
```

---

## Tool bouwen

**Stap 1:** Sluit Excel (of kill het proces):

```powershell
Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue
```

**Stap 2:** Verwijder eventueel de vorige versie:

```powershell
Remove-Item "data\template\GRC_Tool.xlsm" -Force -ErrorAction SilentlyContinue
```

**Stap 3:** Voer het build-script uit:

```powershell
Set-Location "c:\Users\kurtm\Documents\GRC-0.0"
python scripts\build_grc.py
```

Het script:
1. Genereert alle sheets in openpyxl
2. Injecteert alle VBA-code via win32com
3. Slaat op als `data\template\GRC_Tool.xlsm`

---

## Sheets en gebruik

### Tabvolgorde (links → rechts)

| Tab | Inhoud | Hoe gevuld |
|---|---|---|
| Config | Taal (NL/FR/EN), algemene instellingen | Manueel |
| Info | Algemene info over de organisatie | Manueel |
| Processen | Bedrijfsprocessen met classificatie | Import-macro of manueel |
| Informatieassets | Informatie-assets met vertrouwelijkheidslabel | Import-macro of manueel |
| Dependent Assets | Afhankelijke technische assets | Manueel |
| Verantwoordelijken | Eigenaars en rollen | Manueel |
| Import & Export | Knoppenpaneel voor alle macro's | Macro-knoppen |
| Referentiewaarden | Lookuptabellen voor classificatiecodes | Readonly (build-time) |
| CyFun Controls | 218 CyFun 2025 controls + 2023-mapping | Readonly (build-time) |
| Kwetsbaarheden | Kwetsbaarheden × controls matrix | Statisch kader (build) + data (macro) |
| _Lang | Meertalige labels (verborgen) | Readonly (build-time) |

### Processen (11 kolommen)

| Kolom | Inhoud |
|---|---|
| A | ID |
| B | Naam |
| C | Omschrijving |
| D | Eigenaar |
| E | Dienst |
| F | Integriteit code (1–5) |
| G | Integriteit label |
| H | Beschikbaarheid code (1–5) |
| I | Beschikbaarheid label |
| J | Opmerkingen |
| K | Gekoppelde informatieassets (klik = popup) |

### Informatieassets (8 kolommen)

| Kolom | Inhoud |
|---|---|
| A | ID |
| B | Naam |
| C | Omschrijving |
| D | Eigenaar |
| E | Dienst |
| F | Confidentialiteit code (1–5) |
| G | Confidentialiteit label |
| H | Opmerkingen |

---

## Data importeren via macro's

Open `GRC_Tool.xlsm` in Excel. Ga naar het tabblad **Import & Export** en gebruik de knoppen:

| Knop | Functie |
|---|---|
| Importeer Processen | Laadt T - Processes in scope uit Access |
| Importeer Assets | Laadt T - Information Assets uit Access |
| Importeer Koppelingen | Herlaadt proces–asset koppelingen (kolom K) |
| **Importeer Kwetsbaarheden** | Laadt kwetsbaarheden + CIA + control-koppelingen |
| Exporteer Alles | Kopieert Processen + Assets + Kwetsbaarheden naar .xlsx |

Bij elke import vraagt de macro om het `.accdb` bestand te selecteren via een bestandsdialoog.

---

## Kwetsbaarheden sheet

### Layout

```
Rij 1  : Titel (gemerged over alle kolommen)
           → bevat ook: "dubbelklik op een cel om ✔ te plaatsen"
Rij 2  : "Control ID" | "Richtlijn" | [naam kwetsbaarheid 1] | [naam kwetsbaarheid 2] | …
Rij 3  : "C"          | "Vertrouwelijkheid" | ✔ (als CIA-C) | ✔ (als CIA-C) | …
Rij 4  : "I"          | "Integriteit"       | ✔ (als CIA-I) | …
Rij 5  : "A"          | "Beschikbaarheid"   | ✔ (als CIA-A) | …
Rij 6+ : ctrl_id | richtlijn-tekst | ✔ (als die kwetsbaarheid door deze control geremediëerd wordt)
```

- **Kolom A–B**: statisch kader — 218 CyFun 2025 controls, ingeladen bij build
- **Kolom C+**: dynamisch — één kolom per kwetsbaarheid, ingeladen door de macro
- **Vriesvenster**: C6 (rijen 1–5 en kolommen A–B blijven zichtbaar bij scrollen)
- **Kleurcodering kolom A**: groen = Basic, geel = Important, rood = Essential

### Manuele input

Na het importeren via macro kan je:
- **Dubbelklikken** op een cel (rij ≥ 3, kolom ≥ C) om een ✔ te plaatsen of te verwijderen
- De eerste 5 kolommen (C–G) zijn vooraf opgemaakt voor manuele invoer van kwetsbaarheden

### Hoe de macro werkt

1. Bestandsdialoog → gebruiker selecteert `.accdb`
2. Leest kwetsbaarheden uit `T - Vulnerabilities`
3. Leest ID-vertaling uit `T-CyFunEssentiel` (numerieke RefNr → CyFun 2023 ID)
4. Leest 2023→2025 mapping uit het **CyFun Controls** sheet (kolommen F + M)
5. Leest koppelingen uit `LT - Vulnerability to control - fixed`
6. Schrijft voor elke getroffen control een ✔ (groen) in de juiste cel
7. Geeft een samenvatting + lijst van niet-gematchte controls (Afwijkingen sheet)

---

## Bronbestanden

### Access DB: `MNMTool - SocSec.accdb`

| Tabel | Relevante velden |
|---|---|
| T - Processes in scope | ID, ProcessName, ProcessDescription, IntegrityRequirement, AvailabiltyRequirement |
| T - Information Assets | ID, IAName, IA Description, Confidentiality |
| LT - Information Assets to Processes | ID, ProcessID, Information Asset ID |
| T - Vulnerabilities | Reference, Vulnerability, Confidentiality, Integrity, Availability |
| T-CyFunEssentiel | RefNr, Requirement |
| LT - Vulnerability to control - fixed | Vulnerability, CyFunControl |

### CyFun-bestanden in `data\repositories\`

| Bestand | Inhoud |
|---|---|
| CyFun Self-Assessment tool V2025-08-04.xlsx | CyFun 2025 controls (alle niveaus) |
| CyFun2025_Self-Assessment_tool_ESSENTIAL_v3.1.xlsx | Enkel Essential-niveau |
| Mapping_CyFun2023-CyFun2025_v2026-02-25.xlsx | Officiële ID-mapping 2023↔2025 |

---

## Bekende valkuilen

| Situatie | Oplossing |
|---|---|
| "Bestand vergrendeld" bij build | Sluit Excel eerst: `Stop-Process -Name "EXCEL" -Force` |
| VBE syntax error bij build | `Attribute VB_Name`-lijnen worden automatisch gestript door het script |
| `ChrW` vs `Chr` in VBA | Gebruik altijd `ChrW(10004)` voor het ✔-teken — `Chr()` accepteert max 255 |
| Kwetsbaarheden komen niet overeen | Controleer of `T-CyFunEssentiel` gevuld is en `CyFun Controls` sheet de 2023-IDs in kolom M heeft |
| Worksheet CodeName fout | VBA-modules worden aangesproken via CodeName, niet via tab-naam |

---

## Taal

De tool ondersteunt **NL / FR / EN**. De actieve taal wordt ingesteld op het Config-tabblad (cel D9). Labels worden dynamisch opgehaald uit het verborgen `_Lang` sheet via:

```excel
=IFERROR(INDEX(_Lang!$B$2:$D$300, MATCH("key", _Lang!$A$2:$A$300, 0), IF(Config!$D$9="NL",1,IF(Config!$D$9="FR",2,3))), "[key]")
```
