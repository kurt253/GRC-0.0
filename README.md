# GRC Tool v0.3

Excel-gebaseerde tool voor informatieclassificatie, kwetsbaarhedenbeheer en remediëring voor Belgische overheidsorganisaties.  
Gegenereerd via Python (openpyxl + win32com), gevuld via ingebouwde VBA-macro's die data ophalen uit een MS Access databank.

**Kaders:** NIS2 · GDPR · ISO 27001:2022 · BIO  
**Talen:** NL · FR · EN

---

## Snelle start

```powershell
# 1. Sluit Excel
Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue

# 2. Bouw de tool
Set-Location "c:\Users\kurtm\Documents\GRC-0.0"
python scripts\build_grc.py

# 3. Open het resultaat
Start data\template\GRC_Tool.xlsm
```

Zie [HANDLEIDING.md](HANDLEIDING.md) voor de volledige gebruiksdocumentatie.

---

## Vereisten

| Component | Versie |
|---|---|
| Python | 3.10+ |
| openpyxl | 3.x |
| pywin32 | actueel |
| pyodbc | actueel |
| MS Excel | 2016+ (macro's ingeschakeld) |
| MS Access ACE-driver | 16.0 of 12.0 (zelfde bitness als Excel) |

```powershell
pip install openpyxl pywin32 pyodbc
```

---

## Projectstructuur

```
GRC-0.0\
├── scripts\
│   └── build_grc.py              ← enige build-script
├── data\
│   ├── template\
│   │   └── GRC_Tool.xlsm         ← OUTPUT (overschreven bij elke build)
│   ├── Import\
│   │   └── MNMTool - SocSec.accdb
│   └── repositories\
│       ├── CyFun Self-Assessment tool V2025-08-04.xlsx
│       ├── CyFun2025_Self-Assessment_tool_ESSENTIAL_v3.1.xlsx
│       └── Mapping_CyFun2023-CyFun2025_v2026-02-25.xlsx
├── HANDLEIDING.md                ← uitgebreide gebruikshandleiding
└── README.md
```

---

## Sheets

| Tab | Gevuld door | Functie |
|---|---|---|
| Config | Manueel | Taal (NL/FR/EN), instelling, versie |
| Info | Build-time | Doel, kaders, gebruiksaanwijzing |
| Processen | Import of manueel | Bedrijfsprocessen + CIA-classificatie |
| Informatieassets | Import of manueel | Informatieassets + vertrouwelijkheidsniveau |
| Dependent Assets | Import of manueel | Technische assets; basis voor RARM |
| Verantwoordelijken | Manueel | GRC-contactregister |
| Import & Export | Macro-knoppen | Alle importfuncties |
| CyFun Controls | Build-time | 218 CyFun 2025 controls + 2023-mapping |
| Kwetsbaarheden | Macro | Kwetsbaarheden × controls matrix |
| RARM | Macro + manueel | Risicomatrix per afhankelijke asset |
| Referentiewaarden | Build-time | Definities van classificatieniveaus |
| _Lang | Build-time (verborgen) | Meertalige labels |

---

## Importeer-macro's

Alle macro's zijn bereikbaar via het **Import & Export**-tabblad.

| Knop | Functie |
|---|---|
| Importeer Alles | IA → DA → Processen → Kwetsbaarheden in één keer |
| Importeer Informatieassets | Enkel `T - Information Assets` |
| Importeer Afhankelijke Assets | Enkel `T - Dependent assets` |
| Importeer Processen | Processen + koppelingen IA/DA |
| Importeer Kwetsbaarheden | Kwetsbaarheden + CIA + control-koppelingen |
| **Importeer Links DA/Kwetsbaarheden** | Vult RARM: kwetsbaarheden per DA + geselecteerde controls |

De oranje knop **Importeer Links DA/Kwetsbaarheden** staat bewust apart — deze vult het RARM-tabblad en wordt na de andere imports uitgevoerd.

---

## RARM-tabblad

De Risk Assessment & Remediation Matrix combineert:

- 218 CyFun 2025 controls (rijen)
- Alle afhankelijke assets (kolommen, gesynchroniseerd met Dependent Assets-sheet)
- Per DA: de gekoppelde kwetsbaarheid (rij 3) en de geselecteerde controls (✔ in datacellen)

Navigeren naar RARM synchroniseert automatisch de DA-kolommen zonder bestaande data te wissen.

---

## Access-databank — tabellen

| Tabel | Gebruik |
|---|---|
| `T - Processes in scope` | Processen |
| `T - Information Assets` | Informatieassets |
| `LT - Information Assets to Processes` | Koppeling processen ↔ IA |
| `T - Dependent assets` | Afhankelijke assets |
| `T - Vulnerabilities` | Kwetsbaarheden + CIA |
| `T - CyFunEssentiel` | RefNr → 2023-ID vertaling |
| `LT - Vulnerability to control - fixed` | Control-kwetsbaarheid koppelingen |
| `LT - Vulnerabilities to Dependent Assets` | Kwetsbaarheden per DA |
| `LT - Selected controls to DA` | Geselecteerde controls per DA |

---

## Bekende valkuilen

| Situatie | Oplossing |
|---|---|
| Build mislukt — bestand vergrendeld | `Stop-Process -Name "EXCEL" -Force` |
| Macro's werken niet | Klik "Inhoud inschakelen" in de gele Excel-balk |
| "Tabel niet gevonden" bij import | Verkeerde `.accdb` geselecteerd of tabelnaam afwijkend |
| ACE-driver fout (3706) | Installeer MS Access Database Engine 2016, zelfde bitness als Excel |
| RARM blijft leeg | Importeer eerst Afhankelijke Assets (kolom B van Dependent Assets moet gevuld zijn) |
| Labels tonen `[sleutel]` | Bouw de tool opnieuw — sleutel ontbreekt in `_Lang` |
| Classificatielabels fout na taalwijziging | Herselect waarden via dropdown — codes (1–5) zijn correct, labels zijn taalafhankelijk |
