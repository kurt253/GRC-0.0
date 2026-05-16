# GRC Tool — Uitgebreide Handleiding

**Versie 0.3 · Belgische overheidsorganisaties · NIS2 · GDPR · ISO 27001:2022 · BIO**

---

## Inhoudsopgave

1. [Wat is de GRC Tool?](#1-wat-is-de-grc-tool)
2. [Vereisten en installatie](#2-vereisten-en-installatie)
3. [De tool bouwen](#3-de-tool-bouwen)
4. [Eerste gebruik — stap voor stap](#4-eerste-gebruik--stap-voor-stap)
5. [Tabblad Config](#5-tabblad-config)
6. [Tabblad Info](#6-tabblad-info)
7. [Tabblad Processen](#7-tabblad-processen)
8. [Tabblad Informatieassets](#8-tabblad-informatieassets)
9. [Tabblad Dependent Assets](#9-tabblad-dependent-assets)
10. [Tabblad Verantwoordelijken](#10-tabblad-verantwoordelijken)
11. [Tabblad Import & Export](#11-tabblad-import--export)
12. [Tabblad CyFun Controls](#12-tabblad-cyfun-controls)
13. [Tabblad Kwetsbaarheden](#13-tabblad-kwetsbaarheden)
14. [Tabblad RARM](#14-tabblad-rarm)
15. [Tabblad Referentiewaarden](#15-tabblad-referentiewaarden)
16. [Classificatieniveaus (1–5)](#16-classificatieniveaus-15)
17. [Importeer-macro's in detail](#17-importeer-macros-in-detail)
18. [De Access-databank](#18-de-access-databank)
19. [Meertaligheid](#19-meertaligheid)
20. [Veelgestelde vragen en probleemoplossing](#20-veelgestelde-vragen-en-probleemoplossing)

---

## 1. Wat is de GRC Tool?

De GRC Tool is een Excel-werkmap (`.xlsm`) voor **Governance, Risk & Compliance** in Belgische overheidsorganisaties. Ze biedt een gestructureerde manier om:

- **Bedrijfsprocessen** te registreren en te classificeren op integriteit en beschikbaarheid
- **Informatieassets** bij te houden met vertrouwelijkheidsniveaus
- **Afhankelijke (technische) assets** te koppelen aan processen en informatieassets
- **Kwetsbaarheden** in kaart te brengen per CyFun 2025 control
- **Remediëring** te plannen en op te volgen via de RARM-matrix
- Alles te **exporteren** voor rapportering

De tool is volledig gegenereerd via een Python-script (`build_grc.py`) en bevat ingebouwde VBA-macro's die data importeren uit een MS Access databank.

**Ondersteunde kaders:** NIS2 · GDPR · ISO 27001:2022 · BIO  
**Talen:** Nederlands (NL) · Frans (FR) · Engels (EN)

---

## 2. Vereisten en installatie

### Software

| Component | Minimale versie | Opmerking |
|---|---|---|
| Microsoft Excel | 2016+ | Macro's moeten ingeschakeld zijn |
| Python | 3.10+ | Voor het bouwen van de tool |
| openpyxl | 3.x | `pip install openpyxl` |
| pywin32 | actueel | `pip install pywin32` (VBA-injectie) |
| pyodbc | actueel | `pip install pyodbc` (Access-lezing bij build) |
| MS Access ACE-driver | 16.0 of 12.0 | Nodig voor ADODB in VBA tijdens import |

### ACE-driver installeren (indien ontbrekend)

De ACE-driver is nodig zodat de VBA-macro's de Access-databank kunnen openen. Download via Microsoft:
- `Microsoft Access Database Engine 2016 Redistributable`

Controleer de aanwezige driver in het register of via Excel → Bestand → Opties → Invoegtoepassingen → COM-invoegtoepassingen.

> **Let op:** de bitsversie (32-bit of 64-bit) van de ACE-driver moet overeenkomen met die van Excel.

---

## 3. De tool bouwen

De `.xlsm` wordt volledig opnieuw aangemaakt via het Python-script. Voer dit uit telkens wanneer de code of structuur is aangepast.

### Stap 1 — Sluit Excel

```powershell
Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue
```

### Stap 2 — Verwijder de vorige versie (optioneel maar aanbevolen)

```powershell
Remove-Item "c:\Users\kurtm\Documents\GRC-0.0\data\template\GRC_Tool.xlsm" -Force -ErrorAction SilentlyContinue
```

### Stap 3 — Build uitvoeren

```powershell
Set-Location "c:\Users\kurtm\Documents\GRC-0.0"
python scripts\build_grc.py
```

Het script doorloopt drie fasen:
1. **openpyxl-fase**: alle sheets, opmaak, formules en statische data worden aangemaakt; tussenresultaat opgeslagen als `.xlsx`
2. **win32com-fase**: Excel opent de `.xlsx` via COM-automatisering; VBA-code wordt geïnjecteerd en knoppen worden toegevoegd
3. **Opslaan**: het bestand wordt opgeslagen als `GRC_Tool.xlsm`

Verwachte output:
```
Tussenstap opgeslagen: ...GRC_Tool.xlsx
GRC Tool v0.3 (macro-enabled) opgeslagen: ...GRC_Tool.xlsm
Gegenereerd door: [gebruikersnaam] op [datum]
```

---

## 4. Eerste gebruik — stap voor stap

1. Open `GRC_Tool.xlsm` in Excel. Klik op **Macro's inschakelen** wanneer Excel dit vraagt.
2. Ga naar **Config** en stel de naam van de instelling en de gewenste taal in.
3. Ga naar **Import & Export** en gebruik de importknoppen om data uit de Access-databank in te laden.
4. Controleer en corrigeer indien nodig op de datasheets (Processen, Informatieassets, Dependent Assets).
5. Ga naar **Kwetsbaarheden** voor een overzicht van kwetsbaarheden per CyFun-control.
6. Ga naar **RARM** voor de risicomatrix per afhankelijke asset.
7. Sla het bestand op met **Ctrl+S** (als `.xlsm`).

---

## 5. Tabblad Config

Het Config-tabblad bevat de basisinstellingen van de tool.

### Bewerkbare velden

| Veld | Inhoud |
|---|---|
| Naam overheidsinstelling | Volledige naam van de organisatie |
| Dienst / Entiteit | Naam van de dienst of afdeling |
| Taal | `NL`, `FR` of `EN` — wijzigt alle labels in de tool |

### Taalinstelling

De taal wordt ingesteld in cel **D9**. Na het wijzigen passen alle taalformules zich automatisch aan. Opgelet: bestaande dropdownwaarden (bijv. classificatieniveaus) die al ingegeven zijn moeten opnieuw geselecteerd worden, omdat de tekst van de opties verandert.

### Niet-bewerkbare velden

Versienummer, aanmaakdatum en bouwgebruiker worden automatisch ingevuld bij de build en zijn beveiligd.

---

## 6. Tabblad Info

Informatief tabblad met uitleg over:
- het doel van de tool
- de ondersteunde wettelijke kaders (NIS2, GDPR, ISO 27001:2022, BIO)
- de modules die beschikbaar zijn in deze versie
- een beknopte gebruiksaanwijzing
- versie-informatie

Dit tabblad is enkel ter referentie; er zijn geen invoervelden.

---

## 7. Tabblad Processen

Overzicht van alle bedrijfsprocessen in scope, met classificatie op integriteit en beschikbaarheid.

### Kolomstructuur

| Kolom | Naam | Inhoud |
|---|---|---|
| A | ID | Automatisch gegenereerd volgnummer |
| B | Naam | Naam van het proces |
| C | Omschrijving | Beschrijving of toelichting |
| D | Eigenaar | Naam van de proceseigenaar |
| E | Dienst | Organisatorische eenheid |
| F | Integriteit (code) | Numerieke waarde 1–5 — **bewerkbaar via dropdown** |
| G | Integriteit (label) | Automatisch label op basis van kolom F |
| H | Beschikbaarheid (code) | Numerieke waarde 1–5 — **bewerkbaar via dropdown** |
| I | Beschikbaarheid (label) | Automatisch label op basis van kolom H |
| J | Opmerkingen | Vrij tekstveld |
| K | Gekoppelde informatieassets | Lijst van gekoppelde assets (via popup of import) |

### Classificatie invoeren

Klik op een cel in kolom F of H om een dropdown te activeren met de 5 niveaus. Het bijhorende label in kolom G of I wordt automatisch ingevuld via een formule.

### Kleurcodering

Cellen met een classificatiewaarde krijgen een achtergrondkleur:

| Niveau | Kleur |
|---|---|
| 1 — Laag | Lichtgroen |
| 2 — Gemiddeld | Lichtgeel |
| 3 — Hoog | Lichtorange |
| 4 — Zeer Hoog | Lichtrood |
| 5 — Classified | Lichtpaars |

### Informatieassets koppelen (popup)

Klik op een cel in **kolom K** om de AssetPicker te openen — een popup die alle beschikbare informatieassets toont. Meerdere assets kunnen geselecteerd worden. De namen worden gescheiden door een regelafbreking in de cel geplaatst.

### Data importeren

Gebruik de knop **Importeer Processen** op het Import & Export-tabblad. Zie [sectie 17](#17-importeer-macros-in-detail) voor details.

---

## 8. Tabblad Informatieassets

Overzicht van alle informatieassets (gegevenscollecties, systemen, documenten…) met vertrouwelijkheidsclassificatie.

### Kolomstructuur

| Kolom | Naam | Inhoud |
|---|---|---|
| A | ID | Automatisch gegenereerd volgnummer |
| B | Naam | Naam van het informatie-asset |
| C | Omschrijving | Beschrijving |
| D | Eigenaar | Assetverantwoordelijke |
| E | Dienst | Organisatorische eenheid |
| F | Confidentialiteit (code) | Numerieke waarde 1–5 — **bewerkbaar via dropdown** |
| G | Confidentialiteit (label) | Automatisch label |
| H | Opmerkingen | Vrij tekstveld |

### Data importeren

Gebruik de knop **Importeer Informatieassets** op het Import & Export-tabblad.

---

## 9. Tabblad Dependent Assets

Afhankelijke (technische) assets zijn de IT-systemen, netwerken, servers en diensten waarop bedrijfsprocessen steunen. Dit tabblad is de basis voor de RARM-matrix.

### Kolomstructuur (selectie)

| Kolom | Naam | Inhoud |
|---|---|---|
| A | ID | Volgnummer |
| B | Naam | Naam van de afhankelijke asset |
| C | Omschrijving | Beschrijving |
| D | Eigenaar | Verantwoordelijke |
| E | Dienst | Organisatorische eenheid |
| F | Overarching | ✔ als dit een overkoepelende/centrale asset is |
| G | Opmerkingen | Vrij tekstveld |
| P | C-doelstelling | Confidentialiteitsdoelstelling (1–5) |
| R | I-doelstelling | Integriteitsdoelstelling (1–5) |
| T | A-doelstelling | Beschikbaarheidsdoelstelling (1–5) |

### Overarching-vlag

Kolom F bevat een ✔ voor assets die als overkoepelend beschouwd worden. In de RARM-matrix worden deze kolommen in een afwijkende kleur weergegeven (oranje in plaats van geel).

### Dubbelklik-interactie

Dubbelklik op een cel in **kolom F** (rij 6+) om de overarching-vlag te plaatsen of te verwijderen.

### Koppeling met RARM

Elke afhankelijke asset in dit tabblad (rijen 6–105, kolom B) krijgt automatisch een kolom in de RARM-matrix wanneer het RARM-tabblad geactiveerd wordt (`SyncRARMKolommen`). Assets die verwijderd worden uit dit tabblad, verdwijnen ook uit RARM bij de volgende sync.

### Data importeren

Gebruik de knop **Importeer Afhankelijke Assets** op het Import & Export-tabblad.

---

## 10. Tabblad Verantwoordelijken

Contactregister van alle personen die een GRC-rol hebben binnen de organisatie.

### Kolomstructuur

| Kolom | Inhoud |
|---|---|
| B | Naam |
| C | Functietitel |
| D | Rol (dropdown: CISO, DPO, Proceseigenaar…) |
| E | E-mailadres |
| F | Telefoonnummer |
| G | Dienst / Eenheid |
| H | Opmerkingen |

Dit tabblad wordt manueel ingevuld. Er is geen importmacro voorzien.

---

## 11. Tabblad Import & Export

Het centrale bedieningspaneel met knoppen voor alle macro's.

### Importknoppen

| Knop | Macro | Functie |
|---|---|---|
| **Importeer Alles** | `ImportAlles` | Voert de volledige import uit: IA → DA → Processen → Kwetsbaarheden (in vaste volgorde) |
| **Importeer Informatieassets** | `ImportInformatieassets` | Laadt enkel de informatieassets uit `T - Information Assets` |
| **Importeer Afhankelijke Assets** | `ImportAfhankelijkeAssets` | Laadt enkel de afhankelijke assets uit `T - Dependent assets` |
| **Importeer Processen** | `ImportProcessen` | Laadt processen + koppelingen (IA en DA) uit `T - Processes in scope` |
| **Importeer Kwetsbaarheden** | `ImportKwetsbaarheden` | Laadt kwetsbaarheden + CIA-kenmerken + control-koppelingen in de Kwetsbaarheden-sheet |
| **Importeer Links DA/Kwetsbaarheden** | `ImportLinksKwetsbaarheden` | Laadt welke kwetsbaarheden aan welke DA gekoppeld zijn, én welke controls geselecteerd zijn (→ RARM) |

### Volgorde van importeren

Bij **Importeer Alles** worden de imports in deze volgorde uitgevoerd:

1. Informatieassets
2. Afhankelijke assets
3. Processen (inclusief koppelingen aan IA en DA)
4. Kwetsbaarheden

De koppeling tussen kwetsbaarheden/controls en afhankelijke assets (**Links DA/Kwetsbaarheden**) is een afzonderlijke stap die expliciet uitgevoerd moet worden via de oranje knop.

### Bestandsdialoog

Elke importknop opent een **bestandsdialoog** waarmee je de Access-databank (`.accdb` of `.mdb`) selecteert. Het pad wordt niet opgeslagen; de dialoog verschijnt bij elke import opnieuw.

---

## 12. Tabblad CyFun Controls

Referentietabblad met alle **218 CyFun 2025 Essential-controls**, inclusief de mapping naar de vorige CyFun 2023-IDs.

### Structuur

| Kolom | Inhoud |
|---|---|
| A | Domein |
| B | Subdomein |
| C | Niveau (Basic / Important / Essential) |
| D | Sleutelmaatregel (✔ = ja) |
| E | Assurance-niveau |
| F | CyFun 2025 Control ID (bijv. `gv.oc-01.1`) |
| G | Richtlijn-tekst |
| H–L | Extra velden uit het bronbestand |
| M | CyFun 2023 Control ID (bijv. `ID.AM-1.1`) — gebruikt voor de vertaalketen |

### Gebruik

Dit tabblad is **alleen-lezen** (ingevuld bij build-time). Het dient als:
- visuele referentie voor het beheer
- brondata voor de Kwetsbaarheden-sheet en de RARM-matrix
- vertaalketen van 2023 naar 2025 IDs bij de importmacro's

### Kleurcodering kolom C

| Niveau | Kleur |
|---|---|
| Basic | Groen |
| Important | Geel |
| Essential | Rood |

---

## 13. Tabblad Kwetsbaarheden

De **Kwetsbaarheden-matrix** toont per kwetsbaarheid welke CyFun 2025-controls remediëring bieden, en op welke CIA-dimensies de kwetsbaarheid van toepassing is.

### Lay-out

```
Rij 1  : Titelbalk (gemerged over alle kolommen)
Rij 2  : "Control ID" | "Richtlijn" | [naam kwetsbaarheid 1] | [naam kwetsbaarheid 2] | …
Rij 3  : "C"          | "Vertrouwelijkheid" | ✔ (als kwetsbaarheid C-impact heeft)
Rij 4  : "I"          | "Integriteit"       | ✔ (als kwetsbaarheid I-impact heeft)
Rij 5  : "A"          | "Beschikbaarheid"   | ✔ (als kwetsbaarheid A-impact heeft)
Rij 6+ : ctrl_id      | richtlijn-tekst     | ✔ (als control remediëring biedt)
```

### Kolommen

- **Kolom A** — Control ID (bijv. `gv.oc-01.1`): statisch, ingevuld bij build
- **Kolom B** — Richtlijn-tekst: statisch, ingevuld bij build
- **Kolom C+** — één kolom per kwetsbaarheid: **dynamisch**, ingevuld door de importmacro

### Vriesvenster

Kolommen A–B en rijen 1–5 zijn bevroren (vriesvenster op C6). Bij horizontaal en verticaal scrollen blijven de control-IDs en kwetsbaarhedennamen zichtbaar.

### Kleurcodering rijen (kolom A)

| Niveau | Kleur |
|---|---|
| Basic | Groen |
| Important | Geel / Oranje |
| Essential | Rood |

### Manuele invoer (dubbelklik)

Na de import kan je handmatig een ✔ toevoegen of verwijderen:
- **Dubbelklik** op een datacel (rij ≥ 6, kolom ≥ C) om de ✔ te plaatsen of te wissen
- De ✔ wordt automatisch gecentreerd in de cel

### Import via macro

De knop **Importeer Kwetsbaarheden** voert de volgende stappen uit:

1. Leest kwetsbaarheden + CIA-kenmerken uit `T - Vulnerabilities`
2. Leest de ID-vertaling uit `T - CyFunEssentiel` (numerieke RefNr → 2023-ID)
3. Vertaalt 2023-ID naar 2025-ID via het CyFun Controls-tabblad (kolommen F + M)
4. Leest control-koppelingen uit `LT - Vulnerability to control - fixed`
5. Schrijft een **groen ✔** in elke cel waar de control de kwetsbaarheid remediëert

---

## 14. Tabblad RARM

**RARM** staat voor Risk Assessment & Remediation Matrix. Dit tabblad toont per afhankelijke asset (DA) welke kwetsbaarheden aanwezig zijn en welke CyFun-controls geselecteerd zijn als maatregel.

### Lay-out

```
Rij 1      : Titelbalk (gemerged)
Rij 2      : "Control ID" | "Richtlijn" | "Assurance" | "Sleutelmaatregel" | [DA-naam 1] | [DA-naam 2] | …
Rij 3      : (grijs)       | (grijs)     | (grijs)     | (grijs)            | [kwetsbaarheid voor DA] | …
Rij 4+     : ctrl_id       | richtlijn   | assurance   | ✔ (key measure)   | ✔ (als control geselecteerd voor DA)
```

### Kolommen A–D (statisch)

| Kolom | Naam | Inhoud |
|---|---|---|
| A | Control ID | CyFun 2025 ID (bijv. `gv.oc-01.1`) |
| B | Richtlijn | Beschrijving van de control |
| C | Assurance | Assurance-niveau uit CyFun |
| D | Sleutelmaatregel | ✔ als dit een van de 29 CyFun-sleutelmaatregelen is |

### Kolommen E+ (dynamisch — per DA)

Elke afhankelijke asset uit het Dependent Assets-tabblad krijgt hier een kolom. De kleur geeft aan of de asset overarching is:

| Type | Kolomkleur header |
|---|---|
| Gewone asset | Geel |
| Overarching asset | Oranje |

**Rij 3** (grijze balk onder de DA-naam) toont de kwetsbaarheid die via de importmacro aan die DA gekoppeld is.

**Rij 4+** (datacellen) bevatten ✔ voor elke control die geselecteerd werd als maatregel voor die DA.

### Synchronisatie met Dependent Assets

Elke keer je naar het RARM-tabblad navigeert, wordt `SyncRARMKolommen` automatisch uitgevoerd:
- **Nieuwe DAs** in het Dependent Assets-tabblad worden als kolom toegevoegd
- **Verwijderde DAs** worden uit RARM verwijderd
- **Bestaande data** (rij 3 en vinkjes) blijft bewaard

### Dubbelklik-interactie

Dubbelklik op een datacel (rij ≥ 4, kolom ≥ E) om handmatig een ✔ te plaatsen of te wissen.

### Import via macro

De knop **Importeer Links DA/Kwetsbaarheden** voert twee stappen uit:

1. **`ImportRARMKwetsbaarheden`** — leest `LT - Vulnerabilities to Dependent Assets` en plaatst de naam van de gekoppelde kwetsbaarheid in rij 3 van de DA-kolom
2. **`ImportGeselecteerdeControls`** — leest `LT - Selected controls to DA` en schrijft ✔ in de overeenkomstige datacellen

**Vertaalketen voor controls:**
```
ControlReference (RefNr)
  → T-CyFunEssentiel.Requirement (2023-ID)
  → CyFun Controls sheet col F + M (2025-ID)
  → rijnummer in RARM
```

### Kleurcodering datacellen

| Kleur | Betekenis |
|---|---|
| Geel (licht) | Oneven rij, geen vinkje |
| Wit | Even rij, geen vinkje |
| (ongewijzigd) | Geplaatst ✔ via dubbelklik |

---

## 15. Tabblad Referentiewaarden

Overzichtstabblad met de definities van de 5 classificatieniveaus voor alle drie dimensies (C, I, A).

| Niveau | Naam (NL) | Beschrijving |
|---|---|---|
| 1 | Laag | Minimale impact bij schending |
| 2 | Gemiddeld | Beperkte impact |
| 3 | Hoog | Significante impact |
| 4 | Zeer Hoog | Zware impact op werking of reputatie |
| 5 | Classified | Gerubriceerde informatie / nationale veiligheid |

Dit tabblad is alleen-lezen en dient als referentie bij het toekennen van classificatiewaarden.

---

## 16. Classificatieniveaus (1–5)

De tool hanteert vijf niveaus voor drie dimensies:

| Dimensie | Afkorting | Wat wordt gemeten |
|---|---|---|
| Confidentialiteit | C | Gevolgen van ongeoorloofde openbaarmaking |
| Integriteit | I | Gevolgen van foutieve of onvolledige gegevens |
| Beschikbaarheid | A | Gevolgen van niet-beschikbaarheid van het systeem/gegeven |

**Niveaus:**

| Code | NL | FR | EN |
|---|---|---|---|
| 1 | Laag | Faible | Low |
| 2 | Gemiddeld | Moyen | Medium |
| 3 | Hoog | Élevé | High |
| 4 | Zeer Hoog | Très élevé | Very High |
| 5 | Classified | Classified | Classified |

---

## 17. Importeer-macro's in detail

### ImportAlles

Voert alle onderstaande imports achtereenvolgens uit vanuit één enkele bestandsdialoog:

1. Informatieassets (`T - Information Assets` → sheet *Information Assets*)
2. Afhankelijke assets (`T - Dependent assets` → sheet *Dependent Assets*)
3. Processen (`T - Processes in scope` → sheet *Processes*, inclusief koppelingen via `LT - Information Assets to Processes` en `LT - ...`)
4. Kwetsbaarheden (zie hieronder)

**Let op:** De koppeling van controls/kwetsbaarheden aan DAs (RARM-vulling) is géén onderdeel van ImportAlles en moet apart uitgevoerd worden.

---

### ImportInformatieassets

Leest `T - Information Assets` en schrijft naar *Information Assets* sheet (rijen 6–105):

- Naam (dynamisch gedetecteerd op `naam`, `name`, `nom`, …)
- Omschrijving
- Eigenaar
- Dienst
- Confidentialiteitscode (vertaald via `MapCls`)

Bestaande data in kolommen B–F en H wordt eerst gewist.

---

### ImportAfhankelijkeAssets

Leest `T - Dependent assets` en schrijft naar *Dependent Assets* sheet (rijen 6–105):

- Naam, Omschrijving, Eigenaar, Dienst
- Overarching-vlag (kolom F): ✔ als het veld `Overarching` in de DB `True` is
- C/I/A-doelstellingen (kolommen P, R, T)

Na de import worden de RARM-kolommen gesynchroniseerd via `KleurAlleRARMKolommen`.

---

### ImportProcessen

Leest `T - Processes in scope` en schrijft naar *Processes* sheet (rijen 6–105):

- Naam, Omschrijving, Eigenaar, Dienst
- Integriteitscode en beschikbaarheidscode

Na de procesimport worden ook de koppelingen opgehaald:
- Gekoppelde informatieassets (kolom K) via `LT - Information Assets to Processes`
- Gekoppelde afhankelijke assets (kolom L) via de koppelingstabel

---

### ImportKwetsbaarheden

Leest de Kwetsbaarheden-sheet via `CoreImportKwetsbaarheden`:

1. `T - Vulnerabilities` — elke kwetsbaarheid wordt een kolom (naam in rij 2, CIA in rijen 3–5)
2. `T - CyFunEssentiel` — vertaalt numerieke RefNr naar tekstueel 2023-ID
3. `LT - Vulnerability to control - fixed` — per kwetsbaarheid-control koppeling: zoekt het 2025-rijnummer op en schrijft ✔ (groen)

---

### ImportLinksKwetsbaarheden

Vult het RARM-tabblad op basis van twee tabellen:

**`ImportRARMKwetsbaarheden`** (rij 3 van elke DA-kolom):
- SQL JOIN op `T - Dependent assets`, `LT - Vulnerabilities to Dependent Assets`, `T - Vulnerabilities`
- Groepeert kwetsbaarheden per DA-naam; schrijft de kwetsbaarheidsnamen in rij 3

**`ImportGeselecteerdeControls`** (datacellen rij 4+):
- Leest `LT - Selected controls to DA`
- Vertaalketen: ControlReference → RefNr → 2023-ID → 2025-ID → rijnummer in RARM
- Schrijft ✔ in de cel [control-rij, DA-kolom]

---

## 18. De Access-databank

De tool werkt met een MS Access `.accdb`-bestand (standaard: `MNMTool - SocSec.accdb`).

### Relevante tabellen

| Tabel | Velden | Gebruik |
|---|---|---|
| `T - Processes in scope` | ID, ProcessName, ProcessDescription, IntegrityRequirement, AvailabiltyRequirement | Processen importeren |
| `T - Information Assets` | ID, IAName, IA Description, Confidentiality | Informatieassets importeren |
| `LT - Information Assets to Processes` | ID, ProcessID, Information Asset ID | Koppeling processen ↔ IA |
| `T - Dependent assets` | ID, DAName, DADescription, C-objective, I-objective, A-Objective, Overarching | DA's importeren |
| `T - Vulnerabilities` | Reference, Vulnerability, Confidentiality, Integrity, Availability | Kwetsbaarheden importeren |
| `T - CyFunEssentiel` | RefNr, Requirement | Vertaling RefNr → 2023-ID |
| `LT - Vulnerability to control - fixed` | Vulnerability, CyFunControl | Control-kwetsbaarheid koppelingen |
| `LT - Vulnerabilities to Dependent Assets` | DAID, VulnerabilityID, Probability | Kwetsbaarheden per DA (RARM rij 3) |
| `LT - Selected controls to DA` | DAID, ControlReference | Geselecteerde controls per DA (RARM datacellen) |
| `RT - Probability` | ID, Probvalue | Kanswaarden (1–4) |

### Dynamische velddetectie

De importmacro's detecteren veldnamen dynamisch via de `FieldVal`-functie:
- **Pass 1**: exacte match (case-insensitive)
- **Pass 2**: gedeeltelijke match via `InStr`

Dit maakt de macro robuust bij kleine naamvariaties.

> **Bekende afwijking:** het veld `AvailabiltyRequirement` (in `T - Processes in scope`) bevat een typefout — de `i` ontbreekt. De macro houdt hier rekening mee.

---

## 19. Meertaligheid

De tool ondersteunt drie talen die op elk moment gewisseld kunnen worden via de **Config**-sheet (cel D9).

### Hoe het werkt

Alle labels zijn opgeslagen als sleutel-waardeparen in het verborgen `_Lang`-sheet (kolom A = sleutel, B = NL, C = FR, D = EN). Elke cel met een label gebruikt de taalformule:

```excel
=IFERROR(
  INDEX(_Lang!$B$2:$D$300,
    MATCH("sleutel", _Lang!$A$2:$A$300, 0),
    IF(Config!$D$9="NL", 1, IF(Config!$D$9="FR", 2, 3))
  ), "[sleutel]"
)
```

Als een sleutel ontbreekt in de tabel, toont de cel `[sleutel]` als signaal dat de vertaling ontbreekt.

### Dropdownwaarden

Dropdownvalidaties (classificatieniveaus, rollen, assettypes) zijn ook meertalig. Ze wisselen automatisch mee wanneer de taal verandert. Opgelet: de **code** (1–5) blijft taalvrij; enkel het **label** verandert.

### Beperkingen

- Na een taalwijziging moeten bestaande classificatiewaarden in de data-sheets opnieuw geselecteerd worden (de geselecteerde tekst komt niet overeen met de nieuwe taalwaarden).
- Macro-berichten (`MsgBox`) zijn uitsluitend in het **Nederlands** geschreven.

---

## 20. Veelgestelde vragen en probleemoplossing

### De build mislukt met "Bestand vergrendeld" of een COM-fout

Excel is nog open. Sluit Excel volledig:
```powershell
Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue
```

### De macro's werken niet na het openen van de XLSM

Excel heeft macro's geblokkeerd. Controleer:
1. Klik op de gele balk "Macro's inschakelen" bovenaan
2. Of ga naar Bestand → Info → Inhoud inschakelen

### De importmacro geeft "Tabel niet gevonden"

De geselecteerde databank bevat niet de verwachte tabel. Controleer:
- Of je het juiste `.accdb`-bestand geselecteerd hebt
- Of de tabelnaam exact overeenkomt (zie sectie 18)

### De query "QLT2 - Vulnerabilities for DA" geeft geen resultaat

Dit is een bekende beperking van de Access-query die afhankelijk is van form-parameters. De tool omzeilt dit via directe SQL-JOIN op de onderliggende tabellen. Zorg dat `LT - Vulnerabilities to Dependent Assets` gevuld is.

### RARM is leeg na het navigeren naar het tabblad

De `SyncRARMKolommen`-macro wordt bij activatie uitgevoerd. Als de Dependent Assets-sheet leeg is (rijen 6–105 zonder inhoud in kolom B), worden geen kolommen aangemaakt. Importeer eerst de afhankelijke assets.

### Vinkjes staan niet in het midden van de cel

Dit kan voorkomen bij handmatig ingevoerde waarden vóór een recente build. Na een nieuwe import worden alle vinkjes automatisch gecentreerd. Manueel corrigeren: selecteer de betrokken cellen en stel Horizontale en Verticale uitlijning in op "Gecentreerd".

### De ChrW(10004)-tekens worden als vraagtekens weergegeven

Dit is een lettertypeprobleem. Zorg dat het lettertype "Calibri" of een ander Unicode-font actief is voor die cellen.

### Taalformules tonen `[sleutel]`

De sleutel bestaat niet in het `_Lang`-sheet. Dit duidt op een build-probleem. Bouw de tool opnieuw met het laatste script.

### Na taalwijziging kloppen de classificatielabels niet meer

Herselect de waarden via de dropdown in de betreffende cellen. De codes (1–5) zijn correct; alleen het weergegeven label is taalafhankelijk.

### De ACE-driver ontbreekt (fout 3706 of "Provider niet gevonden")

Installeer de Microsoft Access Database Engine 2016 Redistributable. Let op de bitsversie (32/64-bit) — die moet overeenkomen met Excel.

---

*Gegenereerd bij v0.3 · GRC Tool voor Belgische overheidsorganisaties*
