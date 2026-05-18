"""
GRC Tool — Excel Workbook Builder v0.3
Voor overheidsorganisaties
Sheets (volgorde): Config | Info | Processen | Informatieassets | Verantwoordelijken
                   | Import & Export | Referentiewaarden | _Lang (verborgen)
Output: ../data/template/GRC_Tool.xlsm
"""

import datetime
import io
import os
import re
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName

# ── Paden & metadata ──────────────────────────────────────────────────────────
OUT      = Path(__file__).parent.parent / "data" / "template" / "GRC_Tool.xlsm"
OUT_XLSX = OUT.with_suffix(".xlsx")   # tussenstap voor win32com
CYFUN_SRC   = Path(__file__).parent.parent / "data" / "repositories" / "CyFun2025_Self-Assessment_tool_ESSENTIAL_v3.1.xlsx"
MAPPING_SRC = Path(__file__).parent.parent / "data" / "repositories" / "Mapping_CyFun2023-CyFun2025_v2026-02-25.xlsx"
CYFUN23_SRC = Path(__file__).parent.parent / "data" / "repositories" / "CyFun Self-Assessment tool V2025-08-04.xlsx"
USERNAME = os.environ.get("USERNAME", os.environ.get("USER", "Onbekend"))
NOW_STR  = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

# ── Kleurenpalet ──────────────────────────────────────────────────────────────
C = {
    "navy":          "0F2B46",
    "blue_mid":      "2563EB",
    "blue_light":    "DBEAFE",
    "blue_xlight":   "EFF6FF",
    "white":         "FFFFFF",
    "grey_border":   "CBD5E1",
    "grey_light":    "F8FAFC",
    "grey_mid":      "E2E8F0",
    "grey_dark":     "475569",
    "text":          "0F172A",
    "subtext":       "64748B",
    # Classificatieniveaus
    "green":         "15803D",  "green_light":   "DCFCE7",
    "yellow":        "854D0E",  "yellow_light":  "FEF9C3",
    "orange":        "B45309",  "orange_light":  "FEF3C7",
    "red":           "B91C1C",  "red_light":     "FEE2E2",
    "kritiek":       "991B1B",  "kritiek_bg":    "991B1B",
    # UI-accenten
    "accent":        "0EA5E9",  "accent_light":  "E0F2FE",
    "purple":        "6B21A8",  "purple_light":  "F3E8FF",
}

# Vijf classificatieniveaus: index 0 = niveau 1 (Laag) … index 4 = niveau 5 (Classified)
LEVEL_FILLS = ["green_light", "yellow_light", "orange_light", "red_light", "purple_light"]
LEVEL_FONTS = ["green",       "yellow",       "orange",       "red",       "purple"]

# ── Stijlhulpfuncties ─────────────────────────────────────────────────────────
def fill(key): return PatternFill("solid", fgColor=C[key])
def font(size=10, bold=False, color="text", italic=False):
    return Font(name="Calibri", size=size, bold=bold, color=C[color], italic=italic)
def align(h="left", v="center", wrap=False, indent=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=indent)
def border_all(color="grey_border"):
    s = Side(style="thin", color=C[color])
    return Border(left=s, right=s, top=s, bottom=s)
def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
def no_gridlines(ws): ws.sheet_view.showGridLines = False

# ── Taalformule ───────────────────────────────────────────────────────────────
LANG_CELL = "Config!$D$9"

def t(key):
    return (
        f'=IFERROR(INDEX(_Lang!$B$2:$D$300,'
        f'MATCH("{key}",_Lang!$A$2:$A$300,0),'
        f'IF({LANG_CELL}="NL",1,IF({LANG_CELL}="FR",2,3))),"[{key}]")'
    )

# ── Classificatiewaarden (5 niveaus) ─────────────────────────────────────────
CLS = {
    "NL": ["Laag", "Gemiddeld", "Hoog", "Zeer Hoog", "Classified"],
    "FR": ["Faible", "Moyen", "Élevé", "Très élevé", "Classified"],
    "EN": ["Low",  "Medium",  "High",  "Very High",  "Classified"],
}
ROLES = {
    "NL": ["CISO", "DPO / FG", "Proceseigenaar", "Informatiebeheerder",
           "Systeembeheerder", "Intern auditor", "Directie / Management", "Andere"],
    "FR": ["RSSI", "DPD / FG", "Propriétaire du processus", "Gestionnaire de l'information",
           "Administrateur système", "Auditeur interne", "Direction / Management", "Autre"],
    "EN": ["CISO", "DPO", "Process Owner", "Information Manager",
           "System Administrator", "Internal Auditor", "Management", "Other"],
}
ASSET_TYPES = {
    "NL": ["Applicatie", "Databank", "Server", "Hardware", "Netwerkapparatuur",
           "Dienst", "Document", "Persoon", "Proces", "Andere"],
    "FR": ["Application", "Base de données", "Serveur", "Matériel", "Équipement réseau",
           "Service", "Document", "Personne", "Processus", "Autre"],
    "EN": ["Application", "Database", "Server", "Hardware", "Network Equipment",
           "Service", "Document", "Person", "Process", "Other"],
}
# ── Vertalingstabel ───────────────────────────────────────────────────────────
TRANSLATIONS = [
    # Paginatitel Config / Chrome
    ("cfg_title_main",
        "GRC Tool — Informatieveiligheid voor Overheidsdiensten",
        "Outil GRC — Sécurité de l'information pour les services publics",
        "GRC Tool — Information Security for Government Services"),
    ("app_tagline", "Governance · Risico · Compliance",
        "Gouvernance · Risque · Conformité", "Governance · Risk · Compliance"),
    ("ui_tips",         "TIPS",              "CONSEILS",         "TIPS"),
    ("ui_quicklinks",   "SNELKOPPELINGEN",   "RACCOURCIS",       "QUICK LINKS"),
    ("ui_goto_config",  "→  Configuratie",   "→  Configuration", "→  Configuration"),
    ("ui_goto_ref",     "→  Referentiewaarden", "→  Valeurs de référence", "→  Reference Values"),
    ("ui_goto_proc",    "→  Processen",      "→  Processus",     "→  Processes"),
    ("ui_goto_assets",  "→  Informatieassets","→  Actifs informationnels","→  Information Assets"),
    ("ui_goto_verant",  "→  Verantwoordelijken","→  Responsables","→  Responsible Persons"),
    ("ui_goto_import",  "→  Import & Export","→  Import & Export","→  Import & Export"),
    ("ui_goto_info",    "→  Informatie",     "→  Information",   "→  Information"),
    # Config
    ("cfg_section_general", "ALGEMENE INSTELLINGEN",     "PARAMÈTRES GÉNÉRAUX",      "GENERAL SETTINGS"),
    ("cfg_section_display", "WEERGAVE-INSTELLINGEN",     "PARAMÈTRES D'AFFICHAGE",   "DISPLAY SETTINGS"),
    ("cfg_section_meta",    "OVER DIT BESTAND",          "À PROPOS DE CE FICHIER",   "ABOUT THIS FILE"),
    ("cfg_org_label",       "Naam overheidsinstelling",  "Nom de l'institution publique", "Name of public institution"),
    ("cfg_dept_label",      "Dienst / Entiteit",         "Service / Entité",          "Department / Entity"),
    ("cfg_lang_label",      "Taal / Langue / Language",  "Taal / Langue / Language",  "Taal / Langue / Language"),
    ("cfg_lang_hint",
        "Kies NL, FR of EN. Dropdowns en vertalingen passen zich automatisch aan.",
        "Choisissez NL, FR ou EN. Les listes et traductions s'adaptent automatiquement.",
        "Choose NL, FR or EN. Dropdowns and translations adapt automatically."),
    ("cfg_lang_change_hint",
        "Na taalwijziging: herselect bestaande classificatiewaarden in de gegevensbladen.",
        "Après changement de langue : re-sélectionnez les valeurs dans les feuilles de données.",
        "After language change: re-select existing classification values in data sheets."),
    ("cfg_version_label",    "Versie",          "Version",        "Version"),
    ("cfg_updated_label",    "Laatste update",  "Dernière mise à jour", "Last updated"),
    ("cfg_updated_by_label", "Bijgewerkt door", "Mis à jour par", "Updated by"),
    ("cfg_save_hint",
        "Sla het bestand op na elke wijziging (Ctrl+S).",
        "Enregistrez le fichier après chaque modification (Ctrl+S).",
        "Save the file after every change (Ctrl+S)."),
    # Info
    ("info_title",          "Over dit instrument",        "À propos de cet outil",     "About this tool"),
    ("info_section_doel",   "DOEL",                       "OBJECTIF",                  "PURPOSE"),
    ("info_doel_txt",
        "Deze GRC-tool ondersteunt overheidsorganisaties bij het systematisch beheren van informatieveiligheid. "
        "Processen en informatieassets worden geclassificeerd op integriteit, beschikbaarheid en confidentialiteit (5 niveaus: Laag → Kritiek).",
        "Cet outil GRC aide les organisations publiques à gérer systématiquement la sécurité de l'information. "
        "Les processus et actifs sont classifiés sur l'intégrité, la disponibilité et la confidentialité (5 niveaux : Faible → Critique).",
        "This GRC tool supports government organisations in systematically managing information security. "
        "Processes and assets are classified on integrity, availability and confidentiality (5 levels: Low → Critical)."),
    ("info_section_scope",  "TOEPASSINGSGEBIED",  "CHAMP D'APPLICATION", "SCOPE"),
    ("info_scope_txt",
        "Gericht op Belgische en Nederlandse overheidsdiensten (NIS2, GDPR, ISO 27001:2022, BIO).",
        "Destiné aux services publics belges et néerlandais (NIS2, RGPD, ISO 27001:2022, BIO).",
        "Aimed at Belgian and Dutch government services (NIS2, GDPR, ISO 27001:2022, BIO)."),
    ("info_section_modules", "MODULES IN DEZE VERSIE", "MODULES DANS CETTE VERSION", "MODULES IN THIS VERSION"),
    ("info_mod_proc",
        "Processen — Classificeer bedrijfsprocessen op I · B · C.",
        "Processus — Classifiez les processus métier sur I · D · C.",
        "Processes — Classify business processes on I · A · C."),
    ("info_mod_assets",
        "Informatieassets — Classificeer informatieactiva op I · B · C.",
        "Actifs informationnels — Classifiez les actifs sur I · D · C.",
        "Information Assets — Classify information assets on I · A · C."),
    ("info_mod_verant",
        "Verantwoordelijken — Contactregister van GRC-verantwoordelijken.",
        "Responsables — Registre de contacts des responsables GRC.",
        "Responsible Persons — Contact register of GRC responsible persons."),
    ("info_mod_import",
        "Import & Export — Macro's om gegevens te importeren uit externe bestanden.",
        "Import & Export — Macros pour importer des données depuis des fichiers externes.",
        "Import & Export — Macros to import data from external files."),
    ("info_section_guide",  "HOE GEBRUIKEN", "MODE D'EMPLOI", "HOW TO USE"),
    ("info_guide_1",
        "1. Configuratie: vul de naam van de instelling in en kies de taal (NL / FR / EN).",
        "1. Configuration : saisissez le nom et choisissez la langue (NL / FR / EN).",
        "1. Configuration: enter the name of the institution and choose the language (NL / FR / EN)."),
    ("info_guide_2",
        "2. Referentiewaarden: raadpleeg de definitie van elk classificatieniveau (1–5).",
        "2. Valeurs de référence : consultez la définition de chaque niveau de classification (1–5).",
        "2. Reference Values: consult the definition of each classification level (1–5)."),
    ("info_guide_3",
        "3. Processen & Informatieassets: vul de gegevens in. Kies niveau 1–5 per dimensie.",
        "3. Processus & Actifs : saisissez les données. Choisissez le niveau 1–5 par dimension.",
        "3. Processes & Assets: enter data. Choose level 1–5 per dimension."),
    ("info_guide_4",
        "4. Import & Export: gebruik de macro's om gegevens in bulk te importeren.",
        "4. Import & Export : utilisez les macros pour importer des données en masse.",
        "4. Import & Export: use the macros to import data in bulk."),
    ("info_section_version", "VERSIE-INFORMATIE", "INFORMATIONS DE VERSION", "VERSION INFORMATION"),
    ("info_v_tool",      "GRC Tool v0.3",                           "Outil GRC v0.3",                           "GRC Tool v0.3"),
    ("info_v_status",    "Status: Prototype / Fase 3",              "Statut : Prototype / Phase 3",             "Status: Prototype / Phase 3"),
    ("info_v_frameworks","Kaders: ISO 27001:2022 · NIS2 · GDPR · BIO",
                         "Cadres : ISO 27001:2022 · NIS2 · RGPD · BIO",
                         "Frameworks: ISO 27001:2022 · NIS2 · GDPR · BIO"),
    # Referentiewaarden
    ("ref_title",   "Referentiewaarden en Classificatieschalen",
                    "Valeurs de référence et échelles de classification",
                    "Reference Values and Classification Scales"),
    ("ref_subtitle",
        "Gebruik deze schalen bij het invullen van de classificatievelden (1 = Laag … 5 = Kritiek).",
        "Utilisez ces échelles lors du remplissage des champs de classification (1 = Faible … 5 = Critique).",
        "Use these scales when filling in classification fields (1 = Low … 5 = Critical)."),
    ("ref_level",       "Niveau",           "Niveau",           "Level"),
    ("ref_label",       "Benaming",         "Désignation",      "Designation"),
    ("ref_description", "Omschrijving",     "Description",      "Description"),
    ("ref_lvl1_label",  "Laag",             "Faible",           "Low"),
    ("ref_lvl2_label",  "Gemiddeld",        "Moyen",            "Medium"),
    ("ref_lvl3_label",  "Hoog",             "Élevé",            "High"),
    ("ref_lvl4_label",  "Zeer Hoog",        "Très élevé",       "Very High"),
    ("ref_lvl5_label",  "Classified",       "Classified",       "Classified"),
    ("ref_lvl1_desc",
        "Beperkte gevolgen. Verstoring of verlies heeft een verwaarloosbaar effect.",
        "Conséquences limitées. La perturbation ou la perte a un effet négligeable.",
        "Limited consequences. Disruption or loss has a negligible effect."),
    ("ref_lvl2_desc",
        "Merkbare gevolgen. Beperkt maar waarneembaar effect op de dienstverlening.",
        "Conséquences notables. Effet limité mais perceptible sur le service.",
        "Noticeable consequences. Limited but perceptible effect on service delivery."),
    ("ref_lvl3_desc",
        "Ernstige gevolgen. Significante schade aan de dienstverlening of betrokkenen.",
        "Conséquences graves. Dommages significatifs au service ou aux parties concernées.",
        "Serious consequences. Significant damage to service delivery or stakeholders."),
    ("ref_lvl4_desc",
        "Zeer ernstige gevolgen. Grote schade, gevaar voor personen of vitale functies.",
        "Conséquences très graves. Dommages importants, danger pour des personnes ou fonctions vitales.",
        "Very serious consequences. Major damage, danger to persons or vital functions."),
    ("ref_lvl5_desc",
        "Kritieke gevolgen. Onherstelbare schade, gevaar voor nationale veiligheid of grondrechten, of volledige uitval van vitale overheidsfuncties.",
        "Conséquences critiques. Dommages irréparables, danger pour la sécurité nationale ou les droits fondamentaux, ou arrêt total des fonctions gouvernementales vitales.",
        "Critical consequences. Irreparable damage, danger to national security or fundamental rights, or complete failure of vital government functions."),
    ("ref_section_int",   "INTEGRITEIT",      "INTÉGRITÉ",        "INTEGRITY"),
    ("ref_int_subtitle",  "Mate van correctheid, volledigheid en betrouwbaarheid van gegevens.",
                          "Degré de correction, d'exhaustivité et de fiabilité des données.",
                          "Degree of correctness, completeness and reliability of data."),
    ("ref_int_1",  "Fouten hebben geen impact op beslissingen of dienstverlening.",
                   "Les erreurs n'ont aucun impact sur les décisions ou le service.",
                   "Errors have no impact on decisions or service delivery."),
    ("ref_int_2",  "Fouten kunnen leiden tot beperkte fouten in beslissingen.",
                   "Les erreurs peuvent entraîner des erreurs limitées dans les décisions.",
                   "Errors may lead to limited mistakes in decisions."),
    ("ref_int_3",  "Fouten leiden tot verkeerde beslissingen met significante gevolgen.",
                   "Les erreurs entraînent des décisions erronées avec des conséquences significatives.",
                   "Errors lead to wrong decisions with significant consequences."),
    ("ref_int_4",  "Fouten leiden tot grove schade: gevaar, juridische gevolgen, schending van rechten.",
                   "Les erreurs entraînent des dommages graves : danger, conséquences juridiques, violation des droits.",
                   "Errors lead to serious damage: danger, legal consequences, violation of rights."),
    ("ref_int_5",  "Fouten kunnen leiden tot onherstelbare schade, gevaar voor nationale veiligheid of fundamentele schending van grondrechten.",
                   "Les erreurs peuvent entraîner des dommages irréparables, un danger pour la sécurité nationale ou une violation fondamentale des droits.",
                   "Errors may lead to irreparable damage, danger to national security or fundamental violation of rights."),
    ("ref_section_avail", "BESCHIKBAARHEID",  "DISPONIBILITÉ",    "AVAILABILITY"),
    ("ref_avail_subtitle","Mate van toegankelijkheid wanneer dat nodig is.",
                          "Degré de disponibilité lorsque nécessaire.",
                          "Degree of accessibility when needed."),
    ("ref_avail_1","Uitval heeft geen merkbare impact. Processen lopen handmatig verder.",
                   "La panne n'a pas d'impact. Les processus continuent manuellement.",
                   "Outage has no noticeable impact. Processes continue manually."),
    ("ref_avail_2","Uitval veroorzaakt hinder maar geen kritieke verstoring.",
                   "La panne cause des inconvénients mais pas de perturbation critique.",
                   "Outage causes inconvenience but no critical disruption."),
    ("ref_avail_3","Uitval leidt tot significante verstoring van de dienstverlening.",
                   "La panne entraîne une perturbation significative du service.",
                   "Outage leads to significant disruption of services."),
    ("ref_avail_4","Uitval veroorzaakt grote schade of stillegging van belangrijke diensten.",
                   "La panne provoque des dommages importants ou l'arrêt de services importants.",
                   "Outage causes major damage or shutdown of important services."),
    ("ref_avail_5","Uitval leidt tot volledige stopzetting van vitale overheidsfuncties of noodsituaties met gevaar voor leven.",
                   "La panne entraîne l'arrêt total des fonctions gouvernementales vitales ou des situations d'urgence mettant des vies en danger.",
                   "Outage leads to complete failure of vital government functions or emergencies endangering lives."),
    ("ref_section_conf",  "CONFIDENTIALITEIT","CONFIDENTIALITÉ",  "CONFIDENTIALITY"),
    ("ref_conf_subtitle", "Toegankelijkheid enkel voor bevoegde personen of systemen.",
                          "Accessibilité uniquement aux personnes ou systèmes autorisés.",
                          "Accessibility only to authorised persons or systems."),
    ("ref_conf_1", "Openbare gegevens. Verspreiding heeft geen nadelige gevolgen.",
                   "Données publiques. La diffusion n'a pas de conséquences négatives.",
                   "Public data. Disclosure has no adverse consequences."),
    ("ref_conf_2", "Interne gegevens. Verspreiding kan leiden tot beperkte reputatieschade.",
                   "Données internes. La diffusion peut entraîner des dommages à la réputation.",
                   "Internal data. Disclosure may lead to limited reputational damage."),
    ("ref_conf_3", "Vertrouwelijke gegevens. Ongeoorloofde verspreiding heeft ernstige gevolgen.",
                   "Données confidentielles. La diffusion non autorisée a de graves conséquences.",
                   "Confidential data. Unauthorised disclosure has serious consequences."),
    ("ref_conf_4", "Strikt vertrouwelijke gegevens. Verspreiding bedreigt personen, organisaties of vitale processen.",
                   "Données strictement confidentielles. La diffusion menace des personnes, organisations ou processus vitaux.",
                   "Strictly confidential data. Disclosure threatens persons, organisations or vital processes."),
    ("ref_conf_5", "Geheime gegevens. Verspreiding kan nationale veiligheid, grondrechten of vitale staatsfuncties schaden.",
                   "Données secrètes. La diffusion peut nuire à la sécurité nationale, aux droits fondamentaux ou aux fonctions vitales de l'État.",
                   "Secret data. Disclosure may harm national security, fundamental rights or vital state functions."),
    # Processen
    ("proc_title",          "Processen",         "Processus",        "Processes"),
    ("proc_subtitle",
        "Voer de bedrijfsprocessen in en ken de veiligheidclassificatie toe (1 = Laag … 5 = Kritiek).",
        "Saisissez les processus et attribuez la classification de sécurité (1 = Faible … 5 = Critique).",
        "Enter the business processes and assign the security classification (1 = Low … 5 = Critical)."),
    ("proc_col_id",           "ID",                "ID",               "ID"),
    ("proc_col_naam",         "Naam",              "Nom",              "Name"),
    ("proc_col_omschrijving", "Omschrijving",      "Description",      "Description"),
    ("proc_col_eigenaar",     "Proceseigenaar",    "Propriétaire",     "Process Owner"),
    ("proc_col_dienst",       "Dienst / Entiteit", "Service / Entité", "Department / Entity"),
    ("proc_col_integriteit",  "Integriteit",       "Intégrité",        "Integrity"),
    ("proc_col_beschikbaar",  "Beschikbaarheid",   "Disponibilité",    "Availability"),
    ("proc_col_conf",         "Confidentialiteit", "Confidentialité",  "Confidentiality"),
    ("proc_col_assets",       "Informatieassets",  "Actifs informationnels", "Information Assets"),
    ("proc_col_opmerkingen",  "Opmerkingen",       "Remarques",        "Remarks"),
    # Informatieassets
    ("asset_title",           "Informatieassets",  "Actifs informationnels", "Information Assets"),
    ("asset_subtitle",
        "Voer de informatieactiva in en ken de veiligheidclassificatie toe (1 = Laag … 5 = Kritiek).",
        "Saisissez les actifs et attribuez la classification de sécurité (1 = Faible … 5 = Critique).",
        "Enter the information assets and assign the security classification (1 = Low … 5 = Critical)."),
    ("asset_col_id",          "ID",                "ID",               "ID"),
    ("asset_col_naam",        "Naam",              "Nom",              "Name"),
    ("asset_col_type",        "Type",              "Type",             "Type"),
    ("asset_col_omschrijving","Omschrijving",      "Description",      "Description"),
    ("asset_col_eigenaar",    "Eigenaar",          "Propriétaire",     "Owner"),
    ("asset_col_dienst",      "Dienst / Entiteit", "Service / Entité", "Department / Entity"),
    ("asset_col_integriteit", "Integriteit",       "Intégrité",        "Integrity"),
    ("asset_col_beschikbaar", "Beschikbaarheid",   "Disponibilité",    "Availability"),
    ("asset_col_conf",        "Confidentialiteit", "Confidentialité",  "Confidentiality"),
    ("asset_col_opmerkingen", "Opmerkingen",       "Remarques",        "Remarks"),
    # Verantwoordelijken
    ("verant_title",           "Verantwoordelijken", "Responsables",    "Responsible Persons"),
    ("verant_subtitle",
        "Overzicht van contactpersonen en hun GRC-rol.",
        "Aperçu des personnes de contact et de leur rôle GRC.",
        "Overview of contact persons and their GRC role."),
    ("verant_col_id",          "ID",                "ID",               "ID"),
    ("verant_col_naam",        "Naam",              "Nom",              "Name"),
    ("verant_col_voornaam",    "Voornaam",          "Prénom",           "First Name"),
    ("verant_col_functie",     "Functietitel",      "Titre de fonction","Job Title"),
    ("verant_col_dienst",      "Dienst / Entiteit", "Service / Entité", "Department / Entity"),
    ("verant_col_email",       "E-mailadres",       "Adresse e-mail",   "Email Address"),
    ("verant_col_tel",         "Telefoon",          "Téléphone",        "Phone"),
    ("verant_col_rol",         "Rol in GRC",        "Rôle dans le GRC", "Role in GRC"),
    ("verant_col_opmerkingen", "Opmerkingen",       "Remarques",        "Remarks"),
    # Import & Export
    ("ie_title",     "Import & Export",  "Import & Export",  "Import & Export"),
    ("ie_subtitle",
        "Importeer gegevens vanuit externe bronnen of exporteer het GRC-register.",
        "Importez des données depuis des sources externes ou exportez le registre GRC.",
        "Import data from external sources or export the GRC register."),
    ("ie_section_import",  "IMPORTEREN",    "IMPORTATION",    "IMPORT"),
    ("ie_section_export",  "EXPORTEREN",    "EXPORTATION",    "EXPORT"),
    ("ie_section_info",    "TABELSTRUCTUUR BRONBESTAND", "STRUCTURE DU FICHIER SOURCE", "SOURCE FILE TABLE STRUCTURE"),
    ("ie_import_all_btn",
        "Importeer Alles",
        "Importer tout",
        "Import All"),
    ("ie_import_all_desc",
        "Importeert informatieassets, afhankelijke assets, processen én kwetsbaarheden-matrix in één stap.",
        "Importe actifs informationnels, actifs dépendants, processus et matrice des vulnérabilités en une seule étape.",
        "Imports information assets, dependent assets, processes and vulnerability matrix in one step."),
    ("ie_import_ia_btn",
        "Importeer Informatieassets",
        "Importer actifs informationnels",
        "Import Information Assets"),
    ("ie_import_ia_desc",
        "Leest de tabel 'T - Information Assets' uit de Access-database en vult de tab Informatieassets.",
        "Lit la table 'T - Information Assets' depuis la base Access et remplit l'onglet Actifs informationnels.",
        "Reads the 'T - Information Assets' table from the Access database and fills the Information Assets tab."),
    ("ie_import_da_btn",
        "Importeer Afhankelijke Assets",
        "Importer actifs dépendants",
        "Import Dependent Assets"),
    ("ie_import_da_desc",
        "Leest de tabel 'T - Dependent assets' en vult de tab Afhankelijke Assets inclusief CIA-objectives.",
        "Lit la table 'T - Dependent assets' et remplit l'onglet Actifs dépendants avec les objectifs CIA.",
        "Reads the 'T - Dependent assets' table and fills the Dependent Assets tab including CIA objectives."),
    ("ie_import_proc_btn",
        "Importeer Processen",
        "Importer processus",
        "Import Processes"),
    ("ie_import_proc_desc",
        "Leest 'T - Processes in scope' en vult de tab Processen inclusief koppelingen aan assets.",
        "Lit 'T - Processes in scope' et remplit l'onglet Processus avec les liens aux actifs.",
        "Reads 'T - Processes in scope' and fills the Processes tab including asset links."),
    ("ie_import_kwets_btn",
        "Importeer Kwetsbaarheden",
        "Importer vulnérabilités",
        "Import Vulnerabilities"),
    ("ie_import_kwets_desc",
        "Laadt de kwetsbaarheden en hun koppeling aan CyFun-controls. Vereist dat de tab Kwetsbaarheden al aangemaakt is.",
        "Charge les vulnérabilités et leur lien aux contrôles CyFun. Nécessite que l'onglet Vulnérabilités soit déjà créé.",
        "Loads vulnerabilities and their link to CyFun controls. Requires the Vulnerabilities tab to already exist."),
    ("ie_import_links_btn",
        "Importeer Links DA / Kwetsbaarheden",
        "Importer liens actifs dépendants / vulnérabilités",
        "Import DA / Vulnerability Links"),
    ("ie_import_links_desc",
        "Voer dit uit NA de import van kwetsbaarheden. Leest de koppelingen tussen afhankelijke assets en kwetsbaarheden uit de Access-database en vult de RARM-tab aan.",
        "À exécuter APRÈS l'import des vulnérabilités. Lit les liens entre actifs dépendants et vulnérabilités depuis la base Access et complète l'onglet RARM.",
        "Run this AFTER importing vulnerabilities. Reads links between dependent assets and vulnerabilities from the Access database and fills in the RARM tab."),
    ("ie_export_btn",      "Exporteer alle gegevens",     "Exporter toutes les données",   "Export all data"),
    ("ie_export_desc",
        "Exporteert de bladen Processen, Informatieassets en Verantwoordelijken naar een nieuw Excel-bestand.",
        "Exporte les feuilles Processus, Actifs et Responsables vers un nouveau fichier Excel.",
        "Exports the Processes, Information Assets and Responsible Persons sheets to a new Excel file."),
    ("ie_table_proc",  "T - Processes in scope",    "T - Processes in scope",    "T - Processes in scope"),
    ("ie_table_asset", "T - Information Assets",    "T - Information Assets",    "T - Information Assets"),
    ("ie_table_dep",   "T - Dependent assets",      "T - Dependent assets",      "T - Dependent assets"),
    # Afhankelijke assets sheet
    ("dep_title",    "Afhankelijke Assets",
                     "Actifs dépendants",
                     "Dependent Assets"),
    ("dep_subtitle",
        "Overzicht van systemen, infrastructuur en diensten waarvan processen afhankelijk zijn",
        "Vue d'ensemble des systèmes, infrastructures et services dont dépendent les processus",
        "Overview of systems, infrastructure and services that processes depend on"),
    ("dep_col_id",           "ID",                  "ID",                      "ID"),
    ("dep_col_naam",         "Naam",                "Nom",                     "Name"),
    ("dep_col_omschrijving", "Omschrijving",        "Description",             "Description"),
    ("dep_col_eigenaar",     "Eigenaar",            "Propriétaire",            "Owner"),
    ("dep_col_dienst",       "Dienst / Entiteit",   "Service / Entité",        "Department / Entity"),
    ("dep_col_overarching",  "Overarching",         "Transversal",             "Overarching"),
    ("dep_col_opmerkingen",  "Opmerkingen",         "Remarques",               "Remarks"),
    ("dep_col_processes",
        "Gebruikt door processen",
        "Utilisé par les processus",
        "Used by Processes"),
    ("dep_col_linked_info",
        "Informatieassets (via processen)",
        "Actifs informationnels (via processus)",
        "Information Assets (via processes)"),
    ("dep_sec_req_header",
        "Security Requirements",
        "Exigences de sécurité",
        "Security Requirements"),
    ("dep_col_conf",
        "Confidentialiteit",
        "Confidentialité",
        "Confidentiality"),
    ("dep_col_integ",
        "Integriteit",
        "Intégrité",
        "Integrity"),
    ("dep_col_avail",
        "Beschikbaarheid",
        "Disponibilité",
        "Availability"),
    ("dep_sec_obj_header",
        "Security Objectives",
        "Objectifs de sécurité",
        "Security Objectives"),
    ("dep_col_obj_commentaar",
        "Commentaar",
        "Commentaire",
        "Comment"),
    ("dep_gap_header",
        "Gap Analyse",
        "Analyse des écarts",
        "Gap Analysis"),
    ("dep_col_gap_conf",
        "Conf. niet gedekt (info assets)",
        "Conf. non couverte (actifs info)",
        "Conf. not covered (info assets)"),
    ("dep_col_gap_integ",
        "Int. niet gedekt (processen)",
        "Int. non couverte (processus)",
        "Int. not covered (processes)"),
    ("dep_col_gap_avail",
        "Beschikb. niet gedekt (processen)",
        "Disp. non couverte (processus)",
        "Avail. not covered (processes)"),
    # Extra kolom in Processen
    ("proc_col_dep_assets",  "Afhankelijke assets", "Actifs dépendants",       "Dependent Assets"),
    # Informatieassets extra kolom
    ("asset_col_processes",
        "Gebruikt in processen",
        "Utilisé dans les processus",
        "Used in Processes"),
    # Navigatie
    ("ui_goto_dep",          "→  Dependent Assets",    "→  Actifs dépendants", "→  Dependent Assets"),
    ("ie_col_hint",
        "Ondersteunde kolomnamen (niet hoofdlettergevoelig):",
        "Noms de colonnes pris en charge (insensibles à la casse) :",
        "Supported column names (case-insensitive):"),
    ("ie_macro_hint",
        "De macro's zijn ingebouwd in dit bestand (Module: GRC_Macros). Klik op de knoppen hierboven om te starten.",
        "Les macros sont intégrées dans ce fichier (Module : GRC_Macros). Cliquez sur les boutons ci-dessus pour démarrer.",
        "The macros are built into this file (Module: GRC_Macros). Click the buttons above to start."),
]

# ── VBA-macrocode ─────────────────────────────────────────────────────────────
KWETS_SHEET_CODE = '''\
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 3 Or Target.Column < 3 Then Exit Sub
    Cancel = True
    Dim sVal As String
    sVal = ""
    If Not IsError(Target.Value) Then sVal = CStr(Target.Value)
    If sVal = ChrW(10004) Then
        Target.Value = ""
    Else
        Target.Value = ChrW(10004)
        Target.HorizontalAlignment = xlCenter
        Target.VerticalAlignment = xlCenter
    End If
End Sub
'''

INFO_ASSET_SHEET_CODE = '''\
Private Sub Worksheet_Activate()
    Dim procWs As Worksheet
    Dim r As Long, c As Long
    Dim assetNaam As String, pNaam As String, procList As String
    Application.ScreenUpdating = False
    Set procWs = ThisWorkbook.Sheets("Processes")
    Me.Range("I6:I105").ClearContents
    For r = 6 To 105
        assetNaam = CStr(Me.Cells(r, 2).Value)
        If assetNaam <> "" Then
            procList = ""
            For c = 6 To 105
                pNaam = CStr(procWs.Cells(c, 2).Value)
                If pNaam <> "" Then
                    If InStr(LCase(CStr(procWs.Cells(c, 11).Value)), LCase(assetNaam)) > 0 Then
                        If procList <> "" Then procList = procList & Chr(10)
                        procList = procList & pNaam
                    End If
                End If
            Next c
            Me.Cells(r, 9) = procList
            If procList <> "" Then Me.Rows(r).AutoFit
        End If
    Next r
    Application.ScreenUpdating = True
End Sub
'''

DEP_ASSET_SHEET_CODE = '''\
Private Sub Worksheet_Activate()
    Dim procWs As Worksheet, iaWs As Worksheet
    Dim r As Long, c As Long, p As Long, ia As Long
    Dim depNaam As String, pNaam As String, iaName As String
    Dim procList As String, infoList As String, iaStr As String
    Dim parts() As String
    Dim maxConf As Integer, maxInt As Integer, maxAvail As Integer
    Dim objConf As Integer, objInt As Integer, objAvail As Integer
    Dim gapConf As String, gapInt As String, gapAvail As String
    Dim v As Variant
    Application.ScreenUpdating = False
    Me.Unprotect
    Set procWs = ThisWorkbook.Sheets("Processes")
    Set iaWs = ThisWorkbook.Sheets("Information Assets")
    Me.Range("H6:J105").ClearContents
    Me.Range("L6:L105").ClearContents
    Me.Range("N6:N105").ClearContents
    Me.Range("V6:X105").ClearContents
    For r = 6 To 105
        depNaam = CStr(Me.Cells(r, 2).Value)
        If depNaam <> "" Then
            procList = "": infoList = ""
            maxConf = 0: maxInt = 0: maxAvail = 0
            For c = 6 To 105
                pNaam = CStr(procWs.Cells(c, 2).Value)
                If pNaam <> "" Then
                    If InStr(LCase(CStr(procWs.Cells(c, 12).Value)), LCase(depNaam)) > 0 Then
                        If procList <> "" Then procList = procList & Chr(10)
                        procList = procList & pNaam
                        v = procWs.Cells(c, 6).Value
                        If IsNumeric(v) And CInt(v) > maxInt Then maxInt = CInt(v)
                        v = procWs.Cells(c, 8).Value
                        If IsNumeric(v) And CInt(v) > maxAvail Then maxAvail = CInt(v)
                        iaStr = CStr(procWs.Cells(c, 11).Value)
                        If iaStr <> "" Then
                            parts = Split(iaStr, Chr(10))
                            For p = 0 To UBound(parts)
                                iaName = Trim(parts(p))
                                If iaName <> "" Then
                                    If InStr(Chr(10) & infoList & Chr(10), Chr(10) & iaName & Chr(10)) = 0 Then
                                        If infoList <> "" Then infoList = infoList & Chr(10)
                                        infoList = infoList & iaName
                                    End If
                                    For ia = 6 To 105
                                        If LCase(CStr(iaWs.Cells(ia, 2).Value)) = LCase(iaName) Then
                                            v = iaWs.Cells(ia, 6).Value
                                            If IsNumeric(v) And CInt(v) > maxConf Then maxConf = CInt(v)
                                            Exit For
                                        End If
                                    Next ia
                                End If
                            Next p
                        End If
                    End If
                End If
            Next c
            Me.Cells(r, 8) = procList
            Me.Cells(r, 9) = infoList
            If maxConf > 0 Then Me.Cells(r, 10) = maxConf
            If maxInt > 0 Then Me.Cells(r, 12) = maxInt
            If maxAvail > 0 Then Me.Cells(r, 14) = maxAvail
            objConf = 0: objInt = 0: objAvail = 0
            If IsNumeric(Me.Cells(r, 16).Value) Then objConf = CInt(Me.Cells(r, 16).Value)
            If IsNumeric(Me.Cells(r, 18).Value) Then objInt = CInt(Me.Cells(r, 18).Value)
            If IsNumeric(Me.Cells(r, 20).Value) Then objAvail = CInt(Me.Cells(r, 20).Value)
            gapConf = ""
            If objConf > 0 And maxConf > objConf And infoList <> "" Then
                parts = Split(infoList, Chr(10))
                For p = 0 To UBound(parts)
                    iaName = Trim(parts(p))
                    If iaName <> "" Then
                        For ia = 6 To 105
                            If LCase(CStr(iaWs.Cells(ia, 2).Value)) = LCase(iaName) Then
                                v = iaWs.Cells(ia, 6).Value
                                If IsNumeric(v) And CInt(v) > objConf Then
                                    If gapConf <> "" Then gapConf = gapConf & Chr(10)
                                    gapConf = gapConf & iaName
                                End If
                                Exit For
                            End If
                        Next ia
                    End If
                Next p
            End If
            Me.Cells(r, 22) = gapConf
            gapInt = ""
            If objInt > 0 And maxInt > objInt And procList <> "" Then
                parts = Split(procList, Chr(10))
                For p = 0 To UBound(parts)
                    pNaam = Trim(parts(p))
                    If pNaam <> "" Then
                        For c = 6 To 105
                            If LCase(CStr(procWs.Cells(c, 2).Value)) = LCase(pNaam) Then
                                v = procWs.Cells(c, 6).Value
                                If IsNumeric(v) And CInt(v) > objInt Then
                                    If gapInt <> "" Then gapInt = gapInt & Chr(10)
                                    gapInt = gapInt & pNaam
                                End If
                                Exit For
                            End If
                        Next c
                    End If
                Next p
            End If
            Me.Cells(r, 23) = gapInt
            gapAvail = ""
            If objAvail > 0 And maxAvail > objAvail And procList <> "" Then
                parts = Split(procList, Chr(10))
                For p = 0 To UBound(parts)
                    pNaam = Trim(parts(p))
                    If pNaam <> "" Then
                        For c = 6 To 105
                            If LCase(CStr(procWs.Cells(c, 2).Value)) = LCase(pNaam) Then
                                v = procWs.Cells(c, 8).Value
                                If IsNumeric(v) And CInt(v) > objAvail Then
                                    If gapAvail <> "" Then gapAvail = gapAvail & Chr(10)
                                    gapAvail = gapAvail & pNaam
                                End If
                                Exit For
                            End If
                        Next c
                    End If
                Next p
            End If
            Me.Cells(r, 24) = gapAvail
            If procList <> "" Or infoList <> "" Then Me.Rows(r).AutoFit
        End If
    Next r
    Me.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 6 Or Target.Row > 105 Then Exit Sub
    If Target.Column <> 6 Then Exit Sub
    Cancel = True
    If CStr(Target.Value) = ChrW(10004) Then
        Target.Value = ""
    Else
        Target.Value = ChrW(10004)
        Target.HorizontalAlignment = xlCenter
        Target.VerticalAlignment = xlCenter
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    Dim hasObjChange As Boolean
    hasObjChange = False
    For Each cell In Target.Cells
        If cell.Row >= 6 And cell.Row <= 105 Then
            If cell.Column = 16 Or cell.Column = 18 Or cell.Column = 20 Then
                hasObjChange = True
                Exit For
            End If
        End If
    Next cell
    If Not hasObjChange Then Exit Sub
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Me.Unprotect
    Dim procWs As Worksheet, iaWs As Worksheet
    Dim r As Long, c As Long, p As Long, ia As Long
    Dim pNaam As String, iaName As String
    Dim procList As String, infoList As String
    Dim parts() As String
    Dim v As Variant
    Dim maxConf As Integer, maxInt As Integer, maxAvail As Integer
    Dim objConf As Integer, objInt As Integer, objAvail As Integer
    Dim gapConf As String, gapInt As String, gapAvail As String
    Dim rowDone(6 To 105) As Boolean
    Set procWs = ThisWorkbook.Sheets("Processes")
    Set iaWs = ThisWorkbook.Sheets("Information Assets")
    For Each cell In Target.Cells
        r = cell.Row
        If r >= 6 And r <= 105 Then
            If (cell.Column = 16 Or cell.Column = 18 Or cell.Column = 20) And Not rowDone(r) Then
                rowDone(r) = True
                procList = CStr(Me.Cells(r, 8).Value)
                infoList = CStr(Me.Cells(r, 9).Value)
                maxConf = 0: maxInt = 0: maxAvail = 0
                If IsNumeric(Me.Cells(r, 10).Value) Then maxConf = CInt(Me.Cells(r, 10).Value)
                If IsNumeric(Me.Cells(r, 12).Value) Then maxInt = CInt(Me.Cells(r, 12).Value)
                If IsNumeric(Me.Cells(r, 14).Value) Then maxAvail = CInt(Me.Cells(r, 14).Value)
                objConf = 0: objInt = 0: objAvail = 0
                If IsNumeric(Me.Cells(r, 16).Value) Then objConf = CInt(Me.Cells(r, 16).Value)
                If IsNumeric(Me.Cells(r, 18).Value) Then objInt = CInt(Me.Cells(r, 18).Value)
                If IsNumeric(Me.Cells(r, 20).Value) Then objAvail = CInt(Me.Cells(r, 20).Value)
                gapConf = ""
                If objConf > 0 And maxConf > objConf And infoList <> "" Then
                    parts = Split(infoList, Chr(10))
                    For p = 0 To UBound(parts)
                        iaName = Trim(parts(p))
                        If iaName <> "" Then
                            For ia = 6 To 105
                                If LCase(CStr(iaWs.Cells(ia, 2).Value)) = LCase(iaName) Then
                                    v = iaWs.Cells(ia, 6).Value
                                    If IsNumeric(v) And CInt(v) > objConf Then
                                        If gapConf <> "" Then gapConf = gapConf & Chr(10)
                                        gapConf = gapConf & iaName
                                    End If
                                    Exit For
                                End If
                            Next ia
                        End If
                    Next p
                End If
                Me.Cells(r, 22) = gapConf
                gapInt = ""
                If objInt > 0 And maxInt > objInt And procList <> "" Then
                    parts = Split(procList, Chr(10))
                    For p = 0 To UBound(parts)
                        pNaam = Trim(parts(p))
                        If pNaam <> "" Then
                            For c = 6 To 105
                                If LCase(CStr(procWs.Cells(c, 2).Value)) = LCase(pNaam) Then
                                    v = procWs.Cells(c, 6).Value
                                    If IsNumeric(v) And CInt(v) > objInt Then
                                        If gapInt <> "" Then gapInt = gapInt & Chr(10)
                                        gapInt = gapInt & pNaam
                                    End If
                                    Exit For
                                End If
                            Next c
                        End If
                    Next p
                End If
                Me.Cells(r, 23) = gapInt
                gapAvail = ""
                If objAvail > 0 And maxAvail > objAvail And procList <> "" Then
                    parts = Split(procList, Chr(10))
                    For p = 0 To UBound(parts)
                        pNaam = Trim(parts(p))
                        If pNaam <> "" Then
                            For c = 6 To 105
                                If LCase(CStr(procWs.Cells(c, 2).Value)) = LCase(pNaam) Then
                                    v = procWs.Cells(c, 8).Value
                                    If IsNumeric(v) And CInt(v) > objAvail Then
                                        If gapAvail <> "" Then gapAvail = gapAvail & Chr(10)
                                        gapAvail = gapAvail & pNaam
                                    End If
                                    Exit For
                                End If
                            Next c
                        End If
                    Next p
                End If
                Me.Cells(r, 24) = gapAvail
                Me.Rows(r).AutoFit
            End If
        End If
    Next cell
    Me.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
'''

USERFORM_CODE = '''\
Private Sub UserForm_Initialize()
    Me.Caption = "Assets voor: " & GRC_Macros.g_ProcesNaam
    lstAssets.Clear
    Dim assetWs As Worksheet
    Set assetWs = ThisWorkbook.Sheets(GRC_Macros.g_PickerSourceSheet)
    Dim huidigStr As String
    huidigStr = LCase(CStr(ThisWorkbook.Sheets("Processes").Cells(GRC_Macros.g_TargetRow, GRC_Macros.g_PickerTargetCol).Value))
    Dim r As Long
    For r = 6 To 105
        Dim assetNaam As String
        assetNaam = CStr(assetWs.Cells(r, 2).Value)
        If assetNaam <> "" Then
            lstAssets.AddItem assetNaam
            Dim idx As Long: idx = lstAssets.ListCount - 1
            Dim lines() As String
            lines = Split(huidigStr, Chr(10))
            Dim li As Long, isSelected As Boolean
            isSelected = False
            For li = 0 To UBound(lines)
                If Trim(lines(li)) = LCase(Trim(assetNaam)) Then isSelected = True: Exit For
            Next li
            If isSelected Then lstAssets.Selected(idx) = True
        End If
    Next r
End Sub

Private Sub btnOK_Click()
    Dim result As String: result = ""
    Dim i As Long
    For i = 0 To lstAssets.ListCount - 1
        If lstAssets.Selected(i) Then
            If result <> "" Then result = result & Chr(10)
            result = result & lstAssets.List(i)
        End If
    Next i
    Dim procWs As Worksheet
    Set procWs = ThisWorkbook.Sheets("Processes")
    procWs.Cells(GRC_Macros.g_TargetRow, GRC_Macros.g_PickerTargetCol) = result
    procWs.Rows(GRC_Macros.g_TargetRow).AutoFit
    Unload Me
    procWs.Activate
End Sub

Private Sub btnAnnuleer_Click()
    Unload Me
    ThisWorkbook.Sheets("Processes").Activate
End Sub
'''

RARM_SHEET_CODE = '''\
Private Sub Worksheet_Activate()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    GRC_Macros.SyncRARMKolommen
    GRC_Macros.KleurAlleRARMKolommen
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 4 Then Exit Sub
    If Target.Column < 4 Then Exit Sub
    If CStr(Me.Cells(Target.Row, 1).Value) = "" Then Exit Sub
    Cancel = True
    If CStr(Target.Value) = ChrW(10004) Then
        Target.Value = ""
    Else
        Target.Value = ChrW(10004)
        Target.HorizontalAlignment = xlCenter
        Target.VerticalAlignment = xlCenter
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row <> 3 Then Exit Sub
    If Target.Column < 5 Then Exit Sub
    Application.EnableEvents = False
    GRC_Macros.TonenVulnPicker Target.Column
    Application.EnableEvents = True
End Sub
'''

VULNPICKER_CODE = '''\
Private Sub UserForm_Initialize()
    Dim wsK As Worksheet, wsR As Worksheet
    On Error Resume Next
    Set wsK = ThisWorkbook.Sheets("Kwetsbaarheden")
    Set wsR = ThisWorkbook.Sheets("RARM")
    On Error GoTo 0
    If wsK Is Nothing Then Exit Sub

    Dim daName As String
    daName = ""
    If Not wsR Is Nothing Then
        daName = CStr(wsR.Cells(2, GRC_Macros.g_RARMCol).Value)
    End If
    Me.Caption = IIf(daName <> "", "Kwetsbaarheden voor: " & daName, "Kwetsbaarheden selecteren")
    lblTitel.Caption = Me.Caption

    ' Parse existing selection into Dictionary: vname -> probability (1-4)
    ' Storage format: "VulnName (2), VulnName2 (3)"
    Dim selDict As Object
    Set selDict = CreateObject("Scripting.Dictionary")
    If Not wsR Is Nothing Then
        Dim curVal As String
        curVal = CStr(wsR.Cells(3, GRC_Macros.g_RARMCol).Value)
        If curVal <> "" Then
            Dim parts() As String
            parts = Split(curVal, ", ")
            Dim pi As Integer
            For pi = 0 To UBound(parts)
                Dim entry As String
                entry = Trim(parts(pi))
                If entry <> "" Then
                    Dim parenOpen As Integer
                    parenOpen = InStrRev(entry, " (")
                    If parenOpen > 0 Then
                        Dim vKey As String, pNum As String
                        vKey = Left(entry, parenOpen - 1)
                        pNum = Mid(entry, parenOpen + 2, Len(entry) - parenOpen - 2)
                        selDict(vKey) = pNum
                    Else
                        selDict(entry) = "2"
                    End If
                End If
            Next pi
        End If
    End If

    ' Dynamically add CheckBox + ComboBox (1-4) rows to frmVulns
    Dim rowTop As Long
    rowTop = 4
    Dim col As Long
    col = 3
    Do While wsK.Cells(2, col).Value <> ""
        Dim vName As String
        vName = CStr(wsK.Cells(2, col).Value)

        Dim chk As Object
        Set chk = frmVulns.Controls.Add("Forms.CheckBox.1")
        chk.Name = "chk_" & col
        chk.Caption = vName
        chk.Tag = vName
        chk.Left = 4
        chk.Top = rowTop
        chk.Width = 220
        chk.Height = 18
        chk.Font.Size = 9

        Dim cmb As Object
        Set cmb = frmVulns.Controls.Add("Forms.ComboBox.1")
        cmb.Name = "cmb_" & col
        cmb.Tag = vName
        cmb.Left = 228
        cmb.Top = rowTop
        cmb.Width = 162
        cmb.Height = 18
        cmb.Font.Size = 9
        cmb.Style = 2   ' fmStyleDropDownList
        cmb.AddItem "1 - Not probable"
        cmb.AddItem "2 - Low probability"
        cmb.AddItem "3 - Medium probability"
        cmb.AddItem "4 - High probability"
        cmb.ListIndex = 1  ' default = 2 (Low probability)

        Dim isSelected As Boolean
        isSelected = False
        If selDict.Exists(vName) Then
            isSelected = True
            Dim savedProb As String
            savedProb = CStr(selDict(vName))
            Dim savedIdx As Integer
            savedIdx = CInt(savedProb) - 1
            If savedIdx >= 0 And savedIdx <= 3 Then cmb.ListIndex = savedIdx
        End If
        chk.Value = isSelected

        rowTop = rowTop + 22
        col = col + 1
    Loop

    frmVulns.ScrollHeight = rowTop + 4
End Sub

Private Sub btnOK_Click()
    Dim wsR As Worksheet, wsK As Worksheet
    On Error Resume Next
    Set wsR = ThisWorkbook.Sheets("RARM")
    Set wsK = ThisWorkbook.Sheets("Kwetsbaarheden")
    On Error GoTo 0
    If wsR Is Nothing Then Unload Me: Exit Sub

    Dim result As String
    result = ""
    Dim col As Long
    col = 3
    Do While wsK.Cells(2, col).Value <> ""
        Dim chk As Object, cmb As Object
        Set chk = Nothing: Set cmb = Nothing
        On Error Resume Next
        Set chk = frmVulns.Controls("chk_" & col)
        Set cmb = frmVulns.Controls("cmb_" & col)
        On Error GoTo 0
        If Not chk Is Nothing Then
            If chk.Value = True Then
                Dim probNum As String
                probNum = "2"
                If Not cmb Is Nothing Then
                    If cmb.ListIndex >= 0 Then probNum = CStr(cmb.ListIndex + 1)
                End If
                If result <> "" Then result = result & ", "
                result = result & chk.Tag & " (" & probNum & ")"
            End If
        End If
        col = col + 1
    Loop
    wsR.Cells(3, GRC_Macros.g_RARMCol).Value = result
    Unload Me
    GRC_Macros.KleurAlleRARMKolommen
End Sub

Private Sub btnAnnuleer_Click()
    Unload Me
End Sub
'''

PROC_SHEET_CODE = '''\
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count = 1 Then
        If Target.Row >= 6 And Target.Row <= 105 Then
            If Target.Column = 11 Or Target.Column = 12 Then
                Application.EnableEvents = False
                If Target.Column = 11 Then
                    GRC_Macros.TonenAssetPicker Target.Row, "Information Assets", 11
                Else
                    GRC_Macros.TonenAssetPicker Target.Row, "Dependent Assets", 12
                End If
                Application.EnableEvents = True
            End If
        End If
    End If
End Sub
'''

VBA_CODE = r'''Attribute VB_Name = "GRC_Macros"
Option Explicit

Public g_TargetRow As Long         ' Onthoudt welke rij in Processen geklikt werd
Public g_ProcesNaam As String      ' Naam van het geselecteerde proces
Public g_PickerSourceSheet As String ' Bronblad voor de picker (Informatieassets / Afhankelijke assets)
Public g_PickerTargetCol As Long   ' Doelkolom in Processen (11 = info assets, 12 = afhankelijke assets)
Public g_RARMCol As Long           ' Kolom in RARM-sheet waarop geklikt werd (vuln picker)

' GRC Tool v0.3 - Import & Export via MS Access (ADO)

' Zet classificatietekst of -cijfer om naar code 1-5
Private Function MapCls(val As String) As Integer
    Dim s As String
    s = LCase(Trim(val))
    Dim n As Integer
    n = 0
    On Error Resume Next
    n = CInt(s)
    On Error GoTo 0
    If n >= 1 And n <= 5 Then MapCls = n: Exit Function
    If InStr(s, "kritiek") > 0 Or InStr(s, "critique") > 0 Or InStr(s, "critical") > 0 Then
        MapCls = 5
    ElseIf InStr(s, "zeer hoog") > 0 Or InStr(s, "very high") > 0 Or InStr(s, "tr") > 0 And InStr(s, "s " & Chr(233)) > 0 Then
        MapCls = 4
    ElseIf InStr(s, "hoog") > 0 Or InStr(s, "high") > 0 Or InStr(s, Chr(233) & "lev") > 0 Then
        MapCls = 3
    ElseIf InStr(s, "gemiddeld") > 0 Or InStr(s, "moyen") > 0 Or InStr(s, "medium") > 0 Then
        MapCls = 2
    ElseIf InStr(s, "laag") > 0 Or InStr(s, "faible") > 0 Or InStr(s, "low") > 0 Then
        MapCls = 1
    Else
        MapCls = 0
    End If
End Function

' Haal veldwaarde op — pass 1: exacte match, pass 2: gedeeltelijke match
Private Function FieldVal(rs As Object, ParamArray names() As Variant) As String
    Dim n As Variant, fld As Object
    For Each n In names
        For Each fld In rs.Fields
            If LCase(Trim(fld.Name)) = LCase(Trim(CStr(n))) Then
                If IsNull(fld.Value) Then FieldVal = "" Else FieldVal = CStr(fld.Value)
                Exit Function
            End If
        Next fld
    Next n
    For Each n In names
        For Each fld In rs.Fields
            If InStr(LCase(fld.Name), LCase(Trim(CStr(n)))) > 0 Then
                If IsNull(fld.Value) Then FieldVal = "" Else FieldVal = CStr(fld.Value)
                Exit Function
            End If
        Next fld
    Next n
    FieldVal = ""
End Function

' Diagnostiek: toont alle tabellen + queries met veldnamen en SQL
Sub VeldenTonen()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Selecteer de Access-database voor diagnostiek"
    fd.Filters.Clear
    fd.Filters.Add "Access-bestanden", "*.accdb; *.mdb"
    If fd.Show = False Then Exit Sub

    Dim conn As Object
    Set conn = OpenAccess(fd.SelectedItems(1))
    If conn Is Nothing Then Exit Sub

    Dim msg As String
    Dim schemars As Object, rs As Object, fld As Object

    ' ─ Tabellen ─
    msg = "=== TABELLEN ===" & vbCrLf & vbCrLf
    Set schemars = conn.OpenSchema(20)   ' adSchemaTables
    Do While Not schemars.EOF
        If schemars.Fields("TABLE_TYPE").Value = "TABLE" Then
            Dim tname As String
            tname = schemars.Fields("TABLE_NAME").Value
            msg = msg & "[" & tname & "]" & vbCrLf
            Set rs = CreateObject("ADODB.Recordset")
            On Error Resume Next
            rs.Open "SELECT TOP 1 * FROM [" & tname & "]", conn, 0, 1
            If Err.Number = 0 Then
                For Each fld In rs.Fields
                    msg = msg & "    " & fld.Name & vbCrLf
                Next fld
                rs.Close
            Else
                Err.Clear
            End If
            On Error GoTo 0
            msg = msg & vbCrLf
        End If
        schemars.MoveNext
    Loop
    schemars.Close

    ' ─ Queries ─
    msg = msg & "=== QUERIES ===" & vbCrLf & vbCrLf
    On Error Resume Next
    Set schemars = conn.OpenSchema(23)   ' adSchemaViews
    On Error GoTo 0
    If Not schemars Is Nothing Then
        Do While Not schemars.EOF
            Dim qname As String, qsql As String
            qname = schemars.Fields("TABLE_NAME").Value
            On Error Resume Next
            qsql = schemars.Fields("VIEW_DEFINITION").Value
            On Error GoTo 0
            msg = msg & "[" & qname & "]" & vbCrLf
            msg = msg & "    SQL: " & Left(qsql, 200) & vbCrLf & vbCrLf
            schemars.MoveNext
        Loop
        schemars.Close
    End If

    conn.Close

    ' Schrijf resultaat naar nieuw werkblad (te lang voor MsgBox)
    Dim owb As Workbook
    Set owb = Workbooks.Add
    Dim ows As Worksheet
    Set ows = owb.Sheets(1)
    ows.Name = "DB Diagnostiek"
    ows.Cells(1, 1) = msg
    ows.Columns(1).ColumnWidth = 120
    ows.Rows(1).WrapText = True
    ows.Rows(1).AutoFit
    MsgBox "Diagnostiek geschreven naar nieuw werkblad.", vbInformation, "GRC Diagnostiek"
End Sub

' Open Access-database via ADO (probeert ACE 12.0, dan Jet 4.0)
Private Function OpenAccess(dbPath As String) As Object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    On Error Resume Next
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    If Err.Number <> 0 Then
        Err.Clear
        conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";"
    End If
    On Error GoTo 0
    If conn.State = 0 Then
        MsgBox "Verbinding met Access-database mislukt." & vbCrLf & _
               "Zorg dat Microsoft Access Database Engine (ACE) is ge" & Chr(239) & "nstalleerd.", _
               vbExclamation, "GRC Import"
        Set OpenAccess = Nothing
    Else
        Set OpenAccess = conn
    End If
End Function

' Importeer Processen uit Access-tabel "T - Processes in scope"
Sub ImportProcessen()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Selecteer de Access-database (processen)"
    fd.Filters.Clear
    fd.Filters.Add "Access-bestanden", "*.accdb; *.mdb"
    If fd.Show = False Then Exit Sub

    Application.ScreenUpdating = False

    Dim conn As Object
    Set conn = OpenAccess(fd.SelectedItems(1))
    If conn Is Nothing Then Application.ScreenUpdating = True: Exit Sub

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    On Error Resume Next
    rs.Open "SELECT * FROM [T - Processes in scope]", conn, 0, 1
    If Err.Number <> 0 Then
        MsgBox "Tabel 'T - Processes in scope' niet gevonden in de database.", vbExclamation, "GRC Import"
        conn.Close: Application.ScreenUpdating = True: Exit Sub
    End If
    On Error GoTo 0

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Processes")
    ws.Range("B6:E105").ClearContents
    ws.Range("F6:F105").ClearContents  ' Integriteit code
    ws.Range("H6:H105").ClearContents  ' Beschikbaarheid code
    ws.Range("J6:J105").ClearContents  ' Opmerkingen
    ws.Range("K6:K105").ClearContents  ' Gekoppelde informatieassets
    ws.Range("L6:L105").ClearContents  ' Gekoppelde afhankelijke assets

    Dim destRow As Long
    destRow = 6
    Do While Not rs.EOF And destRow <= 105
        ws.Cells(destRow, 2) = FieldVal(rs, "naam", "procesnaam", "process naam", "name", "process name", "nom", "nom du processus", "processus", "titel", "title")
        ws.Cells(destRow, 3) = FieldVal(rs, "omschrijving", "beschrijving", "description", "beschr", "omschr", "toelichting", "detail", "details")
        ws.Cells(destRow, 4) = FieldVal(rs, "eigenaar", "proceseigenaar", "process eigenaar", "process owner", "verantwoordelijke", "owner", "propri" & Chr(233) & "taire", "responsible")
        ws.Cells(destRow, 5) = FieldVal(rs, "dienst", "afdeling", "organisatie", "entiteit", "department", "service", "business unit", "entity")
        ws.Cells(destRow, 6) = MapCls(FieldVal(rs, "integriteit", "integrity", "int" & Chr(233) & "grit" & Chr(233), "integr"))
        ws.Cells(destRow, 8) = MapCls(FieldVal(rs, "beschikbaarheid", "availability", "disponibilit" & Chr(233), "beschikb", "availab"))
        destRow = destRow + 1
        rs.MoveNext
    Loop

    rs.Close

    ' Haal gekoppelde assets op via de koppelingstabellen
    VulGekoppeldeAssets conn, ws, 6, 105
    VulGekoppeldeAfhankelijkeAssets conn, ws, 6, 105
    ws.Rows("6:105").AutoFit

    conn.Close
    Application.ScreenUpdating = True
    MsgBox destRow - 6 & " processen ge" & Chr(239) & "mporteerd.", vbInformation, "GRC Import"
End Sub

' Generieke functie: koppelt een asset-tabel aan processen via een koppelingstabel
Private Sub VulGekoppeldeLijst(conn As Object, ws As Worksheet, startRow As Long, endRow As Long, _
                                junctionTable As String, assetTable As String, _
                                assetKwd1 As String, assetKwd2 As String, targetCol As Long)
    Dim tmpRs As Object, f As Object, fn As String

    ' Stap 1: FK-velden in koppelingstabel
    Set tmpRs = CreateObject("ADODB.Recordset")
    On Error Resume Next
    tmpRs.Open "SELECT TOP 1 * FROM [" & junctionTable & "]", conn, 0, 1
    If Err.Number <> 0 Then
        MsgBox "Tabel '" & junctionTable & "' niet gevonden.", vbExclamation, "GRC Import"
        On Error GoTo 0: Exit Sub
    End If
    On Error GoTo 0
    Dim procFld As String, assetFld As String
    For Each f In tmpRs.Fields
        fn = LCase(f.Name)
        If procFld = "" And (InStr(fn, "process") > 0 Or InStr(fn, "proces") > 0) Then procFld = f.Name
        If assetFld = "" And InStr(fn, assetKwd1) > 0 Then assetFld = f.Name
        If assetFld = "" And assetKwd2 <> "" And InStr(fn, assetKwd2) > 0 Then assetFld = f.Name
    Next f
    tmpRs.Close
    If procFld = "" Or assetFld = "" Then
        MsgBox "FK-velden niet herkend in '" & junctionTable & "'." & vbCrLf & _
               "procFld='" & procFld & "'  assetFld='" & assetFld & "'", vbExclamation, "GRC Import"
        Exit Sub
    End If

    ' Stap 2: naamveld in T - Processes in scope
    Dim procNameFld As String
    Set tmpRs = CreateObject("ADODB.Recordset")
    On Error Resume Next
    tmpRs.Open "SELECT TOP 1 * FROM [T - Processes in scope]", conn, 0, 1
    On Error GoTo 0
    For Each f In tmpRs.Fields
        fn = LCase(f.Name)
        If procNameFld = "" And fn <> "id" Then
            If InStr(fn, "name") > 0 Or InStr(fn, "naam") > 0 Or InStr(fn, "nom") > 0 Then
                procNameFld = f.Name
            End If
        End If
    Next f
    tmpRs.Close
    If procNameFld = "" Then procNameFld = "ProcessName"

    ' Stap 3: naamveld in asset-tabel
    Dim assetNameFld As String
    Set tmpRs = CreateObject("ADODB.Recordset")
    On Error Resume Next
    tmpRs.Open "SELECT TOP 1 * FROM [" & assetTable & "]", conn, 0, 1
    On Error GoTo 0
    For Each f In tmpRs.Fields
        fn = LCase(f.Name)
        If assetNameFld = "" And fn <> "id" Then
            If InStr(fn, "name") > 0 Or InStr(fn, "naam") > 0 Or InStr(fn, "nom") > 0 Then
                assetNameFld = f.Name
            End If
        End If
    Next f
    tmpRs.Close
    If assetNameFld = "" Then assetNameFld = "Name"

    ' Stap 4: JOIN
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    Dim sql As String
    sql = "SELECT p.[" & procNameFld & "] AS PNaam, a.[" & assetNameFld & "] AS ANaam " & _
          "FROM ([T - Processes in scope] AS p " & _
          "INNER JOIN [" & junctionTable & "] AS lt ON lt.[" & procFld & "] = p.ID) " & _
          "INNER JOIN [" & assetTable & "] AS a ON a.ID = lt.[" & assetFld & "] " & _
          "ORDER BY p.[" & procNameFld & "], a.[" & assetNameFld & "]"
    On Error Resume Next
    rs.Open sql, conn, 0, 1
    If Err.Number <> 0 Then
        MsgBox "JOIN mislukt." & vbCrLf & sql, vbExclamation, "GRC Import"
        On Error GoTo 0: Exit Sub
    End If
    On Error GoTo 0

    ' Stap 5: bouw process→assets map
    Dim pkeys(500) As String, pvals(500) As String
    Dim pcount As Long: pcount = 0
    Dim pref As String, aref As String, fi As Long, pfound As Boolean
    Do While Not rs.EOF
        If IsNull(rs.Fields("PNaam").Value) Then pref = "" Else pref = CStr(rs.Fields("PNaam").Value)
        If IsNull(rs.Fields("ANaam").Value) Then aref = "" Else aref = CStr(rs.Fields("ANaam").Value)
        If pref <> "" And aref <> "" Then
            pfound = False
            For fi = 0 To pcount - 1
                If LCase(pkeys(fi)) = LCase(pref) Then
                    pvals(fi) = pvals(fi) & Chr(10) & aref: pfound = True: Exit For
                End If
            Next fi
            If Not pfound And pcount <= 500 Then
                pkeys(pcount) = pref: pvals(pcount) = aref: pcount = pcount + 1
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close

    ' Stap 6: schrijf naar doelkolom
    Dim r As Long, pn As String
    For r = startRow To endRow
        pn = CStr(ws.Cells(r, 2).Value)
        If pn <> "" Then
            For fi = 0 To pcount - 1
                If LCase(pkeys(fi)) = LCase(pn) Then
                    ws.Cells(r, targetCol) = pvals(fi)
                    ws.Rows(r).AutoFit
                    Exit For
                End If
            Next fi
        End If
    Next r
End Sub

' Vult kolom K (11) van Processen via koppelingstabel voor informateassets
Private Sub VulGekoppeldeAssets(conn As Object, ws As Worksheet, startRow As Long, endRow As Long)
    VulGekoppeldeLijst conn, ws, startRow, endRow, _
        "LT - Information Assets to Processes", "T - Information Assets", "asset", "informat", 11
End Sub

' Vult kolom L (12) van Processen via koppelingstabel voor afhankelijke assets
Private Sub VulGekoppeldeAfhankelijkeAssets(conn As Object, ws As Worksheet, startRow As Long, endRow As Long)
    VulGekoppeldeLijst conn, ws, startRow, endRow, _
        "LT - Dependent assets to Processes", "T - Dependent assets", "depend", "", 12
End Sub

' Importeer alles in één stap vanuit één Access-database
Sub ImportAlles()
    Dim fd As FileDialog
    Dim conn As Object
    Dim rs As Object
    Dim ws As Worksheet
    Dim destRow As Long
    Dim nProc As Long
    Dim nAsset As Long
    Dim nDep As Long
    Dim nCls As Integer

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Selecteer de Access-database"
    fd.Filters.Clear
    fd.Filters.Add "Access-bestanden", "*.accdb; *.mdb"
    If fd.Show = False Then Exit Sub

    Application.ScreenUpdating = False
    Set conn = OpenAccess(fd.SelectedItems(1))
    If conn Is Nothing Then Application.ScreenUpdating = True: Exit Sub

    ' --- 1. Informatieassets ---
    Set rs = CreateObject("ADODB.Recordset")
    On Error Resume Next
    rs.Open "SELECT * FROM [T - Information Assets]", conn, 0, 1
    If Err.Number <> 0 Then
        MsgBox "Tabel 'T - Information Assets' niet gevonden.", vbExclamation, "GRC Import"
        conn.Close: Application.ScreenUpdating = True: Exit Sub
    End If
    On Error GoTo 0
    Set ws = ThisWorkbook.Sheets("Information Assets")
    ws.Range("B6:E105").ClearContents
    ws.Range("F6:F105").ClearContents
    ws.Range("H6:H105").ClearContents
    destRow = 6
    Do While Not rs.EOF And destRow <= 105
        ws.Cells(destRow, 2) = FieldVal(rs, "naam", "assetnaam", "asset naam", "name", "asset name", "nom", "actif", "titel", "title")
        ws.Cells(destRow, 3) = FieldVal(rs, "omschrijving", "beschrijving", "description", "beschr", "omschr", "toelichting", "detail")
        ws.Cells(destRow, 4) = FieldVal(rs, "eigenaar", "asseteigenaar", "asset eigenaar", "owner", "propri" & Chr(233) & "taire", "verantwoordelijke")
        ws.Cells(destRow, 5) = FieldVal(rs, "dienst", "afdeling", "organisatie", "entiteit", "department", "service", "business unit")
        ws.Cells(destRow, 6) = MapCls(FieldVal(rs, "confidentialiteit", "confidentiality", "confidentialit" & Chr(233), "conf"))
        destRow = destRow + 1
        rs.MoveNext
    Loop
    rs.Close
    nAsset = destRow - 6

    ' --- 2. Afhankelijke assets ---
    Set rs = CreateObject("ADODB.Recordset")
    On Error Resume Next
    rs.Open "SELECT * FROM [T - Dependent assets]", conn, 0, 1
    If Err.Number <> 0 Then
        MsgBox "Tabel 'T - Dependent assets' niet gevonden.", vbExclamation, "GRC Import"
        conn.Close: Application.ScreenUpdating = True: Exit Sub
    End If
    On Error GoTo 0
    Set ws = ThisWorkbook.Sheets("Dependent Assets")
    ws.Unprotect
    ws.Range("B6:E105").ClearContents
    ws.Range("F6:F105").ClearContents   ' Overarching
    ws.Range("G6:G105").ClearContents   ' Opmerkingen
    ws.Range("P6:P105").ClearContents   ' Conf objective (col 16)
    ws.Range("R6:R105").ClearContents   ' Int objective (col 18)
    ws.Range("T6:T105").ClearContents   ' Avail objective (col 20)
    destRow = 6
    Do While Not rs.EOF And destRow <= 105
        ws.Cells(destRow, 2) = FieldVal(rs, "naam", "name", "nom", "afhankelijk", "dependent", "asset naam", "asset name")
        ws.Cells(destRow, 3) = FieldVal(rs, "omschrijving", "beschrijving", "description", "beschr", "omschr")
        ws.Cells(destRow, 4) = FieldVal(rs, "eigenaar", "owner", "propri" & Chr(233) & "taire", "verantwoordelijke")
        ws.Cells(destRow, 5) = FieldVal(rs, "dienst", "afdeling", "department", "service", "entiteit")
        Dim oarchFld2 As Object
        On Error Resume Next: Set oarchFld2 = rs.Fields("Overarching"): On Error GoTo 0
        If Not oarchFld2 Is Nothing Then
            If Not IsNull(oarchFld2.Value) Then
                If CBool(oarchFld2.Value) Then ws.Cells(destRow, 6) = ChrW(10004)
            End If
            Set oarchFld2 = Nothing
        End If
        nCls = MapCls(FieldVal(rs, "c-objective", "confidentiality objective", "conf objective", "c obj"))
        If nCls > 0 Then ws.Cells(destRow, 16) = nCls
        nCls = MapCls(FieldVal(rs, "i-objective", "integrity objective", "int objective", "i obj"))
        If nCls > 0 Then ws.Cells(destRow, 18) = nCls
        nCls = MapCls(FieldVal(rs, "a-objective", "availability objective", "avail objective", "a obj"))
        If nCls > 0 Then ws.Cells(destRow, 20) = nCls
        destRow = destRow + 1
        rs.MoveNext
    Loop
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    rs.Close
    nDep = destRow - 6

    ' --- 3. Processen ---
    Set rs = CreateObject("ADODB.Recordset")
    On Error Resume Next
    rs.Open "SELECT * FROM [T - Processes in scope]", conn, 0, 1
    If Err.Number <> 0 Then
        MsgBox "Tabel 'T - Processes in scope' niet gevonden.", vbExclamation, "GRC Import"
        conn.Close: Application.ScreenUpdating = True: Exit Sub
    End If
    On Error GoTo 0
    Set ws = ThisWorkbook.Sheets("Processes")
    ws.Range("B6:E105").ClearContents
    ws.Range("F6:F105").ClearContents
    ws.Range("H6:H105").ClearContents
    ws.Range("J6:J105").ClearContents
    ws.Range("K6:K105").ClearContents
    ws.Range("L6:L105").ClearContents
    destRow = 6
    Do While Not rs.EOF And destRow <= 105
        ws.Cells(destRow, 2) = FieldVal(rs, "naam", "procesnaam", "process naam", "name", "process name", "nom", "nom du processus", "processus", "titel", "title")
        ws.Cells(destRow, 3) = FieldVal(rs, "omschrijving", "beschrijving", "description", "beschr", "omschr", "toelichting", "detail", "details")
        ws.Cells(destRow, 4) = FieldVal(rs, "eigenaar", "proceseigenaar", "process eigenaar", "process owner", "verantwoordelijke", "owner", "propri" & Chr(233) & "taire", "responsible")
        ws.Cells(destRow, 5) = FieldVal(rs, "dienst", "afdeling", "organisatie", "entiteit", "department", "service", "business unit", "entity")
        ws.Cells(destRow, 6) = MapCls(FieldVal(rs, "integriteit", "integrity", "int" & Chr(233) & "grit" & Chr(233), "integr"))
        ws.Cells(destRow, 8) = MapCls(FieldVal(rs, "beschikbaarheid", "availability", "disponibilit" & Chr(233), "beschikb", "availab"))
        destRow = destRow + 1
        rs.MoveNext
    Loop
    rs.Close
    nProc = destRow - 6
    VulGekoppeldeAssets conn, ws, 6, 105
    VulGekoppeldeAfhankelijkeAssets conn, ws, 6, 105
    ws.Rows("6:105").AutoFit

    ' Herlaad RARM-kolomkoppen
    RefreshRARMKolommen

    ' Kwetsbaarheden en controls matrix mee importeren
    CoreImportKwetsbaarheden conn

    conn.Close
    KleurAlleRARMKolommen
    Application.ScreenUpdating = True
    MsgBox nProc & " processen, " & nAsset & " informatieassets en " & nDep & _
           " afhankelijke assets ge" & Chr(239) & "mporteerd.", vbInformation, "GRC Import"
End Sub

' Importeer alleen de koppelingen (standalone knop op Import & Export)
Sub ImportKoppelingen()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Selecteer de Access-database (koppelingen)"
    fd.Filters.Clear
    fd.Filters.Add "Access-bestanden", "*.accdb; *.mdb"
    If fd.Show = False Then Exit Sub

    Application.ScreenUpdating = False
    Dim conn As Object
    Set conn = OpenAccess(fd.SelectedItems(1))
    If conn Is Nothing Then Application.ScreenUpdating = True: Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Processes")
    ws.Range("K6:K105").ClearContents
    ws.Range("L6:L105").ClearContents
    VulGekoppeldeAssets conn, ws, 6, 105
    VulGekoppeldeAfhankelijkeAssets conn, ws, 6, 105

    conn.Close
    Application.ScreenUpdating = True
    MsgBox "Koppelingen ge" & Chr(239) & "mporteerd (kolom K: informateassets, kolom L: afhankelijke assets).", vbInformation, "GRC Import"
End Sub

' Importeer Afhankelijke assets uit Access-tabel "T - Dependent assets"
Sub ImportAfhankelijkeAssets()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Selecteer de Access-database (afhankelijke assets)"
    fd.Filters.Clear
    fd.Filters.Add "Access-bestanden", "*.accdb; *.mdb"
    If fd.Show = False Then Exit Sub

    Application.ScreenUpdating = False

    Dim conn As Object
    Set conn = OpenAccess(fd.SelectedItems(1))
    If conn Is Nothing Then Application.ScreenUpdating = True: Exit Sub

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    On Error Resume Next
    rs.Open "SELECT * FROM [T - Dependent assets]", conn, 0, 1
    If Err.Number <> 0 Then
        MsgBox "Tabel 'T - Dependent assets' niet gevonden in de database.", vbExclamation, "GRC Import"
        conn.Close: Application.ScreenUpdating = True: Exit Sub
    End If
    On Error GoTo 0

    Dim ws As Worksheet
    Dim nCls As Integer
    Set ws = ThisWorkbook.Sheets("Dependent Assets")
    ws.Unprotect
    ws.Range("B6:E105").ClearContents
    ws.Range("F6:F105").ClearContents   ' Overarching
    ws.Range("G6:G105").ClearContents   ' Opmerkingen
    ws.Range("P6:P105").ClearContents   ' Conf objective (col 16)
    ws.Range("R6:R105").ClearContents   ' Int objective (col 18)
    ws.Range("T6:T105").ClearContents   ' Avail objective (col 20)

    Dim destRow As Long
    destRow = 6
    Do While Not rs.EOF And destRow <= 105
        ws.Cells(destRow, 2) = FieldVal(rs, "naam", "name", "nom", "afhankelijk", "dependent", "asset naam", "asset name")
        ws.Cells(destRow, 3) = FieldVal(rs, "omschrijving", "beschrijving", "description", "beschr", "omschr")
        ws.Cells(destRow, 4) = FieldVal(rs, "eigenaar", "owner", "propri" & Chr(233) & "taire", "verantwoordelijke")
        ws.Cells(destRow, 5) = FieldVal(rs, "dienst", "afdeling", "department", "service", "entiteit")
        Dim oarchFld As Object
        On Error Resume Next: Set oarchFld = rs.Fields("Overarching"): On Error GoTo 0
        If Not oarchFld Is Nothing Then
            If Not IsNull(oarchFld.Value) Then
                If CBool(oarchFld.Value) Then ws.Cells(destRow, 6) = ChrW(10004)
            End If
            Set oarchFld = Nothing
        End If
        nCls = MapCls(FieldVal(rs, "c-objective", "confidentiality objective", "conf objective", "c obj"))
        If nCls > 0 Then ws.Cells(destRow, 16) = nCls
        nCls = MapCls(FieldVal(rs, "i-objective", "integrity objective", "int objective", "i obj"))
        If nCls > 0 Then ws.Cells(destRow, 18) = nCls
        nCls = MapCls(FieldVal(rs, "a-objective", "availability objective", "avail objective", "a obj"))
        If nCls > 0 Then ws.Cells(destRow, 20) = nCls
        destRow = destRow + 1
        rs.MoveNext
    Loop

    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True

    rs.Close

    ' Koppel ook meteen aan Processen-sheet
    Dim procWs As Worksheet
    Set procWs = ThisWorkbook.Sheets("Processes")
    procWs.Range("L6:L105").ClearContents
    VulGekoppeldeAfhankelijkeAssets conn, procWs, 6, 105

    ' Herlaad RARM-kolomkoppen
    RefreshRARMKolommen

    conn.Close
    Application.ScreenUpdating = True
    KleurAlleRARMKolommen

    MsgBox destRow - 6 & " afhankelijke assets ge" & Chr(239) & "mporteerd.", vbInformation, "GRC Import"
End Sub

' Importeer koppelingen kwetsbaarheden <-> afhankelijke assets (standalone knop)
' Voer dit uit NA de import van kwetsbaarheden.
Sub ImportLinksKwetsbaarheden()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Selecteer de Access-database"
    fd.Filters.Clear
    fd.Filters.Add "Access-bestanden", "*.accdb; *.mdb"
    If fd.Show = False Then Exit Sub
    Application.ScreenUpdating = False
    Dim conn As Object
    Set conn = OpenAccess(fd.SelectedItems(1))
    If conn Is Nothing Then Application.ScreenUpdating = True: Exit Sub
    RefreshRARMKolommen
    ImportRARMKwetsbaarheden conn
    ImportGeselecteerdeControls conn
    conn.Close
    Application.ScreenUpdating = True
    KleurAlleRARMKolommen
    MsgBox "Links kwetsbaarheden / controls per afhankelijke asset ge" & Chr(239) & "mporteerd.", vbInformation, "GRC Import"
End Sub

' Importeert geselecteerde controls per DA vanuit LT - Selected controls to DA
' en schrijft ChrW(10004) in de overeenkomstige RARM-datacellen.
' Vertaalketen: ControlReference -> T-CyFunEssentiel.RefNr -> 2023-ID
'              -> CyFun Controls sheet -> 2025-ID -> rijnummer in RARM
Sub ImportGeselecteerdeControls(conn As Object)
    Const RARM_ROW_DA   As Long = 2
    Const RARM_ROW_DATA As Long = 4
    Const RARM_COL_ID   As Long = 1
    Const RARM_COL_DA   As Long = 5
    Const COL_CF25      As Integer = 6
    Const COL_CF23      As Integer = 13
    Const ROW_CF_DATA   As Integer = 4

    Dim wsR As Worksheet, wsCC As Worksheet
    On Error Resume Next
    Set wsR = ThisWorkbook.Sheets("RARM")
    Set wsCC = ThisWorkbook.Sheets("CyFun Controls")
    On Error GoTo 0
    If wsR Is Nothing Then Exit Sub

    Dim rarmLast As Long
    rarmLast = wsR.Cells(wsR.Rows.Count, RARM_COL_ID).End(xlUp).Row
    Dim lastCol As Long
    lastCol = wsR.Cells(RARM_ROW_DA, wsR.Columns.Count).End(xlToLeft).Column

    ' ── 1. ctrlRowMap: LCase(2025-ID) → rijnummer in RARM ───────────────────────
    Dim ctrlRowMap As Object
    Set ctrlRowMap = CreateObject("Scripting.Dictionary")
    Dim rr As Long
    For rr = RARM_ROW_DATA To rarmLast
        Dim ck As String
        ck = LCase(Trim(CStr(wsR.Cells(rr, RARM_COL_ID).Value)))
        If ck <> "" And Not ctrlRowMap.Exists(ck) Then ctrlRowMap.Add ck, rr
    Next rr

    ' ── 2. refNrMap: RefNr (Long) → CyFun 2023 ID-string ───────────────────────
    Dim refNrMap As Object
    Set refNrMap = CreateObject("Scripting.Dictionary")
    Dim rsCtrl As Object
    Set rsCtrl = CreateObject("ADODB.Recordset")
    rsCtrl.Open "SELECT RefNr, Requirement FROM [T - CyFunEssentiel]", conn
    Do While Not rsCtrl.EOF
        Dim refNr As Long
        refNr = CLng(rsCtrl.Fields("RefNr").Value)
        If Not refNrMap.Exists(refNr) Then
            refNrMap.Add refNr, CStr(rsCtrl.Fields("Requirement").Value)
        End If
        rsCtrl.MoveNext
    Loop
    rsCtrl.Close

    ' ── 3. rev23to25: LCase(2023-ID) → LCase(2025-ID) via CyFun Controls sheet ──
    Dim rev23to25 As Object
    Set rev23to25 = CreateObject("Scripting.Dictionary")
    If Not wsCC Is Nothing Then
        Dim ccLast As Long
        ccLast = Application.Max( _
            wsCC.Cells(wsCC.Rows.Count, COL_CF25).End(xlUp).Row, _
            wsCC.Cells(wsCC.Rows.Count, COL_CF23).End(xlUp).Row)
        Dim rCC As Long
        For rCC = ROW_CF_DATA To ccLast
            Dim r25t As String, r23t As String
            r25t = Trim(CStr(wsCC.Cells(rCC, COL_CF25).Value))
            r23t = Trim(CStr(wsCC.Cells(rCC, COL_CF23).Value))
            If r25t <> "" And r23t <> "" Then
                Dim p25 As Integer, p23 As Integer
                p25 = InStr(r25t, ":"): p23 = InStr(r23t, ":")
                Dim i25 As String, i23 As String
                i25 = LCase(Trim(Replace(IIf(p25 > 0, Left(r25t, p25 - 1), r25t), Chr(9), "")))
                i23 = NormId23(r23t)
                If i23 <> "" And i25 <> "" And Not rev23to25.Exists(i23) Then
                    rev23to25.Add i23, i25
                End If
            End If
        Next rCC
    End If

    ' ── 4. daIdColMap: DAID (Long) → RARM-kolom ─────────────────────────────────
    ' Koppel via naam: T - Dependent assets.DAName ↔ RARM rij 2 (LCase)
    Dim daNameIdMap As Object
    Set daNameIdMap = CreateObject("Scripting.Dictionary")
    Dim rsDA As Object
    Set rsDA = CreateObject("ADODB.Recordset")
    rsDA.Open "SELECT ID, DAName FROM [T - Dependent assets]", conn
    Do While Not rsDA.EOF
        Dim daId As Long
        daId = CLng(rsDA.Fields("ID").Value)
        Dim daNm As String
        daNm = LCase(Trim(CStr(rsDA.Fields("DAName").Value)))
        If daNm <> "" And Not daNameIdMap.Exists(daNm) Then
            daNameIdMap.Add daNm, daId
        End If
        rsDA.MoveNext
    Loop
    rsDA.Close

    Dim daIdColMap As Object
    Set daIdColMap = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = RARM_COL_DA To lastCol
        Dim rarmDaNm As String
        rarmDaNm = LCase(Trim(CStr(wsR.Cells(RARM_ROW_DA, c).Value)))
        If rarmDaNm <> "" And daNameIdMap.Exists(rarmDaNm) Then
            Dim daIdForCol As Long
            daIdForCol = CLng(daNameIdMap(rarmDaNm))
            If Not daIdColMap.Exists(daIdForCol) Then daIdColMap.Add daIdForCol, c
        End If
    Next c

    ' ── 5. Wis bestaande vinkjes in DA-kolommen ──────────────────────────────────
    If lastCol >= RARM_COL_DA And rarmLast >= RARM_ROW_DATA Then
        wsR.Range(wsR.Cells(RARM_ROW_DATA, RARM_COL_DA), _
                  wsR.Cells(rarmLast, lastCol)).ClearContents
    End If

    ' ── 6. Lees LT - Selected controls to DA en schrijf vinkjes in RARM ─────────
    Dim nMatched As Long
    nMatched = 0
    Dim rsLT As Object
    Set rsLT = CreateObject("ADODB.Recordset")
    rsLT.Open "SELECT DAID, ControlReference FROM [LT - Selected controls to DA]", conn
    Do While Not rsLT.EOF
        Dim ltDaId As Long, ctrlRef As Long
        ltDaId  = CLng(rsLT.Fields("DAID").Value)
        ctrlRef = CLng(rsLT.Fields("ControlReference").Value)
        If daIdColMap.Exists(ltDaId) And refNrMap.Exists(ctrlRef) Then
            Dim id23v As String
            id23v = NormId23(CStr(refNrMap.Item(ctrlRef)))
            Dim id25v As String
            If rev23to25.Exists(id23v) Then
                id25v = CStr(rev23to25.Item(id23v))
            Else
                id25v = id23v
            End If
            If ctrlRowMap.Exists(id25v) Then
                Dim ctrlRow As Long
                ctrlRow = CLng(ctrlRowMap.Item(id25v))
                Dim targetCol As Long
                targetCol = CLng(daIdColMap(ltDaId))
                With wsR.Cells(ctrlRow, targetCol)
                    .Value = ChrW(10004)
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
                nMatched = nMatched + 1
            End If
        End If
        rsLT.MoveNext
    Loop
    rsLT.Close
End Sub

' Opent het UserForm VulnPicker voor kwetsbaarheid-selectie in RARM-sheet
Sub TonenVulnPicker(targetCol As Long)
    g_RARMCol = targetCol
    VulnPicker.Show
End Sub

' Verzamelt control IDs die een ✔ hebben voor minstens één van de opgegeven kwetsbaarheden.
' selVulns: kommagescheiden string met kwetsbaarheidsnamen (uit RARM rij 3)
' Retourneert een Scripting.Dictionary {lowercase ctrl_id → True}
Private Function GetMarkedCtrlIDs(wsK As Worksheet, selVulns As String) As Object
    Const KWETS_ROW_HDR    As Long = 2
    Const KWETS_ROW_DATA   As Long = 6
    Const KWETS_COL_ID     As Long = 1
    Const KWETS_COL_VSTART As Long = 3
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    If wsK Is Nothing Or selVulns = "" Then Set GetMarkedCtrlIDs = result: Exit Function
    Dim kwetsLast As Long
    kwetsLast = wsK.Cells(wsK.Rows.Count, KWETS_COL_ID).End(xlUp).Row
    Dim selParts() As String
    selParts = Split(selVulns, ", ")
    Dim p As Integer, kc As Long, r As Long
    Dim vName As String, vCol As Long, hdrVal As String, cv As String, cid As String
    For p = 0 To UBound(selParts)
        vName = Trim(selParts(p))
        ' Strip probability suffix (format: "VulnName (N)")
        Dim parenIdx As Integer
        parenIdx = InStrRev(vName, " (")
        If parenIdx > 0 Then vName = Left(vName, parenIdx - 1)
        If vName <> "" Then
            vCol = 0
            For kc = KWETS_COL_VSTART To KWETS_COL_VSTART + 300
                hdrVal = CStr(wsK.Cells(KWETS_ROW_HDR, kc).Value)
                If hdrVal = "" Then Exit For
                If hdrVal = vName Then vCol = kc: Exit For
            Next kc
            If vCol > 0 Then
                For r = KWETS_ROW_DATA To kwetsLast
                    cv = ""
                    If Not IsError(wsK.Cells(r, vCol).Value) Then cv = CStr(wsK.Cells(r, vCol).Value)
                    If cv = ChrW(10004) Then
                        cid = LCase(Trim(CStr(wsK.Cells(r, KWETS_COL_ID).Value)))
                        If cid <> "" And Not result.Exists(cid) Then result.Add cid, True
                    End If
                Next r
            End If
        End If
    Next p
    Set GetMarkedCtrlIDs = result
End Function

' Importeert kwetsbaarheden + probabiliteit per DA vanuit Access naar RARM rij 3.
' Leest SQL van "QLT2 - Vulnerabilities for DA - assigned" via ADOX, verwijdert de WHERE
' clausule (bevat formulierparameters) en voert de gecleande SQL rechtstreeks uit.
Sub ImportRARMKwetsbaarheden(conn As Object)
    Const RARM_ROW_DA   As Long = 2
    Const RARM_ROW_VULN As Long = 3
    Const RARM_COL_DA   As Long = 5

    Dim wsR As Worksheet
    On Error Resume Next
    Set wsR = ThisWorkbook.Sheets("RARM")
    On Error GoTo 0
    If wsR Is Nothing Then Exit Sub

    ' ── 1. Directe SQL op de onderliggende tabellen (geen Access-query nodig) ─────
    ' Tabel LT - Vulnerabilities to Dependent Assets: DAID, VulnerabilityID, Probability
    Dim qSQL As String
    qSQL = "SELECT [T - Dependent assets].DAName, [T - Vulnerabilities].Vulnerability, " & _
           "[LT - Vulnerabilities to Dependent Assets].Probability " & _
           "FROM ([T - Dependent assets] " & _
           "INNER JOIN [LT - Vulnerabilities to Dependent Assets] " & _
           "ON [T - Dependent assets].ID = [LT - Vulnerabilities to Dependent Assets].DAID) " & _
           "INNER JOIN [T - Vulnerabilities] " & _
           "ON [T - Vulnerabilities].ID = [LT - Vulnerabilities to Dependent Assets].VulnerabilityID"

    ' ── 2. Voer SQL uit en groepeer per DA-naam ──────────────────────────────────
    ' qaMap: LCase(DA-naam) → vbTab-gescheiden "vulnNaam~prob" items
    Dim qaMap As Object
    Set qaMap = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    On Error Resume Next
    rs.Open qSQL, conn, 0, 1
    If Err.Number <> 0 Then
        MsgBox "Fout bij uitvoeren query: " & Err.Description, vbExclamation, "GRC Import"
        Exit Sub
    End If
    On Error GoTo 0

    Do While Not rs.EOF
        Dim daNm As String
        daNm = FieldVal(rs, "naam", "name", "da naam", "da name", "dependent asset", _
                        "asset naam", "asset name", "asset", "da")
        Dim vNm As String
        vNm = FieldVal(rs, "vulnerability", "kwetsbaarheid", "vuln naam", "vuln name", "vuln")
        Dim pStr As String
        pStr = "2"
        On Error Resume Next
        If Not IsNull(rs.Fields("Probability").Value) Then pStr = Trim(CStr(rs.Fields("Probability").Value))
        On Error GoTo 0
        If daNm <> "" And vNm <> "" Then
            Dim qKey As String
            qKey = LCase(Trim(daNm))
            Dim qEntry As String
            qEntry = vNm & "~" & pStr
            If qaMap.Exists(qKey) Then
                qaMap(qKey) = qaMap(qKey) & vbTab & qEntry
            Else
                qaMap(qKey) = qEntry
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close

    ' ── 3. Schrijf naar RARM rij 3 op basis van DA-naam in rij 2 (bovenste lijn) ─
    Dim col As Long
    col = RARM_COL_DA
    Do While Trim(CStr(wsR.Cells(RARM_ROW_DA, col).Value)) <> ""
        Dim rarmKey As String
        rarmKey = LCase(Trim(CStr(wsR.Cells(RARM_ROW_DA, col).Value)))
        Dim result As String
        result = ""
        If qaMap.Exists(rarmKey) Then
            Dim pairs() As String
            pairs = Split(CStr(qaMap(rarmKey)), vbTab)
            Dim pi As Integer
            For pi = 0 To UBound(pairs)
                Dim pp() As String
                pp = Split(pairs(pi), "~")
                If UBound(pp) >= 0 And pp(0) <> "" Then
                    If result <> "" Then result = result & ", "
                    result = result & pp(0)
                    If UBound(pp) >= 1 And pp(1) <> "" Then result = result & " (" & pp(1) & ")"
                End If
            Next pi
        End If
        wsR.Cells(RARM_ROW_VULN, col).Value = result
        col = col + 1
    Loop
End Sub

' Herlaadt DA-kolomkoppen in RARM vanuit de Dependent Assets sheet.
' Leest col 2 (naam) en col 6 (overarching) van rijen 6-105.
Sub RefreshRARMKolommen()
    Const RARM_ROW_DA   As Long = 2
    Const RARM_ROW_VULN As Long = 3
    Const RARM_ROW_DATA As Long = 4
    Const RARM_COL_ID   As Long = 1
    Const RARM_COL_DA   As Long = 5
    Const DA_COL_NAAM   As Long = 2
    Const DA_COL_OARCH  As Long = 6
    Dim wsR As Worksheet, wsD As Worksheet
    On Error Resume Next
    Set wsR = ThisWorkbook.Sheets("RARM")
    Set wsD = ThisWorkbook.Sheets("Dependent Assets")
    On Error GoTo 0
    If wsR Is Nothing Or wsD Is Nothing Then Exit Sub

    ' Bepaal het huidige bereik van DA-kolommen en de laatste datarij
    Dim lastCol As Long, rarmLastData As Long
    lastCol = wsR.Cells(RARM_ROW_DA, wsR.Columns.Count).End(xlToLeft).Column
    rarmLastData = wsR.Cells(wsR.Rows.Count, RARM_COL_ID).End(xlUp).Row

    ' Reset het volledige DA-zone: koppen (rij 2-3) + datacellen — zodat verwijderde
    ' kolommen geen opmaakrestanten achterlaten
    If lastCol >= RARM_COL_DA Then
        Dim hdrRng As Range, dataRng As Range
        Set hdrRng  = wsR.Range(wsR.Cells(RARM_ROW_DA, RARM_COL_DA), _
                                wsR.Cells(RARM_ROW_VULN, lastCol + 5))
        Set dataRng = wsR.Range(wsR.Cells(RARM_ROW_DATA, RARM_COL_DA), _
                                wsR.Cells(rarmLastData, lastCol + 5))
        hdrRng.ClearContents
        hdrRng.Interior.ColorIndex  = xlNone
        hdrRng.Borders.LineStyle    = xlNone
        dataRng.Interior.ColorIndex = xlNone
        dataRng.Borders.LineStyle   = xlNone
    End If

    Dim col As Long
    col = RARM_COL_DA
    Dim r As Long, daNaam As String, isOarch As Boolean
    For r = 6 To 105
        daNaam = Trim(CStr(wsD.Cells(r, DA_COL_NAAM).Value))
        If daNaam <> "" Then
            isOarch = (Trim(CStr(wsD.Cells(r, DA_COL_OARCH).Value)) <> "")
            With wsR.Cells(RARM_ROW_DA, col)
                .Value = daNaam
                .Font.Bold = True: .Font.Size = 10
                .Font.Color = RGB(15, 43, 70)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
                If isOarch Then
                    .Interior.Color = RGB(180, 83, 9)
                Else
                    .Interior.Color = RGB(254, 249, 195)
                End If
                .Borders.LineStyle = xlContinuous
                .Borders.Color     = RGB(15, 43, 70)
            End With
            With wsR.Cells(RARM_ROW_VULN, col)
                .Font.Italic = True: .Font.Size = 9
                .Font.Color = RGB(100, 116, 139)
                .Interior.Color = RGB(239, 246, 255)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
                .Borders.LineStyle = xlContinuous
                .Borders.Color     = RGB(15, 43, 70)
            End With
            wsR.Columns(col).ColumnWidth = 22
            col = col + 1
        End If
    Next r
    ' Herlaad titel-merge (rij 1) zodat die altijd de volledige breedte dekt
    Dim titleLast As Long
    titleLast = col - 1
    If titleLast < RARM_COL_DA + 9 Then titleLast = RARM_COL_DA + 9  ' min 10 DA-kolommen
    On Error Resume Next
    wsR.Cells(1, 1).MergeArea.UnMerge
    On Error GoTo 0
    wsR.Range(wsR.Cells(1, 1), wsR.Cells(1, titleLast)).Merge
End Sub

' Kleurt alle DA-kolommen in RARM door KleurRARMKolom aan te roepen per kolom.
Sub KleurAlleRARMKolommen()
    Const RARM_ROW_DA As Long = 2
    Const RARM_COL_DA As Long = 5
    Dim wsR As Worksheet
    On Error Resume Next
    Set wsR = ThisWorkbook.Sheets("RARM")
    On Error GoTo 0
    If wsR Is Nothing Then Exit Sub
    Dim col As Long
    col = RARM_COL_DA
    Do While Trim(CStr(wsR.Cells(RARM_ROW_DA, col).Value)) <> ""
        KleurRARMKolom col
        col = col + 1
    Loop
End Sub

' Synchroniseert RARM DA-kolommen met de Dependent Assets-sheet.
' Bewaart rij-3-waarden (kwetsbaarheden) en vinkjes voor DA's die nog bestaan.
' Kolommen van weggevallen DA's worden verwijderd; nieuwe DA's krijgen een lege kolom.
' Roep KleurAlleRARMKolommen aan na afloop voor correcte kleuring.
Sub SyncRARMKolommen()
    Const RARM_ROW_DA   As Long = 2
    Const RARM_ROW_VULN As Long = 3
    Const RARM_ROW_DATA As Long = 4
    Const RARM_COL_ID   As Long = 1
    Const RARM_COL_DA   As Long = 5
    Const DA_COL_NAAM   As Long = 2
    Const DA_COL_OARCH  As Long = 6

    Dim wsR As Worksheet, wsD As Worksheet
    On Error Resume Next
    Set wsR = ThisWorkbook.Sheets("RARM")
    Set wsD = ThisWorkbook.Sheets("Dependent Assets")
    On Error GoTo 0
    If wsR Is Nothing Or wsD Is Nothing Then Exit Sub

    Dim rarmLast As Long
    rarmLast = wsR.Cells(wsR.Rows.Count, RARM_COL_ID).End(xlUp).Row
    Dim lastCol As Long
    lastCol = wsR.Cells(RARM_ROW_DA, wsR.Columns.Count).End(xlToLeft).Column

    ' ── 1. Snapshot: bewaar rij-3-waarde en vinkjes per DA-naam ─────────────────
    Dim oldVuln As Object
    Set oldVuln = CreateObject("Scripting.Dictionary")
    Dim oldData As Object
    Set oldData = CreateObject("Scripting.Dictionary")

    Dim c As Long, dKey As String, dNm As String
    For c = RARM_COL_DA To lastCol
        dNm = Trim(CStr(wsR.Cells(RARM_ROW_DA, c).Value))
        If dNm <> "" Then
            dKey = LCase(dNm)
            If Not oldVuln.Exists(dKey) Then
                oldVuln.Add dKey, CStr(wsR.Cells(RARM_ROW_VULN, c).Value)
                Dim rowDict As Object
                Set rowDict = CreateObject("Scripting.Dictionary")
                Dim r As Long
                For r = RARM_ROW_DATA To rarmLast
                    Dim cv As String
                    cv = CStr(wsR.Cells(r, c).Value)
                    If cv <> "" Then rowDict.Add r, cv
                Next r
                oldData.Add dKey, rowDict
            End If
        End If
    Next c

    ' ── 2. Wis DA-zone (koppen + data, incl. formattering) ──────────────────────
    If lastCol >= RARM_COL_DA Then
        Dim clearTo As Long
        clearTo = lastCol + 5
        With wsR.Range(wsR.Cells(RARM_ROW_DA, RARM_COL_DA), wsR.Cells(RARM_ROW_VULN, clearTo))
            .ClearContents
            .Interior.ColorIndex = xlNone
            .Borders.LineStyle   = xlNone
        End With
        With wsR.Range(wsR.Cells(RARM_ROW_DATA, RARM_COL_DA), wsR.Cells(rarmLast, clearTo))
            .ClearContents
            .Interior.ColorIndex = xlNone
            .Borders.LineStyle   = xlNone
        End With
    End If

    ' ── 3. Herbouw DA-kolommen in DA-sheet volgorde; herstel opgeslagen data ─────
    Dim col As Long
    col = RARM_COL_DA
    Dim daR As Long, daNameRaw As String, isOarch As Boolean, dkLow As String
    For daR = 6 To 105
        daNameRaw = Trim(CStr(wsD.Cells(daR, DA_COL_NAAM).Value))
        If daNameRaw <> "" Then
            isOarch = (Trim(CStr(wsD.Cells(daR, DA_COL_OARCH).Value)) <> "")
            dkLow = LCase(daNameRaw)

            With wsR.Cells(RARM_ROW_DA, col)
                .Value = daNameRaw
                .Font.Bold = True: .Font.Size = 10
                .Font.Color = RGB(15, 43, 70)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
                .Interior.Color = IIf(isOarch, RGB(180, 83, 9), RGB(254, 249, 195))
                .Borders.LineStyle = xlContinuous
                .Borders.Color     = RGB(15, 43, 70)
            End With

            With wsR.Cells(RARM_ROW_VULN, col)
                If oldVuln.Exists(dkLow) Then .Value = oldVuln(dkLow)
                .Font.Italic = True: .Font.Size = 9
                .Font.Color = RGB(100, 116, 139)
                .Interior.Color = RGB(239, 246, 255)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
                .Borders.LineStyle = xlContinuous
                .Borders.Color     = RGB(15, 43, 70)
            End With

            If oldData.Exists(dkLow) Then
                Dim rd As Object
                Set rd = oldData(dkLow)
                Dim rk As Variant
                For Each rk In rd.Keys
                    With wsR.Cells(CLng(rk), col)
                        .Value = rd(rk)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With
                Next rk
            End If

            wsR.Columns(col).ColumnWidth = 22
            col = col + 1
        End If
    Next daR

    ' ── 4. Titel-merge ───────────────────────────────────────────────────────────
    Dim titleLast As Long
    titleLast = col - 1
    If titleLast < RARM_COL_DA + 9 Then titleLast = RARM_COL_DA + 9
    On Error Resume Next
    wsR.Cells(1, 1).MergeArea.UnMerge
    On Error GoTo 0
    wsR.Range(wsR.Cells(1, 1), wsR.Cells(1, titleLast)).Merge
End Sub

' Kleurt één DA-kolom in RARM:
'   - Volledige oranje (RGB 255,192,0) voor controls gelinkt aan eigen geselecteerde kwetsbaarheden.
'   - Lichtgele spill (RGB 255,230,153) voor controls van overarching DA-kolommen,
'     maar alleen als de targetCol zelf NIET overarching is.
Sub KleurRARMKolom(targetCol As Long)
    Const RARM_ROW_DA      As Long = 2
    Const RARM_ROW_VULN    As Long = 3
    Const RARM_ROW_DATA    As Long = 4
    Const RARM_COL_ID      As Long = 1
    Const RARM_COL_DA      As Long = 5
    Const DA_COL_NAAM      As Long = 2
    Const DA_COL_OARCH     As Long = 6

    Dim wsR As Worksheet, wsK As Worksheet, wsD As Worksheet
    On Error Resume Next
    Set wsR = ThisWorkbook.Sheets("RARM")
    Set wsK = ThisWorkbook.Sheets("Kwetsbaarheden")
    Set wsD = ThisWorkbook.Sheets("Dependent Assets")
    On Error GoTo 0
    If wsR Is Nothing Or wsK Is Nothing Then Exit Sub

    Dim rarmLast As Long
    rarmLast = wsR.Cells(wsR.Rows.Count, RARM_COL_ID).End(xlUp).Row

    ' Reset naar wisselend grijs/wit met uniforme rand
    Dim bClr As Long
    bClr = RGB(203, 213, 225)   ' grey_border
    Dim r As Long
    For r = RARM_ROW_DATA To rarmLast
        With wsR.Cells(r, targetCol)
            .Interior.Color = IIf((r - RARM_ROW_DATA) Mod 2 = 0, RGB(242, 242, 242), RGB(255, 255, 255))
            .Borders.LineStyle = xlContinuous
            .Borders.Color     = bClr
        End With
    Next r

    ' Eigen kwetsbaarheden → volledige kleur
    Dim selVal As String
    selVal = Trim(CStr(wsR.Cells(RARM_ROW_VULN, targetCol).Value))
    Dim ownMarked As Object
    Set ownMarked = GetMarkedCtrlIDs(wsK, selVal)

    Dim rid As String
    For r = RARM_ROW_DATA To rarmLast
        rid = LCase(Trim(CStr(wsR.Cells(r, RARM_COL_ID).Value)))
        If ownMarked.Exists(rid) Then
            wsR.Cells(r, targetCol).Interior.Color = RGB(255, 192, 0)
        End If
    Next r

    ' Bepaal of targetCol zelf overarching is
    Dim targetDaName As String
    targetDaName = LCase(Trim(CStr(wsR.Cells(RARM_ROW_DA, targetCol).Value)))
    Dim isOarch As Boolean
    isOarch = False
    If Not wsD Is Nothing And targetDaName <> "" Then
        Dim daR As Long
        For daR = 6 To 105
            If LCase(Trim(CStr(wsD.Cells(daR, DA_COL_NAAM).Value))) = targetDaName Then
                isOarch = (Trim(CStr(wsD.Cells(daR, DA_COL_OARCH).Value)) <> "")
                Exit For
            End If
        Next daR
    End If

    ' Als targetCol niet overarching is → ook spill-kleur van overarching kolommen
    If Not isOarch And Not wsD Is Nothing Then
        Dim oarchMarked As Object
        Set oarchMarked = CreateObject("Scripting.Dictionary")
        Dim daCol As Long, oaDaName As String, oaVulns As String
        Dim oaIDs As Object, key As Variant
        daCol = RARM_COL_DA
        Do While Trim(CStr(wsR.Cells(RARM_ROW_DA, daCol).Value)) <> ""
            If daCol <> targetCol Then
                oaDaName = LCase(Trim(CStr(wsR.Cells(RARM_ROW_DA, daCol).Value)))
                For daR = 6 To 105
                    If LCase(Trim(CStr(wsD.Cells(daR, DA_COL_NAAM).Value))) = oaDaName Then
                        If Trim(CStr(wsD.Cells(daR, DA_COL_OARCH).Value)) <> "" Then
                            oaVulns = Trim(CStr(wsR.Cells(RARM_ROW_VULN, daCol).Value))
                            If oaVulns <> "" Then
                                Set oaIDs = GetMarkedCtrlIDs(wsK, oaVulns)
                                For Each key In oaIDs.Keys
                                    If Not oarchMarked.Exists(key) Then oarchMarked.Add key, True
                                Next key
                            End If
                        End If
                        Exit For
                    End If
                Next daR
            End If
            daCol = daCol + 1
        Loop

        For r = RARM_ROW_DATA To rarmLast
            rid = LCase(Trim(CStr(wsR.Cells(r, RARM_COL_ID).Value)))
            If oarchMarked.Exists(rid) Then
                If wsR.Cells(r, targetCol).Interior.Color <> RGB(255, 192, 0) Then
                    wsR.Cells(r, targetCol).Interior.Color = RGB(255, 230, 153)
                End If
            End If
        Next r
    End If
End Sub

' Opent het UserForm AssetPicker voor handmatige selectie van assets bij een proces
Sub TonenAssetPicker(targetRow As Long, sourceSheet As String, targetCol As Long)
    g_TargetRow = targetRow
    g_ProcesNaam = CStr(ThisWorkbook.Sheets("Processes").Cells(targetRow, 2).Value)
    g_PickerSourceSheet = sourceSheet
    g_PickerTargetCol = targetCol
    If g_ProcesNaam = "" Then Exit Sub
    AssetPicker.Show
End Sub

' Importeer Informatieassets uit Access-tabel "T - Information Assets"
Sub ImportInformatieassets()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Selecteer de Access-database (informatieassets)"
    fd.Filters.Clear
    fd.Filters.Add "Access-bestanden", "*.accdb; *.mdb"
    If fd.Show = False Then Exit Sub

    Application.ScreenUpdating = False

    Dim conn As Object
    Set conn = OpenAccess(fd.SelectedItems(1))
    If conn Is Nothing Then Application.ScreenUpdating = True: Exit Sub

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    On Error Resume Next
    rs.Open "SELECT * FROM [T - Information Assets]", conn, 0, 1
    If Err.Number <> 0 Then
        MsgBox "Tabel 'T - Information Assets' niet gevonden in de database.", vbExclamation, "GRC Import"
        conn.Close: Application.ScreenUpdating = True: Exit Sub
    End If
    On Error GoTo 0

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Information Assets")
    ws.Range("B6:E105").ClearContents   ' Naam, Omschrijving, Eigenaar, Dienst
    ws.Range("F6:F105").ClearContents   ' Confidentialiteit code (col 6)
    ws.Range("H6:H105").ClearContents   ' Opmerkingen (col 8)

    Dim destRow As Long
    destRow = 6
    Do While Not rs.EOF And destRow <= 105
        ws.Cells(destRow, 2) = FieldVal(rs, "naam", "assetnaam", "asset naam", "name", "asset name", "nom", "actif", "titel", "title")
        ws.Cells(destRow, 3) = FieldVal(rs, "omschrijving", "beschrijving", "description", "beschr", "omschr", "toelichting", "detail")
        ws.Cells(destRow, 4) = FieldVal(rs, "eigenaar", "asseteigenaar", "asset eigenaar", "owner", "propri" & Chr(233) & "taire", "verantwoordelijke")
        ws.Cells(destRow, 5) = FieldVal(rs, "dienst", "afdeling", "organisatie", "entiteit", "department", "service", "business unit")
        ws.Cells(destRow, 6) = MapCls(FieldVal(rs, "confidentialiteit", "confidentiality", "confidentialit" & Chr(233), "conf"))
        destRow = destRow + 1
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
    Application.ScreenUpdating = True
    MsgBox destRow - 6 & " informatieassets ge" & Chr(239) & "mporteerd.", vbInformation, "GRC Import"
End Sub

' Exporteer alle gegevens naar een nieuw Excel-bestand
Sub ExporteerAlles()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    fd.Title = "Sla de export op als..."
    fd.InitialFileName = "GRC_Export_" & Format(Now, "YYYYMMDD") & ".xlsx"
    If fd.Show = False Then Exit Sub

    Dim exportPath As String
    exportPath = fd.SelectedItems(1)
    If LCase(Right(exportPath, 5)) <> ".xlsx" Then exportPath = exportPath & ".xlsx"

    Application.ScreenUpdating = False
    Dim expWb As Workbook
    Set expWb = Workbooks.Add

    Dim sheetsToCopy As Variant
    sheetsToCopy = Array("Processes", "Information Assets", "Dependent Assets", "Responsible Persons")
    Dim sName As Variant
    For Each sName In sheetsToCopy
        On Error Resume Next
        Dim srcSh As Worksheet
        Set srcSh = ThisWorkbook.Sheets(CStr(sName))
        On Error GoTo 0
        If Not srcSh Is Nothing Then
            srcSh.Copy After:=expWb.Sheets(expWb.Sheets.Count)
            expWb.Sheets(expWb.Sheets.Count).Name = CStr(sName)
        End If
    Next sName

    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In expWb.Worksheets
        If ws.Name = "Sheet1" Or ws.Name = "Blad1" Or ws.Name = "Feuil1" Then ws.Delete
    Next ws
    Application.DisplayAlerts = True

    expWb.SaveAs exportPath, FileFormat:=xlOpenXMLWorkbook
    expWb.Close False
    Application.ScreenUpdating = True
    MsgBox "Export opgeslagen:" & vbCrLf & exportPath, vbInformation, "GRC Export"
End Sub

' Refresh Kwetsbaarheden-matrix vanuit Access.
' FASE 1: kwetsbaarheid-namen + CIA-impact bijwerken op bestaande kolompositie.
' FASE 2: voor elke kwetsbaarheid-kolom alle remediërende controls opzoeken en markeren.
Private Sub CoreImportKwetsbaarheden(conn As Object)
    Const COL_ID      As Integer = 1
    Const COL_TITLE   As Integer = 2
    Const COL_V_START As Integer = 3
    Const ROW_TITLE   As Integer = 1
    Const ROW_VULN    As Integer = 2
    Const ROW_C       As Integer = 3
    Const ROW_I       As Integer = 4
    Const ROW_A       As Integer = 5
    Const ROW_DATA    As Integer = 6
    Const COL_CF25    As Integer = 6
    Const COL_CF23    As Integer = 13
    Const ROW_CF_DATA As Integer = 4

    Dim ws As Worksheet, wsCC As Worksheet, wsAfw As Worksheet
    Dim rs As Object, rsCtrl As Object, rsLT As Object
    Dim rr As Long, v As Integer
    Dim ctrlRowMap As Object, refNrMap As Object
    Dim rev23to25 As Object, enkel23Set As Object, afwList As Object
    Dim lastRow As Long, nVuln As Integer, nMatched As Long

    Set ws = ThisWorkbook.Sheets("Kwetsbaarheden")

    ' ── 1. Ctrl-rij-map: LCase(ID) → rijnummer ───────────────────────────────
    Set ctrlRowMap = CreateObject("Scripting.Dictionary")
    rr = ROW_DATA
    Do While ws.Cells(rr, COL_ID).Value <> ""
        Dim ck As String
        ck = LCase(Trim(CStr(ws.Cells(rr, COL_ID).Value)))
        If Not ctrlRowMap.Exists(ck) Then ctrlRowMap.Add ck, rr
        rr = rr + 1
    Loop
    If ctrlRowMap.Count = 0 Then
        MsgBox "Kwetsbaarheden-sheet bevat geen controls (rij 6+ van kolom A is leeg)." & vbCrLf & _
               "Herbouw het bestand via build_grc.py en probeer opnieuw.", vbExclamation, "GRC Import"
        Exit Sub
    End If
    lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row

    ' ── 2. 2023→2025 reverse mapping + "Enkel 2023"-set via CyFun Controls-sheet ──
    Set rev23to25  = CreateObject("Scripting.Dictionary")
    Set enkel23Set = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set wsCC = ThisWorkbook.Sheets("CyFun Controls")
    On Error GoTo 0
    If Not wsCC Is Nothing Then
        Dim ccLast As Long
        ccLast = Application.Max( _
            wsCC.Cells(wsCC.Rows.Count, COL_CF25).End(xlUp).Row, _
            wsCC.Cells(wsCC.Rows.Count, COL_CF23).End(xlUp).Row)
        Dim rCC As Long
        For rCC = ROW_CF_DATA To ccLast
            Dim r25t As String, r23t As String
            r25t = Trim(CStr(wsCC.Cells(rCC, COL_CF25).Value))
            r23t = Trim(CStr(wsCC.Cells(rCC, COL_CF23).Value))
            If r25t <> "" And r23t <> "" Then
                Dim p25 As Integer, p23 As Integer
                p25 = InStr(r25t, ":"): p23 = InStr(r23t, ":")
                Dim i25 As String, i23 As String
                i25 = LCase(Trim(Replace(IIf(p25 > 0, Left(r25t, p25 - 1), r25t), Chr(9), "")))
                i23 = NormId23(r23t)
                If i23 <> "" And i25 <> "" And Not rev23to25.Exists(i23) Then
                    rev23to25.Add i23, i25
                End If
            ElseIf r25t = "" And r23t <> "" Then
                Dim e23 As String
                e23 = NormId23(r23t)
                If e23 <> "" And Not enkel23Set.Exists(e23) Then
                    enkel23Set.Add e23, True
                End If
            End If
        Next rCC
    End If

    ' ── 3. RefNr → CyFun 2023 ID via T-CyFunEssentiel ────────────────────────
    Set refNrMap = CreateObject("Scripting.Dictionary")
    Set rsCtrl = CreateObject("ADODB.Recordset")
    On Error Resume Next
    rsCtrl.Open "SELECT RefNr, Requirement FROM [T - CyFunEssentiel]", conn
    If Err.Number <> 0 Then
        MsgBox "Tabel 'T - CyFunEssentiel' niet gevonden in de database.", vbExclamation, "GRC Import"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    Do While Not rsCtrl.EOF
        Dim refNr As Long
        refNr = CLng(rsCtrl.Fields("RefNr").Value)
        Dim reqVal As String
        reqVal = CStr(rsCtrl.Fields("Requirement").Value)
        If Not refNrMap.Exists(refNr) Then refNrMap.Add refNr, reqVal
        rsCtrl.MoveNext
    Loop
    rsCtrl.Close

    ' ══════════════════════════════════════════════════════════════════════════
    ' FASE 1 — Kwetsbaarheden en CIA-impact importeren op bestaande positie
    ' ══════════════════════════════════════════════════════════════════════════

    ' ── 4. Scan bestaande kwetsbaarheid-kolommen (rij 2, C+) ─────────────────
    ' existingVulnCols: LCase(naam) → kolomnummer (van vorige import)
    Dim existingVulnCols As Object
    Set existingVulnCols = CreateObject("Scripting.Dictionary")
    Dim usedEndCol As Long
    usedEndCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    Dim scanCol As Long
    For scanCol = COL_V_START To usedEndCol
        Dim scanHdr As String
        scanHdr = Trim(CStr(ws.Cells(ROW_VULN, scanCol).Value))
        If scanHdr <> "" Then
            Dim scanKey As String
            scanKey = LCase(scanHdr)
            If Not existingVulnCols.Exists(scanKey) Then
                existingVulnCols.Add scanKey, scanCol
            End If
        End If
    Next scanCol

    ' Bepaal startkolom voor nieuwe kwetsbaarheden
    Dim nextNewCol As Long
    If existingVulnCols.Count > 0 Then
        Dim maxExistCol As Long: maxExistCol = COL_V_START - 1
        Dim eKey As Variant
        For Each eKey In existingVulnCols.Keys
            Dim ec As Long: ec = CLng(existingVulnCols.Item(eKey))
            If ec > maxExistCol Then maxExistCol = ec
        Next eKey
        nextNewCol = maxExistCol + 1
    Else
        nextNewCol = COL_V_START
    End If

    ' ── 5. Kwetsbaarheden laden; bijwerken op bestaande positie of nieuw toevoegen ──
    Dim vulnIDs()   As Long
    Dim vulnNames() As String
    Dim vulnC()     As Boolean, vulnI() As Boolean, vulnA() As Boolean
    Dim vulnCols()  As Long
    Dim matchedCols As Object  ' bijgehouden kolomnummers (voor orphan-detectie)
    Set matchedCols = CreateObject("Scripting.Dictionary")
    nVuln = 0
    ReDim vulnIDs(0 To 99):   ReDim vulnNames(0 To 99)
    ReDim vulnC(0 To 99):     ReDim vulnI(0 To 99):   ReDim vulnA(0 To 99)
    ReDim vulnCols(0 To 99)

    Set rs = CreateObject("ADODB.Recordset")
    On Error Resume Next
    rs.Open "SELECT Reference, Vulnerability, Confidentialiteit, Integriteit, " & _
            "Beschikbaarheid FROM [T - Vulnerabilities] ORDER BY Reference", conn
    If Err.Number <> 0 Then
        MsgBox "Tabel 'T - Vulnerabilities' niet gevonden of ongeldige query." & vbCrLf & _
               Err.Description, vbExclamation, "GRC Import"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Do While Not rs.EOF And nVuln < 100
        Dim vID As Long, vNm As String
        Dim vC As Boolean, vI As Boolean, vA As Boolean
        vID = CLng(rs.Fields("Reference").Value)
        vNm = CStr(rs.Fields("Vulnerability").Value)
        vC  = CBool(rs.Fields("Confidentialiteit").Value)
        vI  = CBool(rs.Fields("Integriteit").Value)
        vA  = CBool(rs.Fields("Beschikbaarheid").Value)

        ' Gebruik bestaande kolom (op naam) of maak nieuwe aan
        Dim targetCol As Long
        Dim vNmKey As String: vNmKey = LCase(vNm)
        If existingVulnCols.Exists(vNmKey) Then
            targetCol = CLng(existingVulnCols.Item(vNmKey))
        Else
            targetCol = nextNewCol
            nextNewCol = nextNewCol + 1
            ws.Columns(targetCol).ColumnWidth = 10
        End If

        vulnIDs(nVuln)   = vID
        vulnNames(nVuln) = vNm
        vulnC(nVuln) = vC: vulnI(nVuln) = vI: vulnA(nVuln) = vA
        vulnCols(nVuln)  = targetCol
        If Not matchedCols.Exists(targetCol) Then matchedCols.Add targetCol, True

        ' Rij 2: naam (bijwerken)
        With ws.Cells(ROW_VULN, targetCol)
            .Value = vNm
            .Interior.Color = IIf(nVuln Mod 2 = 0, RGB(208, 220, 243), RGB(255, 255, 255))
            .Font.Bold = True: .Font.Size = 9
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        End With

        ' Rijen 3–5: CIA-impact (bijwerken)
        Dim cia As Integer
        Dim ciaRows(1 To 3) As Integer
        Dim ciaFlags(1 To 3) As Boolean
        ciaRows(1) = ROW_C: ciaFlags(1) = vC
        ciaRows(2) = ROW_I: ciaFlags(2) = vI
        ciaRows(3) = ROW_A: ciaFlags(3) = vA
        Dim ciaColors(1 To 3) As Long
        ciaColors(1) = RGB(173, 216, 230)
        ciaColors(2) = RGB(198, 239, 206)
        ciaColors(3) = RGB(255, 235, 156)
        For cia = 1 To 3
            With ws.Cells(ciaRows(cia), targetCol)
                If ciaFlags(cia) Then
                    .Value = ChrW(10004)
                    .Interior.Color = ciaColors(cia)
                    .Font.Bold = True: .Font.Size = 9
                Else
                    .Value = ""
                    .Interior.Color = RGB(217, 217, 217)
                End If
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
            End With
        Next cia

        ' Wis bestaande ✔-marks (worden in fase 2 opnieuw ingevuld)
        If lastRow >= ROW_DATA Then
            ws.Range(ws.Cells(ROW_DATA, targetCol), _
                     ws.Cells(lastRow, targetCol)).ClearContents
        End If

        nVuln = nVuln + 1
        rs.MoveNext
    Loop
    rs.Close

    ' ── 6. Verwijder orphan-kolommen (bestonden maar niet meer in database) ───
    ' Verzamel orphan-kolommen (in existingVulnCols maar NIET in matchedCols)
    Dim orphanCols() As Long
    Dim nOrphan As Integer: nOrphan = 0
    Dim exKey2 As Variant
    For Each exKey2 In existingVulnCols.Keys
        Dim exColNr As Long: exColNr = CLng(existingVulnCols.Item(exKey2))
        If Not matchedCols.Exists(exColNr) Then
            ReDim Preserve orphanCols(0 To nOrphan)
            orphanCols(nOrphan) = exColNr
            nOrphan = nOrphan + 1
        End If
    Next exKey2
    ' Sorteer descenderend (bubble sort) zodat rechts-naar-links wordt gewist
    Dim si As Integer, sj As Integer, tmpCol As Long
    For si = 0 To nOrphan - 2
        For sj = 0 To nOrphan - 2 - si
            If orphanCols(sj) < orphanCols(sj + 1) Then
                tmpCol = orphanCols(sj): orphanCols(sj) = orphanCols(sj + 1): orphanCols(sj + 1) = tmpCol
            End If
        Next sj
    Next si
    ' Verwijder orphan-kolommen en pas vulnCols aan voor kolom-verschuiving
    Dim oi As Integer, adj As Integer
    For oi = 0 To nOrphan - 1
        Dim delCol As Long: delCol = orphanCols(oi)
        ws.Columns(delCol).Delete
        For adj = 0 To nVuln - 1
            If vulnCols(adj) > delCol Then vulnCols(adj) = vulnCols(adj) - 1
        Next adj
    Next oi

    ' ── 7. Titelrij samenvoegen tot laatste kwetsbaarheid-kolom ──────────────
    Dim lastVulnCol As Long: lastVulnCol = COL_V_START - 1
    For v = 0 To nVuln - 1
        If vulnCols(v) > lastVulnCol Then lastVulnCol = vulnCols(v)
    Next v
    On Error Resume Next
    ws.Cells(ROW_TITLE, 1).MergeArea.UnMerge
    On Error GoTo 0
    If lastVulnCol >= COL_V_START Then
        ws.Range(ws.Cells(ROW_TITLE, 1), ws.Cells(ROW_TITLE, lastVulnCol)).Merge
        ws.Cells(ROW_TITLE, 1).HorizontalAlignment = xlLeft
    End If

    ' ══════════════════════════════════════════════════════════════════════════
    ' FASE 2 — Alle kolommen overlopen: controls per kwetsbaarheid opzoeken en markeren
    ' ══════════════════════════════════════════════════════════════════════════

    ' ── 8. LT-tabel in geheugen laden: vulnID → set van ctrlRefs ─────────────
    Set afwList = CreateObject("Scripting.Dictionary")
    nMatched = 0

    Dim vulnLTMap As Object
    Set vulnLTMap = CreateObject("Scripting.Dictionary")

    Set rsLT = CreateObject("ADODB.Recordset")
    On Error Resume Next
    rsLT.Open "SELECT Vulnerability, CyFunControl " & _
              "FROM [LT -  Vulnerability to control - fixed]", conn
    If Err.Number <> 0 Then
        MsgBox "Fout bij openen LT-tabel: " & Err.Description, vbExclamation, "GRC Import"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Do While Not rsLT.EOF
        Dim ltVId As Long, ltCRef As Long
        ltVId  = CLng(rsLT.Fields("Vulnerability").Value)
        ltCRef = CLng(rsLT.Fields("CyFunControl").Value)
        If Not vulnLTMap.Exists(ltVId) Then
            vulnLTMap.Add ltVId, CreateObject("Scripting.Dictionary")
        End If
        Dim ltDict As Object
        Set ltDict = vulnLTMap.Item(ltVId)
        If Not ltDict.Exists(ltCRef) Then ltDict.Add ltCRef, True
        rsLT.MoveNext
    Loop
    rsLT.Close

    ' ── 9. Per kwetsbaarheid: popup tonen en controls markeren ───────────────
    For v = 0 To nVuln - 1
        Application.StatusBar = "Kwetsbaarheden importeren... (" & (v + 1) & "/" & nVuln & "): " & vulnNames(v)
        DoEvents

        If vulnLTMap.Exists(vulnIDs(v)) Then
            Dim vLTDict As Object
            Set vLTDict = vulnLTMap.Item(vulnIDs(v))
            Dim ltRef As Variant
            For Each ltRef In vLTDict.Keys
                Dim ltCtrlRef As Long: ltCtrlRef = CLng(ltRef)
                If refNrMap.Exists(ltCtrlRef) Then
                    Dim id23v As String
                    id23v = NormId23(CStr(refNrMap.Item(ltCtrlRef)))

                    Dim id25v As String
                    If rev23to25.Exists(id23v) Then
                        id25v = CStr(rev23to25.Item(id23v))
                    Else
                        id25v = id23v
                    End If

                    Dim ctrlRow As Long: ctrlRow = 0
                    If ctrlRowMap.Exists(id25v) Then
                        ctrlRow = CLng(ctrlRowMap.Item(id25v))
                        nMatched = nMatched + 1
                        With ws.Cells(ctrlRow, vulnCols(v))
                            .Value = ChrW(10004)
                            .Interior.Color = RGB(198, 239, 206)
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                            .Font.Size = 10
                        End With
                    Else
                        Dim afwK As String
                        afwK = vulnNames(v) & "|" & id23v
                        If Not afwList.Exists(afwK) Then
                            Dim cs As String: cs = ""
                            If vulnC(v) Then cs = cs & "C"
                            If vulnI(v) Then cs = cs & "I"
                            If vulnA(v) Then cs = cs & "A"
                            Dim reden As String
                            If enkel23Set.Exists(id23v) Then
                                reden = "Enkel in 2023"
                            Else
                                reden = "Niet gevonden"
                            End If
                            afwList.Add afwK, Array(vulnNames(v), cs, id23v, reden)
                        End If
                    End If
                End If
            Next ltRef
        End If
    Next v
    Application.StatusBar = False

    ' ── 9. Afwijkingen-sheet bijwerken ────────────────────────────────────────
    On Error Resume Next
    Set wsAfw = ThisWorkbook.Sheets("Afwijkingen")
    On Error GoTo 0
    If wsAfw Is Nothing Then
        Set wsAfw = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsAfw.Name = "Afwijkingen"
    End If

    Dim aLastRow As Long
    aLastRow = wsAfw.Cells(wsAfw.Rows.Count, 1).End(xlUp).Row
    If aLastRow >= 3 Then wsAfw.Rows("3:" & aLastRow).ClearContents

    Dim aRow As Long
    aRow = 3
    Dim afwKey2 As Variant
    For Each afwKey2 In afwList.Keys
        Dim info As Variant
        info = afwList.Item(afwKey2)
        wsAfw.Cells(aRow, 1).Value = info(0)
        wsAfw.Cells(aRow, 2).Value = info(1)
        wsAfw.Cells(aRow, 3).Value = info(2)
        wsAfw.Cells(aRow, 4).Value = info(3)
        aRow = aRow + 1
    Next afwKey2

    MsgBox "Geladen: " & nVuln & " kwetsbaarheden, " & nMatched & " matches, " & _
           afwList.Count & " afwijkingen.", vbInformation, "GRC Import"
End Sub

' Importeer kwetsbaarheden standalone (knop op Kwetsbaarheden-sheet)
Sub ImportKwetsbaarheden()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Selecteer de Access-database (kwetsbaarheden)"
    fd.Filters.Clear
    fd.Filters.Add "Access-bestanden", "*.accdb; *.mdb"
    If fd.Show = False Then Exit Sub
    Application.ScreenUpdating = False
    Dim conn As Object
    Set conn = OpenAccess(fd.SelectedItems(1))
    If conn Is Nothing Then Application.ScreenUpdating = True: Exit Sub
    CoreImportKwetsbaarheden conn
    conn.Close
    Application.ScreenUpdating = True
End Sub

' Normaliseert elke CyFun 2023 ID-string naar lowercase kort ID, ongeacht het formaat:
'   "IMPORTANT_RS.CO-3.2: The organization..." → "rs.co-3.2"
'   "RS.CO-3.2: The organization..."           → "rs.co-3.2"
'   "RS.CO-3.2.a"                              → "rs.co-3.2"
'   "rs.co-3-2"                                → "rs.co-3.2"
'   "RS.CO-3.2"                                → "rs.co-3.2"
Function NormId23(s As String) As String
    Dim r As String
    r = Trim(s)

    ' Stap 1: strip niveau-prefix (IMPORTANT_, BASIC_, ESSENTIAL_)
    If Left(LCase(r), 10) = "important_" Then
        r = Mid(r, 11)
    ElseIf Left(LCase(r), 6) = "basic_" Then
        r = Mid(r, 7)
    ElseIf Left(LCase(r), 10) = "essential_" Then
        r = Mid(r, 11)
    End If

    ' Stap 2: strip beschrijving na dubbele punt ("RS.CO-3.2: tekst" → "RS.CO-3.2")
    Dim colon As Integer
    colon = InStr(r, ":")
    If colon > 0 Then r = Trim(Left(r, colon - 1))

    r = LCase(r)

    ' Stap 3: strip trailing letter-extensie (.a tot .z)
    Dim lastDot As Integer
    lastDot = InStrRev(r, ".")
    If lastDot > 0 Then
        Dim suf As String
        suf = Mid(r, lastDot + 1)
        If Len(suf) = 1 And suf >= "a" And suf <= "z" Then
            r = Left(r, lastDot - 1)
        End If
    End If

    ' Stap 4: trailing -N → .N  (bv. "rs.co-3-2" → "rs.co-3.2")
    Dim i As Integer, lastHyph As Integer
    lastHyph = 0
    For i = Len(r) To 1 Step -1
        Dim ch As String
        ch = Mid(r, i, 1)
        If ch >= "0" And ch <= "9" Then
            ' nog in cijfer-zone, blijf zoeken
        ElseIf ch = "-" Then
            lastHyph = i
            Exit For
        Else
            Exit For
        End If
    Next i
    If lastHyph > 0 Then
        Dim tail As String
        tail = Mid(r, lastHyph + 1)
        If Len(tail) > 0 Then
            r = Left(r, lastHyph - 1) & "." & tail
        End If
    End If

    NormId23 = r
End Function
'''

# ══════════════════════════════════════════════════════════════════════════════
# WERKBOEK OPBOUWEN
# ══════════════════════════════════════════════════════════════════════════════
wb = Workbook()

# ─────────────────────────────────────────────────────────────────────────────
# _Lang  (verborgen vertaalblad + benoemde bereiken)
# ─────────────────────────────────────────────────────────────────────────────
ws_lang = wb.active
ws_lang.title = "_Lang"

for col, h in enumerate(["Sleutel", "NL", "FR", "EN"], 1):
    c = ws_lang.cell(row=1, column=col, value=h)
    c.font = Font(name="Calibri", bold=True, color=C["navy"])
    c.fill = PatternFill("solid", fgColor=C["grey_mid"])

for r, row_data in enumerate(TRANSLATIONS, 2):
    for col, val in enumerate(row_data, 1):
        ws_lang.cell(row=r, column=col, value=val)

ws_lang.column_dimensions["A"].width = 28
for letter in ["B", "C", "D"]:
    ws_lang.column_dimensions[letter].width = 72

# Classificatiewaarden F=NL, G=FR, H=EN  — 5 niveaus
CLS_START   = 250
ROLES_START = 258
TYPES_START = 270

for col_off, lang in enumerate(["NL", "FR", "EN"]):
    col = 6 + col_off
    for row_off, val in enumerate(CLS[lang]):
        ws_lang.cell(row=CLS_START + row_off, column=col, value=val)
    for row_off, val in enumerate(ROLES[lang]):
        ws_lang.cell(row=ROLES_START + row_off, column=col, value=val)
    for row_off, val in enumerate(ASSET_TYPES[lang]):
        ws_lang.cell(row=TYPES_START + row_off, column=col, value=val)

ws_lang.sheet_state = "hidden"

# Named ranges
n_cls   = len(CLS["NL"])    # 5
n_roles = len(ROLES["NL"])  # 8
n_types = len(ASSET_TYPES["NL"])  # 10

for col_off, lang in enumerate(["NL", "FR", "EN"]):
    col = 6 + col_off
    cl = get_column_letter(col)
    wb.defined_names[f"cls_{lang}"]   = DefinedName(f"cls_{lang}",   attr_text=f"_Lang!${cl}${CLS_START}:${cl}${CLS_START+n_cls-1}")
    wb.defined_names[f"roles_{lang}"] = DefinedName(f"roles_{lang}", attr_text=f"_Lang!${cl}${ROLES_START}:${cl}${ROLES_START+n_roles-1}")
    wb.defined_names[f"types_{lang}"] = DefinedName(f"types_{lang}", attr_text=f"_Lang!${cl}${TYPES_START}:${cl}${TYPES_START+n_types-1}")

# ── Layout-hulpfuncties ───────────────────────────────────────────────────────
def sheet_title_bar(ws, text, n_cols):
    last = get_column_letter(n_cols)
    ws.merge_cells(f"A1:{last}1");  ws.row_dimensions[1].height = 52
    c = ws["A1"]; c.value = text
    c.fill = fill("navy"); c.font = Font(name="Calibri", bold=True, size=20, color=C["white"])
    c.alignment = align("center", "center")
    ws.merge_cells(f"A2:{last}2"); ws.row_dimensions[2].height = 22
    c = ws["A2"]; c.value = t("app_tagline")
    c.fill = fill("blue_mid"); c.font = Font(name="Calibri", size=11, color=C["blue_xlight"], italic=True)
    c.alignment = align("center", "center")

def section_header(ws, row, key, n_cols, col_start=1):
    ws.row_dimensions[row].height = 20
    lc = get_column_letter(col_start + n_cols - 1)
    ws.merge_cells(f"{get_column_letter(col_start)}{row}:{lc}{row}")
    c = ws.cell(row=row, column=col_start, value=t(key))
    c.fill = fill("navy"); c.font = Font(name="Calibri", bold=True, size=9, color=C["blue_light"])
    c.alignment = align("left", "center", indent=1)

def label_value_row(ws, row, col_lbl, col_val, label, value,
                    val_bold=False, val_color="text", val_size=10, dropdown=None, editable=True):
    ws.row_dimensions[row].height = 26
    lc = ws.cell(row=row, column=col_lbl, value=label)
    lc.font = font(10, bold=True, color="grey_dark"); lc.alignment = align("right", "center")
    lc.fill = fill("grey_light"); lc.border = border_all()
    vc = ws.cell(row=row, column=col_val, value=value)
    vc.font = Font(name="Calibri", size=val_size, bold=val_bold, color=C[val_color])
    vc.alignment = align("left", "center", indent=1)
    vc.fill = fill("white") if editable else fill("grey_light")
    vc.border = border_all("blue_mid" if editable else "grey_border")
    if dropdown:
        dv = DataValidation(type="list", formula1=dropdown, showDropDown=False, allow_blank=False)
        dv.sqref = f"{get_column_letter(col_val)}{row}"
        ws.add_data_validation(dv)
    return vc

def hint_row(ws, row, n_cols, key, col_start=1):
    ws.row_dimensions[row].height = 18
    lc = get_column_letter(col_start + n_cols - 1)
    ws.merge_cells(f"{get_column_letter(col_start)}{row}:{lc}{row}")
    c = ws.cell(row=row, column=col_start, value=t(key))
    c.font = Font(name="Calibri", size=9, color=C["subtext"], italic=True)
    c.alignment = align("left", "center", indent=1)

def quicklinks_block(ws, row_start, links, n_cols, col_start=1):
    """Renders a SNELKOPPELINGEN block starting at row_start."""
    ws.row_dimensions[row_start].height = 10
    r = row_start + 1
    ws.row_dimensions[r].height = 20
    lc = get_column_letter(col_start + n_cols - 1)
    ws.merge_cells(f"{get_column_letter(col_start)}{r}:{lc}{r}")
    c = ws.cell(row=r, column=col_start, value=t("ui_quicklinks"))
    c.fill = fill("blue_xlight"); c.font = Font(name="Calibri", bold=True, size=9, color=C["blue_mid"])
    c.alignment = align("left", "center", indent=1)
    for key, target in links:
        r += 1
        ws.row_dimensions[r].height = 20
        ws.merge_cells(f"{get_column_letter(col_start)}{r}:{lc}{r}")
        c = ws.cell(row=r, column=col_start, value=t(key))
        c.hyperlink = f"#{target}"; c.fill = fill("blue_xlight")
        c.font = Font(name="Calibri", size=10, color=C["blue_mid"], underline="single")
        c.alignment = align("left", "center", indent=1)

def build_cls_sheet(ws, title_key, subtitle_key, id_prefix,
                    col_headers_before, col_headers_after,
                    widths, data_start=6, data_end=105):
    """
    Bouwt een gegevensblad met classificatiekolommen (code + label).
    col_headers_before : list of (key, col_idx)  — kolommen voor de classificaties
    col_headers_after  : list of (key, col_idx)  — kolommen na de classificaties
    widths             : list of column widths
    Returns: (HDR_ROW, DATA_START, DATA_END, CLS_PAIRS, N)
    """
    N = len(widths)
    set_col_widths(ws, widths)
    no_gridlines(ws)
    sheet_title_bar(ws, t(title_key), N)
    ws.row_dimensions[3].height = 22
    ws.merge_cells(f"A3:{get_column_letter(N)}3")
    c = ws["A3"]; c.value = t(subtitle_key)
    c.fill = fill("blue_xlight"); c.font = font(10, italic=True, color="grey_dark")
    c.alignment = align("center", "center")
    ws.row_dimensions[4].height = 8

    HDR_ROW = 5
    ws.row_dimensions[HDR_ROW].height = 44

    # Enkelvoudige koppen
    for key, col_idx in col_headers_before + col_headers_after:
        c = ws.cell(row=HDR_ROW, column=col_idx, value=t(key))
        c.fill = fill("navy"); c.font = Font(name="Calibri", bold=True, size=10, color=C["white"])
        c.alignment = align("center", "center", wrap=True); c.border = border_all("navy")

    return HDR_ROW, data_start, data_end, N

# ── Gepaarde classificatiekolommen ───────────────────────────────────────────
def add_cls_pairs(ws, cls_pairs, HDR_ROW, DATA_START, DATA_END, max_level=5):
    """Voegt merged koppen, formules, dropdowns en CF toe voor de classificatiekolommen."""
    for key, code_col, label_col in cls_pairs:
        cl = get_column_letter(code_col); ll = get_column_letter(label_col)
        ws.merge_cells(f"{cl}{HDR_ROW}:{ll}{HDR_ROW}")
        c = ws.cell(row=HDR_ROW, column=code_col, value=t(key))
        c.fill = fill("navy"); c.font = Font(name="Calibri", bold=True, size=10, color=C["white"])
        c.alignment = align("center", "center", wrap=True); c.border = border_all("navy")

    for r in range(DATA_START, DATA_END + 1):
        alt = "blue_xlight" if r % 2 == 0 else "white"
        for _, code_col, label_col in cls_pairs:
            cl = get_column_letter(code_col); ll = get_column_letter(label_col)
            cc = ws.cell(row=r, column=code_col)
            cc.alignment = align("center", "center")
            cc.font = Font(name="Calibri", bold=True, size=11, color=C["subtext"])
            lc = ws.cell(row=r, column=label_col)
            lc.value = (
                f'=IF({cl}{r}="","",IFERROR('
                f'INDEX(\'Reference Values\'!$B$8:$B$12,{cl}{r}),"?"))'
            )
            lc.alignment = align("center", "center")
            lc.font = Font(name="Calibri", size=10, italic=True)
            lc.fill = fill("grey_light")

    # Dropdown beperkt tot max_level
    levels_str = ",".join(str(i) for i in range(1, max_level + 1))
    dv = DataValidation(type="list", formula1=f'"{levels_str}"', showDropDown=False, allow_blank=True)
    dv.sqref = " ".join(
        f"{get_column_letter(cc)}{DATA_START}:{get_column_letter(cc)}{DATA_END}"
        for _, cc, _ in cls_pairs
    )
    ws.add_data_validation(dv)

    # Voorwaardelijke opmaak (numeriek, beperkt tot max_level)
    LEVEL_CF = [
        (1, C["green_light"],  C["green"]),
        (2, C["yellow_light"], C["yellow"]),
        (3, C["orange_light"], C["orange"]),
        (4, C["red_light"],    C["red"]),
        (5, C["purple_light"], C["purple"]),
    ]
    for _, code_col, label_col in cls_pairs:
        cl = get_column_letter(code_col); ll = get_column_letter(label_col)
        for lvl, bg, fg in [x for x in LEVEL_CF if x[0] <= max_level]:
            formula = f"{cl}{DATA_START}={lvl}"
            for col_ltr in [cl, ll]:
                rng = f"{col_ltr}{DATA_START}:{col_ltr}{DATA_END}"
                ws.conditional_formatting.add(rng, FormulaRule(
                    formula=[formula],
                    fill=PatternFill("solid", fgColor=bg),
                    font=Font(name="Calibri", bold=True, size=10, color=fg),
                ))


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Config
# ══════════════════════════════════════════════════════════════════════════════
ws_cfg = wb.create_sheet("Config")
no_gridlines(ws_cfg); set_col_widths(ws_cfg, [2, 32, 3, 30, 3, 26, 26])
sheet_title_bar(ws_cfg, t("cfg_title_main"), 7)
ws_cfg.row_dimensions[3].height = 12

section_header(ws_cfg, 4, "cfg_section_general", 5, col_start=2)
label_value_row(ws_cfg, 5, 2, 4, t("cfg_org_label"), "Mijn Overheidsinstelling")
label_value_row(ws_cfg, 6, 2, 4, t("cfg_dept_label"), "")
ws_cfg.row_dimensions[7].height = 8

section_header(ws_cfg, 8, "cfg_section_display", 5, col_start=2)
lv = label_value_row(ws_cfg, 9, 2, 4, t("cfg_lang_label"), "NL",
                     val_bold=True, val_color="blue_mid", val_size=14, dropdown='"NL,FR,EN"')
lv.alignment = align("center", "center")
hint_row(ws_cfg, 10, 5, "cfg_lang_hint",        col_start=2)
hint_row(ws_cfg, 11, 5, "cfg_lang_change_hint", col_start=2)
ws_cfg.row_dimensions[12].height = 8

section_header(ws_cfg, 13, "cfg_section_meta", 5, col_start=2)
label_value_row(ws_cfg, 14, 2, 4, t("cfg_version_label"),    "0.3", editable=False)
label_value_row(ws_cfg, 15, 2, 4, t("cfg_updated_label"),    NOW_STR, editable=False)
label_value_row(ws_cfg, 16, 2, 4, t("cfg_updated_by_label"), USERNAME)
ws_cfg.row_dimensions[17].height = 8

# Tips
ws_cfg.row_dimensions[18].height = 20; ws_cfg.merge_cells("B18:F18")
c = ws_cfg.cell(row=18, column=2, value=t("ui_tips"))
c.fill = fill("accent_light"); c.font = Font(name="Calibri", bold=True, size=9, color=C["accent"])
c.alignment = align("left", "center", indent=1)
for tip_r, tip_k in [(19, "cfg_save_hint"), (20, "cfg_lang_change_hint")]:
    ws_cfg.row_dimensions[tip_r].height = 20; ws_cfg.merge_cells(f"B{tip_r}:F{tip_r}")
    c = ws_cfg.cell(row=tip_r, column=2, value=t(tip_k))
    c.fill = fill("accent_light"); c.font = Font(name="Calibri", size=9, color=C["grey_dark"])
    c.alignment = align("left", "center", indent=1)

quicklinks_block(ws_cfg, 21,
    [("ui_goto_info",    "Info!A1"),
     ("ui_goto_proc",   "Processes!A1"),
     ("ui_goto_assets", "Information Assets!A1"),
     ("ui_goto_dep",    "Dependent Assets!A1"),
     ("ui_goto_verant", "Responsible Persons!A1"),
     ("ui_goto_import", "Import & Export!A1"),
     ("ui_goto_ref",    "Reference Values!A1")], 5, col_start=2)


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Info
# ══════════════════════════════════════════════════════════════════════════════
ws_info = wb.create_sheet("Info")
no_gridlines(ws_info); set_col_widths(ws_info, [2, 22, 4, 58, 8])
sheet_title_bar(ws_info, t("info_title"), 5)
ws_info.row_dimensions[3].height = 12

def info_section(ws, row, key): section_header(ws, row, key, 4, col_start=2)
def info_text(ws, row, key, height=52, bg="white"):
    ws.row_dimensions[row].height = height; ws.merge_cells(f"B{row}:D{row}")
    c = ws.cell(row=row, column=2, value=t(key))
    c.fill = fill(bg); c.font = font(10, color="text")
    c.alignment = align("left", "top", wrap=True, indent=1); c.border = border_all()

info_section(ws_info, 4, "info_section_doel");   info_text(ws_info, 5, "info_doel_txt", 60)
ws_info.row_dimensions[6].height = 8
info_section(ws_info, 7, "info_section_scope");  info_text(ws_info, 8, "info_scope_txt", 44)
ws_info.row_dimensions[9].height = 8
info_section(ws_info, 10, "info_section_modules")
for i, key in enumerate(["info_mod_proc", "info_mod_assets", "info_mod_verant", "info_mod_import"], 11):
    info_text(ws_info, i, key, 28, "blue_xlight")
ws_info.row_dimensions[15].height = 8
info_section(ws_info, 16, "info_section_guide")
for i, key in enumerate(["info_guide_1","info_guide_2","info_guide_3","info_guide_4"], 17):
    info_text(ws_info, i, key, 28, "grey_light")
ws_info.row_dimensions[21].height = 8
info_section(ws_info, 22, "info_section_version")
for i, key in enumerate(["info_v_tool","info_v_status","info_v_frameworks"], 23):
    ws_info.row_dimensions[i].height = 22; ws_info.merge_cells(f"B{i}:D{i}")
    c = ws_info.cell(row=i, column=2, value=t(key))
    c.fill = fill("grey_light"); c.font = font(10, color="subtext")
    c.alignment = align("left", "center", indent=1); c.border = border_all()

quicklinks_block(ws_info, 26,
    [("ui_goto_config",  "Config!A1"),
     ("ui_goto_proc",   "Processes!A1"),
     ("ui_goto_assets", "Information Assets!A1"),
     ("ui_goto_dep",    "Dependent Assets!A1"),
     ("ui_goto_verant", "Responsible Persons!A1"),
     ("ui_goto_import", "Import & Export!A1"),
     ("ui_goto_ref",    "Reference Values!A1")], 4, col_start=2)


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Processen
# ══════════════════════════════════════════════════════════════════════════════
ws_proc = wb.create_sheet("Processes")

PROC_WIDTHS = [5, 20, 36, 18, 16,  4, 15,  4, 15,  26,  35,  35]
PROC_CLS_PAIRS = [
    ("proc_col_integriteit",  6,  7),
    ("proc_col_beschikbaar",  8,  9),
]
HDR_PROC, DS_PROC, DE_PROC, N_PROC = build_cls_sheet(
    ws_proc, "proc_title", "proc_subtitle", "P",
    col_headers_before=[(k, i) for i, k in enumerate(
        ["proc_col_id","proc_col_naam","proc_col_omschrijving","proc_col_eigenaar","proc_col_dienst"], 1)],
    col_headers_after=[("proc_col_opmerkingen", 10), ("proc_col_assets", 11), ("proc_col_dep_assets", 12)],
    widths=PROC_WIDTHS)

# Datarijen: basis stijl
for r in range(DS_PROC, DE_PROC + 1):
    ws_proc.row_dimensions[r].height = 24
    alt = "blue_xlight" if r % 2 == 0 else "white"
    for col_idx in range(1, N_PROC + 1):
        c = ws_proc.cell(row=r, column=col_idx)
        c.fill = fill(alt); c.font = font(10, color="text")
        c.alignment = align("left", "center", wrap=True); c.border = border_all()
    id_c = ws_proc.cell(row=r, column=1)
    id_c.value = f'=IF(B{r}="","",TEXT({r - DS_PROC + 1},"P000"))'
    id_c.font = Font(name="Calibri", size=9, bold=True, color=C["subtext"])
    id_c.alignment = align("center", "center")

add_cls_pairs(ws_proc, PROC_CLS_PAIRS, HDR_PROC, DS_PROC, DE_PROC)
ws_proc.freeze_panes = f"A{DS_PROC}"
ws_proc.auto_filter.ref = f"A{HDR_PROC}:{get_column_letter(N_PROC)}{HDR_PROC}"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Informatieassets
# ══════════════════════════════════════════════════════════════════════════════
ws_asset = wb.create_sheet("Information Assets")

ASSET_WIDTHS = [5, 20, 30, 16, 16,  4, 15,  24,  35]
ASSET_CLS_PAIRS = [
    ("asset_col_conf",  6,  7),
]
HDR_ASSET, DS_ASSET, DE_ASSET, N_ASSET = build_cls_sheet(
    ws_asset, "asset_title", "asset_subtitle", "A",
    col_headers_before=[(k, i) for i, k in enumerate(
        ["asset_col_id","asset_col_naam","asset_col_omschrijving",
         "asset_col_eigenaar","asset_col_dienst"], 1)],
    col_headers_after=[("asset_col_opmerkingen", 8), ("asset_col_processes", 9)],
    widths=ASSET_WIDTHS)



for r in range(DS_ASSET, DE_ASSET + 1):
    ws_asset.row_dimensions[r].height = 24
    alt = "blue_xlight" if r % 2 == 0 else "white"
    for col_idx in range(1, N_ASSET + 1):
        c = ws_asset.cell(row=r, column=col_idx)
        c.fill = fill(alt); c.font = font(10, color="text")
        c.alignment = align("left", "center", wrap=True); c.border = border_all()
    id_c = ws_asset.cell(row=r, column=1)
    id_c.value = f'=IF(B{r}="","",TEXT({r - DS_ASSET + 1},"A000"))'
    id_c.font = Font(name="Calibri", size=9, bold=True, color=C["subtext"])
    id_c.alignment = align("center", "center")

add_cls_pairs(ws_asset, ASSET_CLS_PAIRS, HDR_ASSET, DS_ASSET, DE_ASSET)
# Col 9 = computed door VBA (Worksheet_Activate) — markeer als read-only/grijs
for r in range(DS_ASSET, DE_ASSET + 1):
    c9 = ws_asset.cell(row=r, column=9)
    c9.fill = fill("grey_light")
    c9.font = Font(name="Calibri", size=9, italic=True, color=C["subtext"])
    c9.alignment = align("left", "top", wrap=True)
ws_asset.freeze_panes = f"A{DS_ASSET}"
ws_asset.auto_filter.ref = f"A{HDR_ASSET}:{get_column_letter(N_ASSET)}{HDR_ASSET}"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Afhankelijke assets
# ══════════════════════════════════════════════════════════════════════════════
ws_dep = wb.create_sheet("Dependent Assets")

# 25 kolommen: 1-9 basis + 10-15 Security Requirements + 16-21 Security Objectives + 22-24 Gap Analyse + 25 Commentaar
# Col 6 = Overarching (nieuw): boolean — niet-leeg betekent dat dit een organisatorische asset is.
DEP_WIDTHS = [5, 20, 30, 16, 16, 10, 24, 35, 35, 4, 15, 4, 15, 4, 15, 4, 15, 4, 15, 4, 15, 35, 35, 35, 35]
DEP_CLS_PAIRS = [
    ("dep_col_conf",  10, 11),
    ("dep_col_integ", 12, 13),
    ("dep_col_avail", 14, 15),
]
HDR_DEP, DS_DEP, DE_DEP, N_DEP = build_cls_sheet(
    ws_dep, "dep_title", "dep_subtitle", "D",
    col_headers_before=[(k, i) for i, k in enumerate(
        ["dep_col_id","dep_col_naam","dep_col_omschrijving",
         "dep_col_eigenaar","dep_col_dienst","dep_col_overarching",
         "dep_col_opmerkingen","dep_col_processes","dep_col_linked_info"], 1)],
    col_headers_after=[],
    widths=DEP_WIDTHS)

# ── Groepskoppen rij 4: Security Requirements (paars) + Security Objectives (navy)
ws_dep.row_dimensions[4].height = 22
ws_dep.merge_cells("J4:O4")
c = ws_dep.cell(row=4, column=10, value=t("dep_sec_req_header"))
c.fill = fill("purple"); c.font = Font(name="Calibri", bold=True, size=10, color=C["white"])
c.alignment = align("center", "center"); c.border = border_all("purple")

ws_dep.merge_cells("P4:U4")
c = ws_dep.cell(row=4, column=16, value=t("dep_sec_obj_header"))
c.fill = fill("navy"); c.font = Font(name="Calibri", bold=True, size=10, color=C["white"])
c.alignment = align("center", "center"); c.border = border_all("navy")

ws_dep.merge_cells("V4:X4")
c = ws_dep.cell(row=4, column=22, value=t("dep_gap_header"))
c.fill = fill("red"); c.font = Font(name="Calibri", bold=True, size=10, color=C["white"])
c.alignment = align("center", "center"); c.border = border_all("red")

for r in range(DS_DEP, DE_DEP + 1):
    ws_dep.row_dimensions[r].height = 24
    alt = "blue_xlight" if r % 2 == 0 else "white"
    for col_idx in range(1, N_DEP + 1):
        c = ws_dep.cell(row=r, column=col_idx)
        c.fill = fill(alt); c.font = font(10, color="text")
        c.alignment = align("left", "center", wrap=True); c.border = border_all()
    id_c = ws_dep.cell(row=r, column=1)
    id_c.value = f'=IF(B{r}="","","D"&TEXT({r - DS_DEP + 1},"000"))'
    id_c.font = Font(name="Calibri", size=9, bold=True, color=C["subtext"])
    id_c.alignment = align("center", "center")

# Col 6 = Overarching — centre, groen lettertype als hint
for r in range(DS_DEP, DE_DEP + 1):
    c6 = ws_dep.cell(row=r, column=6)
    c6.alignment = align("center", "center")
    c6.font = Font(name="Calibri", size=11, color=C["green"])

# Cols 8-9 = computed door VBA — grijs/italic
for r in range(DS_DEP, DE_DEP + 1):
    for col_idx in [8, 9]:
        cd = ws_dep.cell(row=r, column=col_idx)
        cd.fill = fill("grey_light")
        cd.font = Font(name="Calibri", size=9, italic=True, color=C["subtext"])
        cd.alignment = align("left", "top", wrap=True)

# Security Requirements: code+label paren via add_cls_pairs (CF + formules)
add_cls_pairs(ws_dep, [("dep_col_conf",  10, 11)], HDR_DEP, DS_DEP, DE_DEP)
add_cls_pairs(ws_dep, [("dep_col_integ", 12, 13)], HDR_DEP, DS_DEP, DE_DEP)
add_cls_pairs(ws_dep, [("dep_col_avail", 14, 15)], HDR_DEP, DS_DEP, DE_DEP)

# Override kolomkoppen rij 5 naar paarse stijl (add_cls_pairs zette navy)
for code_col in [10, 12, 14]:
    c = ws_dep.cell(row=HDR_DEP, column=code_col)
    c.fill = fill("purple"); c.font = Font(name="Calibri", bold=True, size=10, color=C["white"])
    c.border = border_all("purple")

# Code-cols Security Requirements → grijs (VBA vult automatisch in, niet bedoeld voor handmatige invoer)
for r in range(DS_DEP, DE_DEP + 1):
    for col_idx in [10, 12, 14]:
        c = ws_dep.cell(row=r, column=col_idx)
        c.fill = fill("grey_light")
        c.font = Font(name="Calibri", size=9, bold=True, italic=True, color=C["subtext"])
        c.alignment = align("center", "center")

# Security Objectives: 3×code+label paren (editeerbaar, navy — zelfde layout als requirements)
add_cls_pairs(ws_dep, [("dep_col_conf",  16, 17)], HDR_DEP, DS_DEP, DE_DEP)
add_cls_pairs(ws_dep, [("dep_col_integ", 18, 19)], HDR_DEP, DS_DEP, DE_DEP)
add_cls_pairs(ws_dep, [("dep_col_avail", 20, 21)], HDR_DEP, DS_DEP, DE_DEP)

# Gap Analyse-kolommen (22-24) — niet editeerbaar, VBA-berekend
for col_idx, key in [(22, "dep_col_gap_conf"), (23, "dep_col_gap_integ"), (24, "dep_col_gap_avail")]:
    c = ws_dep.cell(row=HDR_DEP, column=col_idx, value=t(key))
    c.fill = fill("red"); c.font = Font(name="Calibri", bold=True, size=10, color=C["white"])
    c.alignment = align("center", "center", wrap=True); c.border = border_all("red")
for r in range(DS_DEP, DE_DEP + 1):
    for col_idx in [22, 23, 24]:
        c = ws_dep.cell(row=r, column=col_idx)
        c.fill = fill("grey_light")
        c.font = Font(name="Calibri", size=9, italic=True, color=C["subtext"])
        c.alignment = align("left", "top", wrap=True)
        c.border = border_all()

# CF: gap-cel niet leeg → rood (er is een tekort)
for gap_col in [22, 23, 24]:
    cl = get_column_letter(gap_col)
    rng = f"{cl}{DS_DEP}:{cl}{DE_DEP}"
    ws_dep.conditional_formatting.add(rng, FormulaRule(
        formula=[f'{cl}{DS_DEP}<>""'],
        fill=PatternFill("solid", fgColor=C["red_light"]),
        font=Font(name="Calibri", bold=True, size=9, color=C["red"]),
    ))

# Commentaar-kolom (25) — vrij tekstveld
c = ws_dep.cell(row=HDR_DEP, column=25, value=t("dep_col_obj_commentaar"))
c.fill = fill("navy"); c.font = Font(name="Calibri", bold=True, size=10, color=C["white"])
c.alignment = align("center", "center", wrap=True); c.border = border_all("navy")
for r in range(DS_DEP, DE_DEP + 1):
    c = ws_dep.cell(row=r, column=25)
    c.alignment = align("left", "top", wrap=True)

# Celbeveiliging: standaard alles vergrendeld; ontgrendel alleen editeerbare kolommen
# Editeerbaar: B-G (2-7) = basisgegevens incl. Overarching + Opmerkingen, P-U (16-21) = objectives, Y (25) = commentaar
# Gap-kolommen V-X (22-24) zijn vergrendeld (VBA-berekend)
EDITABLE_COLS = list(range(2, 8)) + list(range(16, 22)) + [25]
for r in range(DS_DEP, DE_DEP + 1):
    for col_idx in EDITABLE_COLS:
        ws_dep.cell(row=r, column=col_idx).protection = Protection(locked=False)

ws_dep.protection.sheet = True
ws_dep.protection.autoFilter = False   # filteren toegestaan
ws_dep.protection.sort = False         # sorteren toegestaan

ws_dep.freeze_panes = f"A{DS_DEP}"
ws_dep.auto_filter.ref = f"A{HDR_DEP}:{get_column_letter(N_DEP)}{HDR_DEP}"


def build_rarm(ws):
    """
    RARM — Risk Assessment & Remediation Matrix.
    Kolommen: A=Control ID  B=Richtlijn  C=Assurance  D=Sleutelmaatregel  E+=per Dependent Asset (dynamisch via macro)
    Rijen:    1=Titel  2=DA-namen  3=Kwetsbaarheden-selector  4+=CyFun 2025 controls
    """
    ROW_DA       = 2
    ROW_VULN     = 3
    ROW_DATA     = 4
    COL_ID       = 1
    COL_TITLE    = 2
    COL_ASS      = 3
    COL_KM       = 4   # Sleutelmaatregel (key measure)
    COL_DA_START = 5
    N_TEMPLATE   = 10

    CYFUN_TABS = ["GOVERN", "IDENTIFY", "PROTECT", "DETECT", "RESPOND", "RECOVER"]
    ASS_STYLE  = {"Basic": ("green_light","green"), "Important": ("yellow_light","yellow"),
                  "Essential": ("red_light","red")}
    LAST_COL = get_column_letter(COL_DA_START + N_TEMPLATE - 1)

    no_gridlines(ws)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 13
    ws.column_dimensions["D"].width = 18   # Sleutelmaatregel
    for k in range(N_TEMPLATE):
        ws.column_dimensions[get_column_letter(COL_DA_START + k)].width = 22

    controls = []
    if CYFUN_SRC.exists():
        src_wb = load_workbook(str(CYFUN_SRC), read_only=True, data_only=True)
        for tab in CYFUN_TABS:
            if tab not in src_wb.sheetnames:
                continue
            for row in src_wb[tab].iter_rows(min_row=3, values_only=True):
                if row[5] is None:
                    continue
                assurance = str(row[4]).strip() if row[4] is not None else ""
                req_text  = str(row[5]).strip().replace("\n", " ").replace("\t", "")
                parts     = req_text.split(":", 1)
                req_id    = parts[0].strip()
                req_title = parts[1].strip() if len(parts) > 1 else req_text
                is_km     = row[2] is not None and str(row[2]).strip() != ""
                controls.append((req_id, req_title, assurance, is_km))
        src_wb.close()

    ws.merge_cells(f"A1:{LAST_COL}1")
    c = ws["A1"]
    c.value = "RARM — Risk Assessment & Remediation Matrix"
    c.fill = fill("navy"); c.font = font(13, bold=True, color="white")
    c.alignment = align("left", "center", indent=1)
    ws.row_dimensions[1].height = 34

    ws.row_dimensions[2].height = 36
    for col_idx, label in [(COL_ID, "Control ID"), (COL_TITLE, "Richtlijn"),
                           (COL_ASS, "Assurance"), (COL_KM, "Sleutelmaatregel")]:
        c = ws.cell(row=ROW_DA, column=col_idx, value=label)
        c.fill = fill("navy"); c.font = font(10, bold=True, color="white")
        c.alignment = align("center", "center", wrap=True); c.border = border_all("navy")
    for k in range(N_TEMPLATE):
        c = ws.cell(row=ROW_DA, column=COL_DA_START + k)
        c.fill = fill("yellow_light"); c.font = font(10, bold=True)
        c.alignment = align("center", "center", wrap=True); c.border = border_all("navy")

    ws.row_dimensions[3].height = 28
    for col_idx in [COL_ID, COL_TITLE, COL_ASS, COL_KM]:
        c = ws.cell(row=ROW_VULN, column=col_idx)
        c.fill = fill("grey_light"); c.border = border_all("grey_border")
    for k in range(N_TEMPLATE):
        c = ws.cell(row=ROW_VULN, column=COL_DA_START + k)
        c.fill = fill("blue_xlight"); c.font = font(9, italic=True, color="subtext")
        c.alignment = align("center", "center", wrap=True); c.border = border_all("navy")

    thick = Side(border_style="medium", color=C["navy"])
    for row_idx in range(1, ROW_DATA + len(controls) + 1):
        cell = ws.cell(row=row_idx, column=COL_DA_START)
        b = cell.border
        cell.border = Border(left=thick, right=b.right, top=b.top, bottom=b.bottom)

    ws.freeze_panes = "E4"

    for i, (req_id, req_title, assurance, is_km) in enumerate(controls):
        row_idx = ROW_DATA + i
        bg_key, fg_key = ASS_STYLE.get(assurance, ("white", "text"))
        row_bg = "grey_light" if i % 2 == 0 else "white"

        c = ws.cell(row=row_idx, column=COL_ID, value=req_id)
        c.fill = fill(bg_key); c.font = font(9, bold=True, color=fg_key)
        c.alignment = align("center", "center"); c.border = border_all("grey_border")

        c = ws.cell(row=row_idx, column=COL_TITLE, value=req_title)
        c.fill = fill(row_bg); c.font = font(9)
        c.alignment = align("left", "top", wrap=True); c.border = border_all("grey_border")
        ws.row_dimensions[row_idx].height = 30

        c = ws.cell(row=row_idx, column=COL_ASS, value=assurance)
        c.fill = fill(bg_key); c.font = font(9, bold=True, color=fg_key)
        c.alignment = align("center", "center"); c.border = border_all("grey_border")

        # Sleutelmaatregel — ✔ indien key measure in CyFun bron, anders leeg; togglebaar via dubbelklik
        c = ws.cell(row=row_idx, column=COL_KM)
        c.fill = fill(row_bg)
        c.value = "✔" if is_km else ""
        c.font = font(10, bold=True, color="green")
        c.alignment = align("center", "center"); c.border = border_all("grey_border")

        for k in range(N_TEMPLATE):
            c = ws.cell(row=row_idx, column=COL_DA_START + k)
            c.fill = fill(row_bg); c.border = border_all("grey_border")
            c.alignment = align("center", "center")


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: RARM
# ══════════════════════════════════════════════════════════════════════════════
ws_rarm = wb.create_sheet("RARM")
build_rarm(ws_rarm)

# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Verantwoordelijken
# ══════════════════════════════════════════════════════════════════════════════
ws_ver = wb.create_sheet("Responsible Persons")
no_gridlines(ws_ver)
VER_COLS = [("verant_col_id",6),("verant_col_naam",20),("verant_col_voornaam",18),
            ("verant_col_functie",24),("verant_col_dienst",22),("verant_col_email",28),
            ("verant_col_tel",16),("verant_col_rol",22),("verant_col_opmerkingen",28)]
N_VER = len(VER_COLS)
set_col_widths(ws_ver, [w for _, w in VER_COLS])
sheet_title_bar(ws_ver, t("verant_title"), N_VER)
ws_ver.row_dimensions[3].height = 22; ws_ver.merge_cells(f"A3:{get_column_letter(N_VER)}3")
c = ws_ver["A3"]; c.value = t("verant_subtitle")
c.fill = fill("blue_xlight"); c.font = font(10, italic=True, color="grey_dark")
c.alignment = align("center", "center")
ws_ver.row_dimensions[4].height = 8

VER_HDR = 5; ws_ver.row_dimensions[VER_HDR].height = 44
for col_idx, (key, _) in enumerate(VER_COLS, 1):
    c = ws_ver.cell(row=VER_HDR, column=col_idx, value=t(key))
    c.fill = fill("navy"); c.font = Font(name="Calibri", bold=True, size=10, color=C["white"])
    c.alignment = align("center", "center", wrap=True); c.border = border_all("navy")

VER_START, VER_END = 6, 55
dv_rol = DataValidation(type="list", formula1=f'=INDIRECT("roles_"&{LANG_CELL})',
                        showDropDown=False, allow_blank=True)
dv_rol.sqref = f"H{VER_START}:H{VER_END}"; ws_ver.add_data_validation(dv_rol)

for r in range(VER_START, VER_END + 1):
    ws_ver.row_dimensions[r].height = 24
    alt = "blue_xlight" if r % 2 == 0 else "white"
    for col_idx in range(1, N_VER + 1):
        c = ws_ver.cell(row=r, column=col_idx)
        c.fill = fill(alt); c.font = font(10, color="text")
        c.alignment = align("left", "center", wrap=True); c.border = border_all()
    id_c = ws_ver.cell(row=r, column=1)
    id_c.value = f'=IF(B{r}="","",TEXT({r - VER_START + 1},"V00"))'
    id_c.font = Font(name="Calibri", size=9, bold=True, color=C["subtext"])
    id_c.alignment = align("center", "center")

ws_ver.freeze_panes = f"A{VER_START}"
ws_ver.auto_filter.ref = f"A{VER_HDR}:{get_column_letter(N_VER)}{VER_HDR}"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Import & Export
# ══════════════════════════════════════════════════════════════════════════════
ws_ie = wb.create_sheet("Import & Export")
no_gridlines(ws_ie); set_col_widths(ws_ie, [2, 26, 2, 56, 8])
sheet_title_bar(ws_ie, t("ie_title"), 5)
ws_ie.row_dimensions[3].height = 22; ws_ie.merge_cells("A3:E3")
c = ws_ie["A3"]; c.value = t("ie_subtitle")
c.fill = fill("blue_xlight"); c.font = font(10, italic=True, color="grey_dark")
c.alignment = align("center", "center")
ws_ie.row_dimensions[4].height = 10

def ie_action_block(ws, row, btn_key, desc_key, btn_color="blue_mid"):
    """Maakt een import/export-actieblok met knoopplaats en beschrijving."""
    ws.row_dimensions[row].height = 34
    btn_cell = ws.cell(row=row, column=2)
    btn_cell.value = t(btn_key)
    btn_cell.fill = fill(btn_color); btn_cell.font = Font(name="Calibri", bold=True, size=11, color=C["white"])
    btn_cell.alignment = align("center", "center"); btn_cell.border = border_all(btn_color)

    ws.row_dimensions[row + 1].height = 44
    ws.merge_cells(f"B{row+1}:D{row+1}")
    desc = ws.cell(row=row + 1, column=2, value=t(desc_key))
    desc.fill = fill("grey_light"); desc.font = font(10, color="subtext", italic=True)
    desc.alignment = align("left", "top", wrap=True, indent=1); desc.border = border_all()
    ws.row_dimensions[row + 2].height = 10
    return row + 3

section_header(ws_ie, 5, "ie_section_import", 4, col_start=2)
cur = 6
cur = ie_action_block(ws_ie, cur, "ie_import_all_btn",   "ie_import_all_desc")
cur = ie_action_block(ws_ie, cur, "ie_import_ia_btn",    "ie_import_ia_desc")
cur = ie_action_block(ws_ie, cur, "ie_import_da_btn",    "ie_import_da_desc")
cur = ie_action_block(ws_ie, cur, "ie_import_proc_btn",  "ie_import_proc_desc")
cur = ie_action_block(ws_ie, cur, "ie_import_kwets_btn", "ie_import_kwets_desc")
cur = ie_action_block(ws_ie, cur, "ie_import_links_btn", "ie_import_links_desc", btn_color="orange")

section_header(ws_ie, cur, "ie_section_export", 4, col_start=2)
cur += 1

section_header(ws_ie, cur, "ie_section_info", 4, col_start=2)
cur += 1

# Tabelstructuur bronbestand tonen als referentie
def ie_table_info(ws, row, tbl_key, cols_str):
    ws.row_dimensions[row].height = 22; ws.merge_cells(f"B{row}:D{row}")
    c = ws.cell(row=row, column=2, value=t(tbl_key))
    c.fill = fill("navy"); c.font = Font(name="Calibri", bold=True, size=10, color=C["white"])
    c.alignment = align("left", "center", indent=1)
    ws.row_dimensions[row+1].height = 18; ws.merge_cells(f"B{row+1}:D{row+1}")
    c2 = ws.cell(row=row+1, column=2, value=t("ie_col_hint"))
    c2.fill = fill("grey_light"); c2.font = font(9, italic=True, color="subtext")
    c2.alignment = align("left", "center", indent=1)
    ws.row_dimensions[row+2].height = 52; ws.merge_cells(f"B{row+2}:D{row+2}")
    c3 = ws.cell(row=row+2, column=2, value=cols_str)
    c3.fill = fill("grey_light"); c3.font = Font(name="Courier New", size=9, color=C["text"])
    c3.alignment = align("left", "top", wrap=True, indent=1); c3.border = border_all()
    return row + 4

proc_cols = ("naam / name / nom\nomschrijving / description\neigenaar / owner / propriétaire\n"
             "dienst / department / service\nintegriteit / integrity / intégrité\n"
             "beschikbaarheid / availability / disponibilité\nconfidentialiteit / confidentiality / confidentialité")
asset_cols = ("naam / name / nom\ntype / asset type\nomschrijving / description\n"
              "eigenaar / owner / propriétaire\ndienst / department / service\n"
              "integriteit / integrity / intégrité\nbeschikbaarheid / availability / disponibilité\n"
              "confidentialiteit / confidentiality / confidentialité")

cur = ie_table_info(ws_ie, cur, "ie_table_proc",  proc_cols)
cur = ie_table_info(ws_ie, cur, "ie_table_asset", asset_cols)

ws_ie.row_dimensions[cur].height = 22; ws_ie.merge_cells(f"B{cur}:D{cur}")
c = ws_ie.cell(row=cur, column=2, value=t("ie_macro_hint"))
c.fill = fill("accent_light"); c.font = Font(name="Calibri", size=9, color=C["grey_dark"], italic=True)
c.alignment = align("left", "center", indent=1)


# ══════════════════════════════════════════════════════════════════════════════
# SHEET: Referentiewaarden  (naar achter geschoven)
# ══════════════════════════════════════════════════════════════════════════════
ws_ref = wb.create_sheet("Reference Values")
no_gridlines(ws_ref); set_col_widths(ws_ref, [2, 18, 52, 8, 8, 8]); N_REF = 6
sheet_title_bar(ws_ref, t("ref_title"), N_REF)
ws_ref.merge_cells("A3:F3"); ws_ref.row_dimensions[3].height = 30
c = ws_ref["A3"]; c.value = t("ref_subtitle")
c.fill = fill("blue_xlight"); c.font = font(10, italic=True, color="grey_dark")
c.alignment = align("center", "center", wrap=True)
ws_ref.row_dimensions[4].height = 10

def ref_dim_table(ws, start_row, section_key, subtitle_key, desc_keys, n_cols=6):
    section_header(ws, start_row, section_key, n_cols)
    row = start_row + 1
    ws.row_dimensions[row].height = 22; ws.merge_cells(f"A{row}:{get_column_letter(n_cols)}{row}")
    c = ws.cell(row=row, column=1, value=t(subtitle_key))
    c.fill = fill("grey_light"); c.font = font(10, italic=True, color="subtext")
    c.alignment = align("left", "center", indent=1)
    row += 1
    ws.row_dimensions[row].height = 26
    for col, key in enumerate(["ref_level", "ref_label", "ref_description"], 1):
        c = ws.cell(row=row, column=col, value=t(key))
        c.fill = fill("blue_mid"); c.font = font(10, bold=True, color="white")
        c.alignment = align("center", "center", wrap=True); c.border = border_all("navy")
    lvl_keys = ["ref_lvl1_label","ref_lvl2_label","ref_lvl3_label","ref_lvl4_label","ref_lvl5_label"]
    for i, (lvl_key, desc_key, fill_k, font_k) in enumerate(
            zip(lvl_keys, desc_keys, LEVEL_FILLS, LEVEL_FONTS)):
        row += 1; ws.row_dimensions[row].height = 52
        # Nummer
        c = ws.cell(row=row, column=1, value=i + 1)
        c.fill = fill(fill_k); c.font = Font(name="Calibri", bold=True, size=14, color=C[font_k])
        c.alignment = align("center", "center"); c.border = border_all()
        # Benaming
        c = ws.cell(row=row, column=2, value=t(lvl_key))
        c.fill = fill(fill_k); c.font = Font(name="Calibri", bold=True, size=10, color=C[font_k])
        c.alignment = align("center", "center", wrap=True); c.border = border_all()
        # Beschrijving
        ws.merge_cells(f"C{row}:F{row}")
        c = ws.cell(row=row, column=3, value=t(desc_key))
        c.fill = fill(fill_k); c.font = font(10, color=font_k)
        c.alignment = align("left", "center", wrap=True, indent=1); c.border = border_all()
    return row + 2

cur_ref = 5
cur_ref = ref_dim_table(ws_ref, cur_ref, "ref_section_int",   "ref_int_subtitle",
                         ["ref_int_1","ref_int_2","ref_int_3","ref_int_4","ref_int_5"])
cur_ref = ref_dim_table(ws_ref, cur_ref, "ref_section_avail", "ref_avail_subtitle",
                         ["ref_avail_1","ref_avail_2","ref_avail_3","ref_avail_4","ref_avail_5"])
cur_ref = ref_dim_table(ws_ref, cur_ref, "ref_section_conf",  "ref_conf_subtitle",
                         ["ref_conf_1","ref_conf_2","ref_conf_3","ref_conf_4","ref_conf_5"])


# ── Kwetsbaarheden sheet ─────────────────────────────────────────────────────
def _load_kwets_db():
    """Laadt kwetsbaarheden + control-links uit Access DB.
    Retourneert (vulns, ctrl_links_2025, unmatched_23ids):
      vulns            = [(id, name, C_bool, I_bool, A_bool), ...]
      ctrl_links_2025  = {vuln_id: set_of_2025_req_ids}   — via 2023→2025 mapping
      unmatched_23ids  = {norm_2023_id}   — 2023 IDs zonder 2025 equivalent
    """
    DB_PATH = Path(__file__).parent.parent / "data" / "Import" / "MNMTool - SocSec.accdb"
    if not DB_PATH.exists():
        print(f"[kwetsbaarheden] Access DB niet gevonden: {DB_PATH}")
        return [], {}, set()
    try:
        import pyodbc
        conn = pyodbc.connect(
            f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={DB_PATH};"
        )
        cur = conn.cursor()
        cur.execute(
            "SELECT ID, Reference, Vulnerability, Confidentialiteit, Integriteit, Beschikbaarheid "
            "FROM [T - Vulnerabilities] ORDER BY Reference"
        )
        vulns = [(int(r[0]), str(r[2]), bool(r[3]), bool(r[4]), bool(r[5]))
                 for r in cur.fetchall()]
        cur.execute("SELECT RefNr, Requirement FROM [T - CyFunEssentiel]")
        refnr_map = {int(r[0]): str(r[1]) for r in cur.fetchall()}
        cur.execute(
            "SELECT Vulnerability, CyFunControl "
            "FROM [LT -  Vulnerability to control - fixed]"
        )
        raw_links = [(r[0], r[1]) for r in cur.fetchall()
                     if r[0] is not None and r[1] is not None]
        conn.close()
    except Exception as e:
        print(f"[kwetsbaarheden] DB-fout: {e}")
        return [], {}, set()

    # ── Bouw 2023→2025 reverse mapping via MAPPING_SRC ────────────────────────
    # _load_all_mappings() geeft {norm_2025_id → "2023_id"} (both_map)
    # We draaien dit om: {norm_2023_id → 2025_id}
    rev_map = {}   # norm_2023_id → canonical 2025_id (string uit CYFUN_SRC)
    try:
        both_map, _, _ = _load_all_mappings()
        for norm25, id23 in both_map.items():
            n23 = re.sub(r'-(\d+)$', r'.\1', str(id23).strip().lower())
            # norm25 is already normalized; de-normalize to get original 2025 ID
            # We need the original (non-normalized) 2025 ID to match sheet headers.
            # Store norm_2023 → norm_2025 (sheet headers are compared normalized)
            rev_map[n23] = norm25
    except Exception as e:
        print(f"[kwetsbaarheden] mapping-fout: {e}")

    # ── Verwerk LT-links: 2023 ID → 2025 ID via rev_map ──────────────────────
    ctrl_links_2025 = {}
    unmatched_23ids = set()
    for vuln_id_raw, ctrl_ref_raw in raw_links:
        vid = int(vuln_id_raw)
        cid = int(ctrl_ref_raw)
        id23 = refnr_map.get(cid)
        if not id23:
            continue
        n23 = re.sub(r'-(\d+)$', r'.\1', str(id23).strip().lower())
        if n23 in rev_map:
            # Sla normalized 2025 ID op voor matching met sheet-headers
            ctrl_links_2025.setdefault(vid, set()).add(rev_map[n23])
        else:
            unmatched_23ids.add(id23)
            # Controleer ook of het ID al in 2025-formaat is (bijv. RefNr 219-227)
            # door het direct op te slaan als norm
            ctrl_links_2025.setdefault(vid, set()).add(n23)

    return vulns, ctrl_links_2025, unmatched_23ids


def build_kwetsbaarheden(ws):
    """
    Statische structuur — data wordt ENKEL via macro geladen (ImportKwetsbaarheden).

    Kolom A = Control ID, kolom B = Richtlijn  (statisch vanuit CYFUN_SRC)
    Kolom C+ = per kwetsbaarheid 1 kolom       (dynamisch, door macro)

    Rij 1  : Titel
    Rij 2  : "Control ID" | "Richtlijn" | [kwetsbaarheid-namen]
    Rij 3  : "C" | "Vertrouwelijkheid"  | [✔ als kwetsbaarheid C-impact heeft]
    Rij 4  : "I" | "Integriteit"        | [✔ als kwetsbaarheid I-impact heeft]
    Rij 5  : "A" | "Beschikbaarheid"    | [✔ als kwetsbaarheid A-impact heeft]
    Rij 6+ : ctrl_id | richtlijn        | [✔ als control remediëring biedt]
    """
    ROW_TITLE = 1
    ROW_VULN  = 2
    ROW_C     = 3
    ROW_I     = 4
    ROW_A     = 5
    ROW_DATA  = 6
    COL_ID    = 1   # A
    COL_TITLE = 2   # B
    CYFUN_TABS = ["GOVERN", "IDENTIFY", "PROTECT", "DETECT", "RESPOND", "RECOVER"]
    ASSURANCE_STYLE = {
        "Basic":     ("green_light",  "green"),
        "Important": ("yellow_light", "yellow"),
        "Essential": ("red_light",    "red"),
    }

    no_gridlines(ws)

    # ── Controls laden uit CYFUN_SRC (statische referentiedata) ──────────────
    controls = []   # [(req_id, req_title, assurance), ...]
    if CYFUN_SRC.exists():
        src_wb = load_workbook(str(CYFUN_SRC), read_only=True, data_only=True)
        for tab in CYFUN_TABS:
            if tab not in src_wb.sheetnames:
                continue
            for row in src_wb[tab].iter_rows(min_row=3, values_only=True):
                if row[5] is None:
                    continue
                assurance = str(row[4]).strip() if row[4] is not None else ""
                req_text  = str(row[5]).strip().replace("\n", " ").replace("\t", "")
                parts     = req_text.split(":", 1)
                req_id    = parts[0].strip()
                req_title = parts[1].strip() if len(parts) > 1 else req_text
                controls.append((req_id, req_title, assurance))
        src_wb.close()

    N_TEMPLATE    = 5   # lege invoerkolommen voor manuele input
    COL_INP_START = 3   # C
    LAST_HDR_COL  = get_column_letter(COL_INP_START + N_TEMPLATE - 1)

    # ── Kolombreedtes ─────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 65
    for k in range(N_TEMPLATE):
        ws.column_dimensions[get_column_letter(COL_INP_START + k)].width = 14

    # ── Rij 1: titel (gemerged over A t/m laatste invoerkolom) ───────────────
    ws.merge_cells(f"A1:{LAST_HDR_COL}1")
    c = ws["A1"]
    c.value = "Kwetsbaarheden — Remediëring via CyFun 2025 Controls  ·  dubbelklik op een cel om ✔ te plaatsen"
    c.fill = fill("navy"); c.font = font(11, bold=True, color="white")
    c.alignment = align("left", "center", indent=1)
    ws.row_dimensions[ROW_TITLE].height = 34

    # ── Rij 2: "Control ID" / "Richtlijn" + lege invoerkoppen ────────────────
    ws.row_dimensions[ROW_VULN].height = 55
    for col, lbl in [(COL_ID, "Control ID"), (COL_TITLE, "Richtlijn")]:
        c = ws.cell(row=ROW_VULN, column=col, value=lbl)
        c.fill = fill("navy"); c.font = font(10, bold=True, color="white")
        c.alignment = align("center", "center", wrap=True)
        c.border = border_all("navy")

    # ── Rijen 3-5: C/I/A rijlabels ───────────────────────────────────────────
    CIA_DEFS = [
        (ROW_C, "C", "Vertrouwelijkheid", "blue_light",   "blue_mid"),
        (ROW_I, "I", "Integriteit",       "green_light",  "green"),
        (ROW_A, "A", "Beschikbaarheid",   "orange_light", "orange"),
    ]
    for row, short, full, bg, fg in CIA_DEFS:
        ws.row_dimensions[row].height = 18
        c = ws.cell(row=row, column=COL_ID, value=short)
        c.fill = fill(bg); c.font = Font(name="Calibri", size=9, bold=True, color=C[fg])
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_all("grey_border")
        c = ws.cell(row=row, column=COL_TITLE, value=full)
        c.fill = fill("white"); c.font = Font(name="Calibri", size=9, color=C["text"])
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = border_all("grey_border")

    # Bevroren venster: rijen 1-5 EN kolommen A-B altijd zichtbaar
    ws.freeze_panes = "C6"

    if not controls:
        ws.cell(row=ROW_DATA, column=COL_ID,
                value="CYFUN_SRC niet gevonden — controleer het pad")
        return []

    # ── Rijen 6+: controls (A = ID, B = richtlijn) — ✔ marks door macro ──────
    _bdr = border_all("grey_border")
    for j, (ctrl_id, ctrl_title, assurance) in enumerate(controls):
        row    = ROW_DATA + j
        row_bg = "grey_light" if j % 2 == 0 else "white"
        bg_key, fg_key = ASSURANCE_STYLE.get(assurance, ("white", "text"))

        c = ws.cell(row=row, column=COL_ID, value=ctrl_id)
        c.fill = fill(bg_key)
        c.font = Font(name="Calibri", size=9, bold=True, color=C[fg_key])
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = _bdr

        c = ws.cell(row=row, column=COL_TITLE, value=ctrl_title)
        c.fill = fill(row_bg)
        c.font = Font(name="Calibri", size=9, color=C["text"])
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c.border = _bdr

        ws.row_dimensions[row].height = 30

    # ── 5 lege invoerkolommen voor manuele input (C t/m G) ───────────────────
    # Dikke linker scheidingslijn op col C, dunne randen op D-G
    _thin   = Side(style="thin",   color=C["grey_border"])
    _medium = Side(style="medium", color=C["navy"])
    _align_inp = Alignment(horizontal="center", vertical="center", wrap_text=True)

    CIA_COLORS = [("blue_light", ROW_C), ("green_light", ROW_I), ("orange_light", ROW_A)]

    for k in range(N_TEMPLATE):
        col  = COL_INP_START + k
        left = _medium if k == 0 else _thin

        def _inp_bdr(left_side):
            return Border(left=left_side, right=_thin, top=_thin, bottom=_thin)

        # Rij 2: kwetsbaarheid-naam invoerveld (geel = actief invoerveld)
        c = ws.cell(row=ROW_VULN, column=col)
        c.fill = fill("yellow_light")
        c.alignment = _align_inp
        c.border = _inp_bdr(left)

        # Rijen 3-5: CIA-impact invoervelden (gekleurde achtergrond)
        for bg, row in CIA_COLORS:
            c = ws.cell(row=row, column=col)
            c.fill = fill(bg)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = _inp_bdr(left)

        # Rijen 6+: control-koppeling invoervelden (wisselende achtergrond)
        for j in range(len(controls)):
            r      = ROW_DATA + j
            row_bg = "grey_light" if j % 2 == 0 else "white"
            c = ws.cell(row=r, column=col)
            c.fill = fill(row_bg)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = _inp_bdr(left)

    ws.auto_filter.ref = "A2:B2"
    return []


def build_afwijkingen(ws, afw_rows):
    """
    Rapportagesheet: controls die in de Access DB als remediëring zijn gelinkt
    maar NIET bestaan in de CyFun 2025 ESSENTIAL-set (Enkel-2023 controls).
    """
    ROW_TITLE = 1
    ROW_HDR   = 2
    ROW_DATA  = 3

    no_gridlines(ws)

    ws.column_dimensions["A"].width = 48
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 40

    # Titel
    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value = "Afwijkingen — CyFun 2023 controls zonder 2025 equivalent"
    c.fill = fill("orange"); c.font = font(13, bold=True, color="white")
    c.alignment = align("left", "center", indent=1)
    ws.row_dimensions[ROW_TITLE].height = 34

    # Kolomkoppen
    headers = ["Kwetsbaarheid", "C / I / A", "Control ID (DB)", "Reden"]
    for col, hdr in enumerate(headers, 1):
        c = ws.cell(row=ROW_HDR, column=col, value=hdr)
        c.fill = fill("navy"); c.font = font(10, bold=True, color="white")
        c.alignment = align("center", "center", wrap=True)
        c.border = border_all("navy")
    ws.row_dimensions[ROW_HDR].height = 30

    if not afw_rows:
        c = ws.cell(row=ROW_DATA, column=1,
                    value="Geen afwijkingen — alle DB-controls bestaan in CyFun 2025.")
        c.font = font(9, color="subtext"); c.alignment = align("left", "center")
        return

    for idx, (vname, cia_str, ctrl_id) in enumerate(afw_rows):
        r = ROW_DATA + idx
        row_bg = "grey_light" if idx % 2 == 0 else "white"
        vals = [vname, cia_str, ctrl_id,
                "Control bestaat enkel in CyFun 2023 — geen 2025 equivalent"]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=r, column=col, value=val)
            c.fill = fill(row_bg)
            c.font = font(9, color="text")
            c.alignment = align("left", "center", wrap=(col == 1))
            c.border = border_all("grey_border")
        ws.row_dimensions[r].height = 18

    ws.auto_filter.ref = "A2:D2"

# ── CyFun Controls sheet ─────────────────────────────────────────────────────
# Kolomindeling:  A–F = 2025 data  |  G = Versie  |  H–M = 2023 data
COL_25_START = 1   # A
COL_VERSIE   = 7   # G
COL_23_START = 8   # H

def _norm_id(s):
    """Lowercase + trailing -N → .N  (bv. 'ID.AM-03-3' → 'id.am-03.3')."""
    return re.sub(r'-(\d+)$', r'.\1', str(s).strip().lower())

def _load_all_mappings():
    """
    Doorloop de 3 sheets van het mapping-workbook.
    Geeft terug:
      both_map  : {norm_2025_id → "2023_id"}
      only25    : {norm 2025-only IDs}
      only23    : lijst van {id, ctrl_linked, key_meas} — 2023-only controls
    BASIC-sheet heeft geen DELETED/NEW-kolommen; IMPORTANT/ESSENTIAL wel.
    """
    both_map, only25, only23 = {}, set(), []
    if not MAPPING_SRC.exists():
        return both_map, only25, only23

    wb = load_workbook(str(MAPPING_SRC), read_only=True, data_only=True)

    for sheet in ["CyFun 2023>25 BASIC", "CyFun 2023>25 IMPORTANT", "CyFun 2023>25 ESSENTIAL"]:
        ws = wb[sheet]
        is_basic = sheet.endswith("BASIC")

        for row in ws.iter_rows(min_row=2, values_only=True):
            if is_basic:
                # BASIC: [2023_ID, req23, km23, score, 2025_ID, req25, km25, score]
                id23, km23  = row[0], row[2]
                id25        = row[4]
                deleted     = False
                new_flag    = False
            else:
                # IMPORTANT/ESSENTIAL: [2023_ID, req23, km23, DELETED, score, 2025_ID, req25, km25, NEW, score]
                id23, km23  = row[0], row[2]
                deleted     = row[3] == "DELETED"
                id25        = row[5]
                new_flag    = row[8] == "New"

            # Lege rijen overslaan
            if id23 is None and id25 is None:
                continue

            km23_s = str(km23).strip() if km23 else ""

            if deleted:
                only23.append({
                    "id":          str(id23).strip(),
                    "ctrl_linked": km23_s if "Control" in km23_s else "",
                    "key_meas":    "Key Measure" if "Key" in km23_s else "",
                })
            elif new_flag:
                if id25:
                    only25.add(_norm_id(id25))
            else:
                if id23 and id25:
                    both_map[_norm_id(id25)] = str(id23).strip()

    wb.close()
    return both_map, only25, only23

def _load_2023_details(target_ids):
    """
    Zoek target_ids op in alle 3 Details-sheets van het 2023-workbook.
    Requirement-prefix: 'BASIC_' → Basic, 'IMPORTANT_' → Important, geen → Essential.
    Geeft {req_id → {category, subcategory, assurance, requirement, key_meas}}.
    """
    found = {}
    if not CYFUN23_SRC.exists():
        return found

    wb = load_workbook(str(CYFUN23_SRC), read_only=True, data_only=True)

    # ESSENTIAL Details eerst: heeft expliciete BASIC_/IMPORTANT_ prefixes voor alle niveaus.
    # Daarna IMPORTANT Details en BASIC Details als fallback voor controls die enkel daar staan.
    # Als een control geen prefix heeft: niveau = wat de sheet aangeeft (Essential/Important/Basic).
    sheet_default = {
        "ESSENTIAL Details": "Essential",
        "IMPORTANT Details": "Important",
        "BASIC Details":     "Basic",
    }
    for sheet in ["ESSENTIAL Details", "IMPORTANT Details", "BASIC Details"]:
        ws = wb[sheet]
        cur_cat, cur_sub = "", ""
        for row in ws.iter_rows(min_row=3, values_only=True):
            if row[1]: cur_cat = str(row[1]).strip()
            if row[3]: cur_sub = str(row[3]).strip()
            if not row[4]:
                continue
            req_str = str(row[4]).strip()
            clean   = re.sub(r'^(BASIC_|IMPORTANT_)', '', req_str)
            parts   = clean.split(":", 1)
            req_id  = parts[0].strip()
            if req_id not in target_ids or req_id in found:
                continue
            if req_str.startswith("BASIC_"):
                level = "Basic"
            elif req_str.startswith("IMPORTANT_"):
                level = "Important"
            else:
                level = sheet_default[sheet]   # geen prefix → niveau van de sheet
            req_text = f"{req_id}: {parts[1].strip()}" if len(parts) > 1 else req_str
            km_raw   = str(row[2]).strip() if row[2] else ""
            key_m    = "Key Measure" if "Key" in km_raw else ""
            ctrl_l   = km_raw if "Control" in km_raw else ""
            found[req_id] = {
                "category":    cur_cat,
                "subcategory": cur_sub,
                "assurance":   level,
                "requirement": req_text.replace("\n", " ").replace("\t", ""),
                "key_meas":    key_m,
                "ctrl_linked": ctrl_l,
            }

    wb.close()
    return found

STATUS_STYLE = {
    "Beide":      ("green_light",  "green"),
    "Enkel 2025": ("blue_xlight",  "blue_mid"),
    "Enkel 2023": ("orange_light", "orange"),
}

def _write_side(ws, data_row, col_start, vals6, assurance, row_bg, assurance_style):
    """
    Schrijf 6 datacellen (Category/Controls/KeyMeasure/Subcategory/Assurance/Requirement)
    vanaf col_start. vals6=None schrijft lege cellen met rand.
    """
    bg_key, fg_key = assurance_style.get(assurance, ("white", "text"))
    values = vals6 if vals6 is not None else [""] * 6
    for i, val in enumerate(values):
        col = col_start + i
        pos = i + 1                  # 1=Cat 2=Ctrl 3=KM 4=Sub 5=Ass 6=Req
        c = ws.cell(row=data_row, column=col, value=val if val else None)
        c.border = border_all("grey_border")
        if not val:
            c.fill = fill(row_bg)
            c.alignment = align("left", "top")
        elif pos == 5:               # Assurance level
            c.fill = fill(bg_key)
            c.font = font(9, bold=True, color=fg_key)
            c.alignment = align("center", "top")
        elif pos == 3:               # Key Measure badge
            c.fill = fill("accent_light")
            c.font = font(9, bold=True, color="accent")
            c.alignment = align("center", "top")
        else:
            c.fill = fill(row_bg)
            c.font = font(9, color="text")
            c.alignment = align("left", "top", wrap=True)

def _write_versie(ws, data_row, status):
    """Schrijf de Versie-cel (kolom G)."""
    sbg, sfg = STATUS_STYLE.get(status, ("white", "text"))
    c = ws.cell(row=data_row, column=COL_VERSIE, value=status)
    c.border = border_all("grey_border")
    c.fill = fill(sbg)
    c.font = font(9, bold=True, color=sfg)
    c.alignment = align("center", "top")

def build_cyfun_controls(ws):
    CYFUN_TABS = ["GOVERN", "IDENTIFY", "PROTECT", "DETECT", "RESPOND", "RECOVER"]
    # 6 kolomnamen per zijde
    HDR_25 = ["Category",
              "Controls linked to\nthe management aspects",
              "Key Measure",
              "Subcategory",
              "Assurance level",
              "Requirement"]
    HDR_23 = ["Category (2023)",
              "Controls linked (2023)",
              "Key Measure (2023)",
              "Subcategory (2023)",
              "Assurance level (2023)",
              "Requirement (2023)"]
    # Kolombreedtes: 6×(2025) + Versie + 6×(2023)
    COL_WIDTHS = [30, 24, 14, 36, 16, 90,   # A–F
                  13,                         # G
                  30, 24, 14, 36, 16, 90]     # H–M
    N        = len(COL_WIDTHS)               # 13
    LAST_COL = get_column_letter(N)
    ASSURANCE_STYLE = {
        "Basic":     ("green_light",  "green"),
        "Important": ("yellow_light", "yellow"),
        "Essential": ("red_light",    "red"),
    }

    # ── Mapping + 2023-details laden ─────────────────────────────────────────
    both_map, only25_set, only23_meta = _load_all_mappings()
    # Laad 2023-details voor ALLE gekoppelde IDs + only-2023
    all_2023_ids = set(both_map.values()) | {d["id"] for d in only23_meta}
    details_2023 = _load_2023_details(all_2023_ids)

    # ── Opmaak ───────────────────────────────────────────────────────────────
    no_gridlines(ws)
    set_col_widths(ws, COL_WIDTHS)

    # Rij 1: titelregel
    ws.merge_cells(f"A1:{LAST_COL}1")
    c = ws["A1"]
    c.value = "CyFun 2025 + 2023 — Controls & Requirements (ESSENTIAL) — Vergelijking"
    c.fill = fill("navy"); c.font = font(13, bold=True, color="white")
    c.alignment = align("left", "center", indent=1)
    ws.row_dimensions[1].height = 34

    # Rij 2: groepskoppen
    GRP = 2
    ws.row_dimensions[GRP].height = 22
    lc25 = get_column_letter(COL_25_START + 5)
    lc23 = get_column_letter(COL_23_START + 5)
    ws.merge_cells(f"A{GRP}:{lc25}{GRP}")
    c = ws[f"A{GRP}"]
    c.value = "CyFun 2025"; c.fill = fill("navy")
    c.font = font(10, bold=True, color="white"); c.alignment = align("center", "center")
    c = ws.cell(row=GRP, column=COL_VERSIE, value="Versie")
    c.fill = fill("grey_dark"); c.font = font(10, bold=True, color="white")
    c.alignment = align("center", "center")
    ws.merge_cells(f"H{GRP}:{lc23}{GRP}")
    c = ws[f"H{GRP}"]
    c.value = "CyFun 2023"; c.fill = fill("blue_mid")
    c.font = font(10, bold=True, color="white"); c.alignment = align("center", "center")

    # Rij 3: kolomkoppen
    HDR = 3
    ws.row_dimensions[HDR].height = 36
    for i, h in enumerate(HDR_25):
        c = ws.cell(row=HDR, column=COL_25_START + i, value=h)
        c.fill = fill("navy"); c.font = font(10, bold=True, color="white")
        c.alignment = align("center", "center", wrap=True); c.border = border_all("navy")
    c = ws.cell(row=HDR, column=COL_VERSIE, value="Versie")
    c.fill = fill("grey_dark"); c.font = font(10, bold=True, color="white")
    c.alignment = align("center", "center"); c.border = border_all("grey_dark")
    for i, h in enumerate(HDR_23):
        c = ws.cell(row=HDR, column=COL_23_START + i, value=h)
        c.fill = fill("blue_mid"); c.font = font(10, bold=True, color="white")
        c.alignment = align("center", "center", wrap=True); c.border = border_all("navy")

    ws.auto_filter.ref = f"A{HDR}:{LAST_COL}{HDR}"
    ws.freeze_panes   = f"A{HDR + 1}"

    if not CYFUN_SRC.exists():
        ws[f"A{HDR+1}"].value = f"BESTAND NIET GEVONDEN: {CYFUN_SRC}"
        return

    # ── 2025 requirements ────────────────────────────────────────────────────
    src_wb   = load_workbook(str(CYFUN_SRC), read_only=True, data_only=True)
    data_row = HDR + 1

    for tab in CYFUN_TABS:
        if tab not in src_wb.sheetnames:
            continue
        src_ws  = src_wb[tab]
        cur_cat = ""
        cur_sub = ""
        for row in src_ws.iter_rows(min_row=3, values_only=True):
            if row[0] is not None:
                cur_cat = str(row[0]).strip()
                cur_sub = ""
            if row[3] is not None:
                cur_sub = str(row[3]).strip()
            if row[5] is None:
                continue
            ctrl_linked = str(row[1]).strip() if (row[1] is not None and isinstance(row[1], str)) else ""
            key_meas    = "Key Measure" if (row[2] is not None and "Key" in str(row[2])) else ""
            assurance   = str(row[4]).strip() if row[4] is not None else ""
            req_text    = str(row[5]).strip().replace("\n", " ").replace("\t", "")
            req_id_norm = _norm_id(req_text.split(":")[0])

            if req_id_norm in both_map:
                status   = "Beide"
                id23     = both_map[req_id_norm]
                d23      = details_2023.get(id23, {})
                vals_23  = [d23.get("category", ""),
                            d23.get("ctrl_linked", ""),
                            d23.get("key_meas",    ""),
                            d23.get("subcategory", ""),
                            d23.get("assurance",   ""),
                            d23.get("requirement", "")]
                ass_23   = d23.get("assurance", "")
            else:
                status   = "Enkel 2025"
                vals_23  = None
                ass_23   = ""

            row_bg = "grey_light" if data_row % 2 == 0 else "white"
            vals_25 = [cur_cat, ctrl_linked, key_meas, cur_sub, assurance, req_text]
            _write_side(ws, data_row, COL_25_START, vals_25, assurance, row_bg, ASSURANCE_STYLE)
            _write_versie(ws, data_row, status)
            _write_side(ws, data_row, COL_23_START, vals_23, ass_23, row_bg, ASSURANCE_STYLE)
            data_row += 1

    src_wb.close()

    # ── Scheidingsregel ──────────────────────────────────────────────────────
    ws.merge_cells(f"A{data_row}:{LAST_COL}{data_row}")
    c = ws.cell(row=data_row, column=1,
                value="CyFun 2023 — Controls niet opgenomen in versie 2025  (Enkel 2023)")
    c.fill = fill("orange"); c.font = font(11, bold=True, color="white")
    c.alignment = align("left", "center", indent=1)
    ws.row_dimensions[data_row].height = 24
    data_row += 1

    # ── 2023-only requirements (A–F leeg, H–M gevuld) ────────────────────────
    for meta in only23_meta:
        rid    = meta["id"]
        d      = details_2023.get(rid, {})
        ass    = d.get("assurance",   "Essential")
        vals_23 = [d.get("category",    f"(2023) {rid}"),
                   meta["ctrl_linked"] or d.get("ctrl_linked", ""),
                   meta["key_meas"]    or d.get("key_meas",    ""),
                   d.get("subcategory", ""),
                   ass,
                   d.get("requirement", rid)]
        row_bg = "grey_light" if data_row % 2 == 0 else "white"
        _write_side(ws, data_row, COL_25_START, None,    "",  row_bg, ASSURANCE_STYLE)
        _write_versie(ws, data_row, "Enkel 2023")
        _write_side(ws, data_row, COL_23_START, vals_23, ass, row_bg, ASSURANCE_STYLE)
        data_row += 1

ws_cyfun = wb.create_sheet("CyFun Controls")
build_cyfun_controls(ws_cyfun)

ws_kwets = wb.create_sheet("Kwetsbaarheden")
afw_rows = build_kwetsbaarheden(ws_kwets)

ws_afw = wb.create_sheet("Afwijkingen")
build_afwijkingen(ws_afw, afw_rows)

# ══════════════════════════════════════════════════════════════════════════════
# OPSLAAN ALS XLSX + VBA INJECTEREN VIA WIN32COM → XLSM
# ══════════════════════════════════════════════════════════════════════════════
OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
# _Lang werd als eerste aangemaakt (wb.active) — verschuif naar einde
wb.move_sheet("_Lang", offset=len(wb.sheetnames) - 1)
wb.save(OUT_XLSX)
print(f"Tussenstap opgeslagen: {OUT_XLSX}")

# ── VBA-injectie via win32com ─────────────────────────────────────────────────
vba_ok = False
try:
    import win32com.client as win32
    import pythoncom

    pythoncom.CoInitialize()
    xl = win32.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    wb_com = xl.Workbooks.Open(str(OUT_XLSX.resolve()))

    # Voeg VBA-module toe
    mod = wb_com.VBProject.VBComponents.Add(1)   # 1 = vbext_ct_StdModule
    mod.Name = "GRC_Macros"
    # Attribute VB_Name is VBE-metadata en mag NIET via AddFromString worden doorgegeven
    code_for_vba = "\n".join(
        line for line in VBA_CODE.splitlines()
        if not line.startswith("Attribute VB_Name")
    )
    mod.CodeModule.AddFromString(code_for_vba)

    # Voeg knoppen toe op Import & Export-blad
    ws_ie_com = wb_com.Sheets("Import & Export")
    def add_button(ws, left, top, w, h, caption, macro):
        btn = ws.Shapes.AddShape(1, left, top, w, h)   # 1 = msoShapeRectangle
        btn.Fill.ForeColor.RGB = int("2563EB", 16)
        btn.Line.Visible = False
        btn.TextFrame.Characters().Text = caption
        btn.TextFrame.Characters().Font.Bold = True
        btn.TextFrame.Characters().Font.Size = 11
        btn.TextFrame.Characters().Font.Color = int("FFFFFF", 16)
        btn.TextFrame.HorizontalAlignment = -4108   # xlHAlignCenter
        btn.TextFrame.VerticalAlignment = -4108     # xlVAlignCenter
        btn.OnAction = macro

    add_button(ws_ie_com, 10, 130, 200, 44, "▶  Importeer Alles",                    "GRC_Macros.ImportAlles")
    add_button(ws_ie_com, 10, 218, 200, 44, "▶  Importeer Informatieassets",          "GRC_Macros.ImportInformatieassets")
    add_button(ws_ie_com, 10, 306, 200, 44, "▶  Importeer Afhankelijke Assets",       "GRC_Macros.ImportAfhankelijkeAssets")
    add_button(ws_ie_com, 10, 394, 200, 44, "▶  Importeer Processen",                 "GRC_Macros.ImportProcessen")
    add_button(ws_ie_com, 10, 482, 200, 44, "▶  Importeer Kwetsbaarheden",            "GRC_Macros.ImportKwetsbaarheden")
    add_button(ws_ie_com, 10, 570, 200, 44, "▶  Importeer Links DA/Kwetsbaarheden",   "GRC_Macros.ImportLinksKwetsbaarheden")

    # Maak UserForm "AssetPicker" aan
    uf = wb_com.VBProject.VBComponents.Add(3)   # 3 = vbext_ct_MSForm
    uf.Name = "AssetPicker"
    uf.Properties("Caption").Value = "Selecteer Informatieassets"
    uf.Properties("Width").Value  = 340
    uf.Properties("Height").Value = 360

    # Label bovenaan
    lbl = uf.Designer.Controls.Add("Forms.Label.1")
    lbl.Name    = "lblTitel"
    lbl.Caption = "Selecteer de gekoppelde assets:"
    lbl.Left    = 10; lbl.Top = 10; lbl.Width = 300; lbl.Height = 18
    lbl.Font.Bold = True

    # ListBox met checkboxen (MultiSelect = 1, ListStyle = 1)
    lb = uf.Designer.Controls.Add("Forms.ListBox.1")
    lb.Name        = "lstAssets"
    lb.Left        = 10; lb.Top = 34; lb.Width = 306; lb.Height = 252
    lb.MultiSelect = 1   # fmMultiSelectMulti
    lb.ListStyle   = 1   # fmListStyleOption (checkboxen)

    # OK-knop (groen)
    btnOK = uf.Designer.Controls.Add("Forms.CommandButton.1")
    btnOK.Name    = "btnOK"
    btnOK.Caption = "OK"
    btnOK.Left    = 10;  btnOK.Top = 296; btnOK.Width = 146; btnOK.Height = 28
    btnOK.BackColor = int("15803D", 16)
    btnOK.ForeColor = int("FFFFFF", 16)

    # Annuleer-knop (rood)
    btnAnn = uf.Designer.Controls.Add("Forms.CommandButton.1")
    btnAnn.Name    = "btnAnnuleer"
    btnAnn.Caption = "Annuleer"
    btnAnn.Left    = 170; btnAnn.Top = 296; btnAnn.Width = 146; btnAnn.Height = 28
    btnAnn.BackColor = int("B91C1C", 16)
    btnAnn.ForeColor = int("FFFFFF", 16)

    # Voeg UserForm code toe
    uf.CodeModule.AddFromString(USERFORM_CODE)

    # Voeg Worksheet_SelectionChange toe aan Processes sheet module
    proc_code_name = wb_com.Sheets("Processes").CodeName
    proc_mod = wb_com.VBProject.VBComponents(proc_code_name)
    proc_mod.CodeModule.AddFromString(PROC_SHEET_CODE)

    # Voeg Worksheet_Activate toe aan Information Assets sheet module
    ia_code_name = wb_com.Sheets("Information Assets").CodeName
    ia_mod = wb_com.VBProject.VBComponents(ia_code_name)
    ia_mod.CodeModule.AddFromString(INFO_ASSET_SHEET_CODE)

    # Voeg Worksheet_Activate toe aan Dependent Assets sheet module
    dep_code_name = wb_com.Sheets("Dependent Assets").CodeName
    dep_mod = wb_com.VBProject.VBComponents(dep_code_name)
    dep_mod.CodeModule.AddFromString(DEP_ASSET_SHEET_CODE)

    # Voeg BeforeDoubleClick toggle toe aan Kwetsbaarheden sheet module
    kwets_code_name = wb_com.Sheets("Kwetsbaarheden").CodeName
    kwets_mod = wb_com.VBProject.VBComponents(kwets_code_name)
    kwets_mod.CodeModule.AddFromString(KWETS_SHEET_CODE)

    # Maak UserForm "VulnPicker" aan (kwetsbaarheidselectie + probabiliteit per vuln)
    vp = wb_com.VBProject.VBComponents.Add(3)   # 3 = vbext_ct_MSForm
    vp.Name = "VulnPicker"
    vp.Properties("Caption").Value = "Kwetsbaarheden selecteren"
    vp.Properties("Width").Value  = 458   # +28 pt (≈1 cm) rechts
    vp.Properties("Height").Value = 514   # +28 pt (≈1 cm) onder + 16 pt rand onder knoppen
    vp_des = vp.Designer
    # Titel-label
    lbl_vp = vp_des.Controls.Add("Forms.Label.1")
    lbl_vp.Name = "lblTitel"; lbl_vp.Caption = "Kwetsbaarheden"
    lbl_vp.Left = 8; lbl_vp.Top = 8; lbl_vp.Width = 426; lbl_vp.Height = 20
    lbl_vp.Font.Bold = True
    # Kolomkopjes
    lbl_vuln_hdr = vp_des.Controls.Add("Forms.Label.1")
    lbl_vuln_hdr.Name = "lblVulnHdr"; lbl_vuln_hdr.Caption = "Kwetsbaarheid"
    lbl_vuln_hdr.Left = 8; lbl_vuln_hdr.Top = 30; lbl_vuln_hdr.Width = 216; lbl_vuln_hdr.Height = 14
    lbl_vuln_hdr.Font.Italic = True
    lbl_prob_hdr = vp_des.Controls.Add("Forms.Label.1")
    lbl_prob_hdr.Name = "lblProbHdr"; lbl_prob_hdr.Caption = "Kans"
    lbl_prob_hdr.Left = 228; lbl_prob_hdr.Top = 30; lbl_prob_hdr.Width = 198; lbl_prob_hdr.Height = 14
    lbl_prob_hdr.Font.Italic = True
    # Scrollable Frame voor CheckBox + ComboBox rijen (inhoud wordt dynamisch gevuld in VBA)
    frm_vp = vp_des.Controls.Add("Forms.Frame.1")
    frm_vp.Name = "frmVulns"
    frm_vp.Caption = ""
    frm_vp.Left = 8; frm_vp.Top = 46; frm_vp.Width = 426; frm_vp.Height = 388
    frm_vp.ScrollBars = 2          # fmScrollBarsVertical
    frm_vp.KeepScrollBarsVisible = 2  # alleen verticaal tonen wanneer nodig
    # Annuleer-knop (links)
    btn_vp_ann = vp_des.Controls.Add("Forms.CommandButton.1")
    btn_vp_ann.Name = "btnAnnuleer"; btn_vp_ann.Caption = "Annuleer"
    btn_vp_ann.Left = 8; btn_vp_ann.Top = 446; btn_vp_ann.Width = 120; btn_vp_ann.Height = 30
    btn_vp_ann.BackColor = int("C0504D", 16)
    btn_vp_ann.ForeColor = int("FFFFFF", 16)
    # OK-knop (rechts)
    btn_vp_ok = vp_des.Controls.Add("Forms.CommandButton.1")
    btn_vp_ok.Name = "btnOK"; btn_vp_ok.Caption = "OK"
    btn_vp_ok.Left = 322; btn_vp_ok.Top = 446; btn_vp_ok.Width = 120; btn_vp_ok.Height = 30
    btn_vp_ok.BackColor  = int("70AD47", 16)
    btn_vp_ok.ForeColor  = int("FFFFFF", 16)
    vp.CodeModule.AddFromString(VULNPICKER_CODE)

    # Voeg SelectionChange toe aan RARM sheet module
    rarm_code_name = wb_com.Sheets("RARM").CodeName
    rarm_mod = wb_com.VBProject.VBComponents(rarm_code_name)
    rarm_mod.CodeModule.AddFromString(RARM_SHEET_CODE)

    # Sla op als .xlsm
    wb_com.SaveAs(str(OUT.resolve()), FileFormat=52)   # 52 = xlOpenXMLWorkbookMacroEnabled
    wb_com.Close(False)
    xl.Quit()
    pythoncom.CoUninitialize()

    OUT_XLSX.unlink()   # verwijder tussenstap
    vba_ok = True
    print(f"GRC Tool v0.3 (macro-enabled) opgeslagen: {OUT}")

except ImportError:
    print("INFO: pywin32 niet gevonden — bestand opgeslagen als .xlsx zonder macro's.")
    print(f"      Installeer pywin32 via: pip install pywin32")
except Exception as e:
    print(f"WAARSCHUWING: VBA-injectie mislukt ({e})")
    print(f"              Bestand opgeslagen als .xlsx: {OUT_XLSX}")

print(f"Gegenereerd door: {USERNAME} op {NOW_STR}")
