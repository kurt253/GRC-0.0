"""
GRC Tool v0.3 – Trainingspresentatie generator
Voegt screenshots van de echte tool in en bouwt een stap-voor-stap handleiding.
Uitvoer: data/template/GRC_Tool_Presentatie.pptx
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image as PILImage
import os

# ── Paden ────────────────────────────────────────────────────────────────────
SHOT_DIR = r"c:\Users\kurtm\Documents\GRC-0.0\data\screenshots"
OUT      = r"c:\Users\kurtm\Documents\GRC-0.0\data\template\GRC_Tool_Presentatie.pptx"

# ── Kleurpalet ────────────────────────────────────────────────────────────────
NAVY       = RGBColor(0x1F, 0x3B, 0x6E)
ACCENT     = RGBColor(0x2E, 0x86, 0xC1)
LIGHT_BLUE = RGBColor(0xD0, 0xDC, 0xF3)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
DARK_GREY  = RGBColor(0x2C, 0x2C, 0x2C)
MID_GREY   = RGBColor(0x5A, 0x5A, 0x5A)
GREEN      = RGBColor(0x1E, 0x8A, 0x44)
GREEN_LT   = RGBColor(0xD4, 0xED, 0xDA)
ORANGE     = RGBColor(0xD4, 0x6A, 0x1A)
ORANGE_LT  = RGBColor(0xFD, 0xF0, 0xE3)
RED        = RGBColor(0xC0, 0x39, 0x2B)
YELLOW     = RGBColor(0xF5, 0xCB, 0x5C)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


# ══ Hulpfuncties ══════════════════════════════════════════════════════════════

def slide_bg(slide, color=WHITE):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = color

def add_rect(slide, l, t, w, h, fill=None, line_color=None, line_width=Pt(1)):
    shape = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, l, t, w, h, size=13, bold=False, color=DARK_GREY,
             align=PP_ALIGN.LEFT, wrap=True, italic=False):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return tb

def add_bullets(slide, items, l, t, w, h, size=12, color=DARK_GREY, bullet="▸"):
    """Meerdere bullet-regels in één tekstvak."""
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(3)
        run = p.add_run()
        run.text = f"{bullet}  {item}"
        run.font.size = Pt(size)
        run.font.color.rgb = color

def add_numbered(slide, items, l, t, w, h, size=12, color=DARK_GREY):
    """Genummerde stappen."""
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(4)
        run = p.add_run()
        run.text = f"{i+1}.  {item}"
        run.font.size = Pt(size)
        run.font.color.rgb = color

def add_callout(slide, text, l, t, w, h, bg=ORANGE_LT, border=ORANGE, icon="⚠"):
    """Gekleurde waarschuwings- of infobalk."""
    add_rect(slide, l, t, w, h, fill=bg, line_color=border)
    add_text(slide, f"{icon}  {text}", l+0.12, t+0.05, w-0.2, h-0.1,
             size=11, color=DARK_GREY, wrap=True)

def add_tip(slide, text, l, t, w, h):
    add_callout(slide, text, l, t, w, h, bg=GREEN_LT, border=GREEN, icon="💡")

def add_warning(slide, text, l, t, w, h):
    add_callout(slide, text, l, t, w, h, bg=ORANGE_LT, border=ORANGE, icon="⚠")

def add_screenshot(slide, filename, l, t, max_w, max_h, border=True):
    """
    Voegt een screenshot in met behoud van beeldverhouding.
    Schaalt naar max_w × max_h (inches) zonder te vervormen.
    Geeft (werkelijke breedte, werkelijke hoogte) terug in inches.
    """
    path = os.path.join(SHOT_DIR, filename)
    if not os.path.exists(path):
        # Placeholder als screenshot ontbreekt
        r = add_rect(slide, l, t, max_w, max_h, fill=LIGHT_BLUE, line_color=ACCENT)
        add_text(slide, f"[screenshot: {filename}]", l+0.1, t+max_h/2-0.2,
                 max_w-0.2, 0.4, size=11, color=ACCENT, align=PP_ALIGN.CENTER)
        return max_w, max_h

    img = PILImage.open(path)
    px_w, px_h = img.size
    ratio = px_w / px_h

    # Schaal zodat het past binnen max_w × max_h
    if (max_w / ratio) <= max_h:
        w = max_w
        h = max_w / ratio
    else:
        h = max_h
        w = max_h * ratio

    pic = slide.shapes.add_picture(path, Inches(l), Inches(t), Inches(w), Inches(h))
    if border:
        pic.line.color.rgb = MID_GREY
        pic.line.width = Pt(0.75)
    return w, h

def header_bar(slide, step_nr, step_total, title, subtitle=None):
    """Genummerde kopbalk: 'Stap X/Y — Titel'."""
    add_rect(slide, 0, 0, 13.33, 1.15, fill=NAVY)
    # Stapnummer-badge
    add_rect(slide, 0.25, 0.15, 0.85, 0.85, fill=ACCENT)
    add_text(slide, f"{step_nr}", 0.25, 0.15, 0.85, 0.85,
             size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, f"/ {step_total}", 1.15, 0.38, 0.5, 0.4,
             size=11, color=LIGHT_BLUE, align=PP_ALIGN.LEFT)
    add_text(slide, title, 1.75, 0.1, 10.5, 0.65,
             size=24, bold=True, color=WHITE)
    if subtitle:
        add_text(slide, subtitle, 1.75, 0.72, 10.5, 0.38,
                 size=12, color=LIGHT_BLUE, italic=True)
    add_rect(slide, 0, 1.15, 13.33, 0.05, fill=ACCENT)

def footer(slide, note=None):
    add_rect(slide, 0, 7.1, 13.33, 0.4, fill=NAVY)
    txt = note or "GRC Tool v0.3  ·  Trainingsdocument  ·  Belgische Overheid"
    add_text(slide, txt, 0.3, 7.12, 12.7, 0.35,
             size=9, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)

def section_label(slide, text, l, t, w=3.5, color=NAVY):
    """Klein blauw sectielabel boven een blok."""
    add_rect(slide, l, t, w, 0.3, fill=color)
    add_text(slide, text, l+0.1, t+0.02, w-0.15, 0.28,
             size=9, bold=True, color=WHITE)

TOTAL = 13   # totaal aantal inhoudelijke slides


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — TITELSLIDE
# ══════════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
slide_bg(sl, NAVY)

add_rect(sl, 8.1, 0, 5.23, 7.5, fill=ACCENT)
add_rect(sl, 7.85, 0, 0.28, 7.5, fill=LIGHT_BLUE)

add_text(sl, "GRC Tool", 0.7, 1.2, 7, 1.4,
         size=52, bold=True, color=WHITE)
add_text(sl, "v0.3", 0.7, 2.6, 3, 0.7,
         size=34, color=LIGHT_BLUE)
add_text(sl, "Gebruikerstraining",
         0.7, 3.3, 7, 0.65, size=22, bold=True, color=WHITE)
add_text(sl,
    "Stap-voor-stap handleiding voor het importeren\nen beheren van GRC-data",
    0.7, 4.05, 7, 0.9, size=15, color=LIGHT_BLUE, wrap=True)

add_text(sl, "Inhoud van deze training", 8.4, 1.3, 4.5, 0.45,
         size=13, bold=True, color=WHITE)
for i, item in enumerate([
    "Opstarten & configuratie",
    "Sheets en navigatie",
    "Kwetsbaarheden importeren",
    "RARM lezen en invullen",
    "Links DA/Kwetsbaarheden",
    "Rapportage & afwijkingen",
    "Tips, valkuilen & veelgestelde vragen",
]):
    add_text(sl, f"›  {item}", 8.4, 1.85 + i*0.6, 4.6, 0.52,
             size=12, color=WHITE)

add_text(sl, "NIS2  ·  GDPR  ·  ISO 27001:2022  ·  BIO  ·  CyFun 2025",
         0.7, 6.8, 7, 0.4, size=10, color=LIGHT_BLUE, italic=True)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — OVERZICHT: ALLE SHEETS
# ══════════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
slide_bg(sl)
header_bar(sl, "—", TOTAL, "Overzicht van de tool",
           "Alle tabbladen in GRC_Tool.xlsm en hun functie")
footer(sl)

sheets = [
    (NAVY,   "Config",              "Taal instellen (NL/FR/EN), organisatienaam"),
    (NAVY,   "Info",                "Introductie, kaders en gebruiksaanwijzing"),
    (ACCENT, "Processen",           "Bedrijfsprocessen + CIA-classificatie"),
    (ACCENT, "Informatieassets",    "Informatieassets + vertrouwelijkheidsniveau"),
    (ACCENT, "Dependent Assets",    "Technische assets — basis voor RARM"),
    (ACCENT, "Verantwoordelijken",  "Contactregister GRC-verantwoordelijken"),
    (GREEN,  "Import & Export",     "Centrale pagina met alle importknoppen"),
    (GREEN,  "CyFun Controls",      "218 controls + mapping 2023↔2025"),
    (GREEN,  "Kwetsbaarheden",      "Matrix: kwetsbaarheden × controls (CIA + ✔)"),
    (GREEN,  "RARM",                "Risicomatrix per afhankelijke asset"),
    (MID_GREY,"Referentiewaarden",  "Definities van de 5 classificatieniveaus"),
    (ORANGE, "Controls 2023 - DA",  "Rapport: 2023-controls zonder 2025-equivalent"),
]

for i, (col, name, desc) in enumerate(sheets):
    r, c = divmod(i, 4)
    lft = 0.3  + c * 3.25
    top = 1.35 + r * 1.45
    add_rect(sl, lft, top, 0.15, 1.1, fill=col)
    add_text(sl, name, lft+0.25, top+0.05, 2.85, 0.38,
             size=11, bold=True, color=DARK_GREY)
    add_text(sl, desc, lft+0.25, top+0.45, 2.85, 0.55,
             size=10, color=MID_GREY, wrap=True, italic=True)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — STAP 1: OPSTARTEN & CONFIG
# ══════════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
slide_bg(sl)
header_bar(sl, 1, TOTAL, "Opstarten & configuratie",
           "Bouw het bestand en stel de taal in")
footer(sl)

# Linker kolom: stappen
section_label(sl, "WAT TE DOEN", 0.3, 1.3)
add_numbered(sl, [
    "Sluit Excel volledig af",
    "Open een terminal in de projectmap",
    "Voer uit:  python scripts\\build_grc.py",
    "Open het gegenereerde GRC_Tool.xlsm",
    'Klik "Inhoud inschakelen" in de gele balk',
    "Ga naar het tabblad Config",
    "Stel taal in op NL, FR of EN (cel D9)",
    "Vul organisatienaam en dienst in",
], 0.3, 1.65, 4.5, 5.1, size=12)

add_warning(sl,
    "Na een taalwijziging: herselect bestaande classificatiewaarden "
    "in de gegevensbladen — de codes (1–5) blijven correct, "
    "maar de labels zijn taalafhankelijk.",
    0.3, 6.2, 4.5, 0.75)

# Rechter kolom: screenshot Config
section_label(sl, "SCREENSHOT — Config-sheet", 5.2, 1.3, 7.8)
add_screenshot(sl, "01_config.png", 5.2, 1.65, 7.85, 5.3)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — STAP 2: IMPORT & EXPORT
# ══════════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
slide_bg(sl)
header_bar(sl, 2, TOTAL, "Import & Export — centraal startpunt",
           "Alle importknoppen staan op één tabblad")
footer(sl)

# Screenshot groot links
section_label(sl, "SCREENSHOT — Import & Export-sheet", 0.3, 1.3, 7.5)
add_screenshot(sl, "02_import_export.png", 0.3, 1.65, 7.5, 5.3)

# Rechts: uitleg knoppen
section_label(sl, "DE KNOPPEN UITGELEGD", 8.1, 1.3, 4.9)
knoppen = [
    ("Importeer Alles",             "IA → DA → Processen → Kwetsbaarheden in één keer",   ACCENT),
    ("Importeer Informatieassets",  "Enkel T - Information Assets",                         NAVY),
    ("Importeer Afhankelijke Assets","Enkel T - Dependent assets",                          NAVY),
    ("Importeer Processen",         "Processen + koppelingen IA/DA",                        NAVY),
    ("Importeer Kwetsbaarheden",    "Kwetsbaarheden + CIA + control-koppelingen",           GREEN),
    ("Importeer Links DA/Kwets.",   "Vult RARM + maakt Controls 2023 - DA rapport",        ORANGE),
]
for i, (btn, desc, col) in enumerate(knoppen):
    top = 1.65 + i * 0.82
    add_rect(sl, 8.1, top, 0.18, 0.5, fill=col)
    add_text(sl, btn,  8.35, top,       4.6, 0.3, size=11, bold=True,  color=DARK_GREY)
    add_text(sl, desc, 8.35, top+0.3,  4.6, 0.38, size=10, color=MID_GREY, italic=True, wrap=True)

add_warning(sl,
    "De oranje knop (Importeer Links DA/Kwets.) uitvoeren "
    "NADAT Afhankelijke Assets én Kwetsbaarheden zijn geïmporteerd.",
    8.1, 6.6, 4.9, 0.45)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — STAP 3: AFHANKELIJKE ASSETS
# ══════════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
slide_bg(sl)
header_bar(sl, 3, TOTAL, "Afhankelijke Assets importeren",
           "Technische assets — de basis voor de RARM-kolommen")
footer(sl)

section_label(sl, "SCREENSHOT — Dependent Assets-sheet", 0.3, 1.3, 7.5)
add_screenshot(sl, "07_dep_assets.png", 0.3, 1.65, 7.5, 4.0)

add_tip(sl,
    "Kolom B (DAName) is de sleutel: de naam hier moet exact overeenkomen "
    "met de naam in de Access-database (T - Dependent assets.DAName). "
    "De RARM-kolommen worden op basis van deze naam gesynchroniseerd.",
    0.3, 5.75, 7.5, 0.65)

section_label(sl, "WAT GEBEURT ER?", 8.1, 1.3, 4.9)
add_bullets(sl, [
    "Leest T - Dependent assets uit de .accdb",
    "Voegt elke asset toe als rij in de sheet",
    "Koppelt automatisch aan Processen en Informatieassets",
    "Maakt kolommen klaar in de RARM-sheet",
    "Bestaande data wordt niet gewist — alleen aangevuld",
], 8.1, 1.65, 4.9, 3.0, size=12)

section_label(sl, "ACCESS-TABEL", 8.1, 4.85, 4.9)
add_rect(sl, 8.1, 5.2, 4.9, 1.7, fill=LIGHT_BLUE)
for i, row in enumerate([
    "Tabel:  T - Dependent assets",
    "Veld:   ID  (intern — niet gebruikt voor LT-koppeling)",
    "Veld:   DAName  (naam — sleutel voor RARM-mapping)",
    "Aantal: 18 records in de demo-database",
]):
    add_text(sl, row, 8.25, 5.28 + i*0.38, 4.6, 0.35, size=10, color=NAVY)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — STAP 4: KWETSBAARHEDEN IMPORTEREN (OVERZICHT)
# ══════════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
slide_bg(sl)
header_bar(sl, 4, TOTAL, "Kwetsbaarheden importeren — overzicht",
           "Twee fasen: eerst CIA-impact, dan alle remediërende controls markeren")
footer(sl)

# Fase 1 blok
add_rect(sl, 0.3, 1.3, 6.0, 2.85, fill=LIGHT_BLUE)
add_rect(sl, 0.3, 1.3, 6.0, 0.45, fill=NAVY)
add_text(sl, "FASE 1 — Kwetsbaarheden & CIA-impact", 0.45, 1.33, 5.7, 0.38,
         size=13, bold=True, color=WHITE)
fase1 = [
    "Scant bestaande kolommen (rij 2) — naam als sleutel",
    "Kwetsbaarheid al aanwezig? → bijwerken op dezelfde positie",
    "Nieuwe kwetsbaarheid? → nieuwe kolom achteraan",
    "Rijen 3–5: CIA-impact bijwerken (✔ met kleur of grijs)",
    "Rijen 6+: bestaande ✔-marks wissen (fase 2 vult opnieuw)",
    "Verwijdert orphan-kolommen (niet meer in database)",
    "Herberekent de samengevoegde titelrij",
]
add_bullets(sl, fase1, 0.45, 1.82, 5.65, 2.2, size=11, color=NAVY)

# Fase 2 blok
add_rect(sl, 0.3, 4.3, 6.0, 2.85, fill=GREEN_LT)
add_rect(sl, 0.3, 4.3, 6.0, 0.45, fill=GREEN)
add_text(sl, "FASE 2 — Controls markeren (✔)", 0.45, 4.33, 5.7, 0.38,
         size=13, bold=True, color=WHITE)
fase2 = [
    "Laadt volledige LT-tabel in geheugen (één DB-query)",
    "Statusbalk toont actieve kwetsbaarheid: '(3/37): Phishing'",
    "Per kwetsbaarheid: zoek alle gekoppelde controls op",
    "Vertaalketen:  ctrlRef → refNrMap → 2023-ID → rev23to25 → 2025-ID → rij",
    "Gevonden → ✔ (groen) in de juiste cel",
    "Niet gevonden → geregistreerd in afwijkingsrapport",
]
add_bullets(sl, fase2, 0.45, 4.82, 5.65, 2.2, size=11, color=DARK_GREY)

# Vertaalketen diagram rechts
section_label(sl, "VERTAALKETEN — DETAIL", 6.7, 1.3, 6.3)
add_rect(sl, 6.7, 1.75, 6.3, 5.4, fill=NAVY)
chain = [
    ("ctrlRef",    "Getal uit LT-tabel (bv. 10)",                    YELLOW),
    ("refNrMap",   "T-CyFunEssentiel: 10 → 'ID.AM-3.1'",            LIGHT_BLUE),
    ("NormId23",   "Normaliseer: 'id.am-3.1'  (lowercase, strip)",   LIGHT_BLUE),
    ("rev23to25",  "CyFun Controls-sheet: → 'id.am-07.1'",           LIGHT_BLUE),
    ("ctrlRowMap", "Kwetsbaarheden sheet: → rij 60",                  LIGHT_BLUE),
    ("Resultaat",  "✔ in cel (rij 60, kwetsbaarheid-kolom)",         GREEN_LT),
]
for i, (lbl, desc, col) in enumerate(chain):
    top = 1.95 + i * 0.82
    add_rect(sl, 6.9, top, 1.5, 0.42, fill=col)
    add_text(sl, lbl,  6.92, top+0.04, 1.45, 0.34, size=10, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_text(sl, desc, 8.5,  top+0.06, 4.3,  0.3,  size=10, color=WHITE)
    if i < 5:
        add_text(sl, "↓", 7.55, top+0.44, 0.4, 0.36, size=14, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — STAP 4b: KWETSBAARHEDEN-SHEET IN DETAIL
# ══════════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
slide_bg(sl)
header_bar(sl, 4, TOTAL, "Kwetsbaarheden-sheet — na import",
           "Wat je ziet nadat de import klaar is")
footer(sl)

section_label(sl, "SCREENSHOT — Kwetsbaarheden-sheet (rijen 1–15)", 0.3, 1.3, 8.5)
sw, sh = add_screenshot(sl, "03_kwetsbaarheden.png", 0.3, 1.65, 12.7, 3.5)

# Uitleg rijen
section_label(sl, "RIJSTRUCTUUR", 0.3, 5.35, 4.0)
add_rect(sl, 0.3, 5.7, 12.7, 1.55, fill=LIGHT_BLUE)
rows_info = [
    ("Rij 1", "Titelrij (samengevoegd over alle kolommen)",          NAVY),
    ("Rij 2", "Naam van de kwetsbaarheid (kolomhoofd)",              ACCENT),
    ("Rij 3", "Confidentialiteit-impact  (✔ = van toepassing)",     ACCENT),
    ("Rij 4", "Integriteit-impact",                                  ACCENT),
    ("Rij 5", "Beschikbaarheid-impact",                              ACCENT),
    ("Rij 6+","Één rij per CyFun 2025 control — ✔ = remediëert deze kwetsbaarheid", GREEN),
]
for i, (lbl, desc, col) in enumerate(rows_info):
    lft = 0.45 + (i % 3) * 4.2
    top = 5.75 + (i // 3) * 0.5
    add_rect(sl, lft, top, 0.75, 0.38, fill=col)
    add_text(sl, lbl,  lft+0.02, top+0.04, 0.7, 0.3, size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(sl, desc, lft+0.82, top+0.06, 3.2, 0.3, size=10, color=DARK_GREY)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — STAP 5: CYFUN CONTROLS
# ══════════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
slide_bg(sl)
header_bar(sl, 5, TOTAL, "CyFun Controls raadplegen",
           "218 controls met mapping 2023 ↔ 2025")
footer(sl)

section_label(sl, "SCREENSHOT — CyFun Controls-sheet (eerste rijen)", 0.3, 1.3, 8.5)
add_screenshot(sl, "05_cyfun_controls.png", 0.3, 1.65, 8.3, 4.1)

section_label(sl, "KOLOMSTRUCTUUR", 8.85, 1.3, 4.15)
add_rect(sl, 8.85, 1.65, 4.15, 4.1, fill=NAVY)
cols_info = [
    ("A",  "CyFun 2025 — Categorie"),
    ("B",  "CyFun 2025 — Subcategorie"),
    ("C",  "CyFun 2025 — Requirement (ID)"),
    ("D",  "CyFun 2025 — Requirement tekst"),
    ("E",  "Assurance-niveau"),
    ("F",  "CyFun 2025 Requirement (vol)  ← mapping-sleutel"),
    ("G–L","ISO 27001 / NIS2 / BIO / GDPR referenties"),
    ("M",  "CyFun 2023 Requirement  ← 2023-ID bron"),
    ("N",  "Versie-status (2023 only / nieuw / gewijzigd)"),
]
for i, (col, desc) in enumerate(cols_info):
    add_rect(sl, 8.95, 1.72 + i*0.41, 0.38, 0.35, fill=ACCENT)
    add_text(sl, col,  8.96, 1.75 + i*0.41, 0.36, 0.28, size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(sl, desc, 9.4,  1.77 + i*0.41, 3.5,  0.28, size=9.5, color=WHITE)

add_tip(sl,
    "Kolom F (2025-ID) en M (2023-ID) zijn de sleutels voor de automatische "
    "vertaling in de importmacro. Wijzig deze kolommen niet.",
    0.3, 5.85, 12.7, 0.55)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — STAP 6: RARM LEZEN
# ══════════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
slide_bg(sl)
header_bar(sl, 6, TOTAL, "RARM — Risk Assessment & Remediation Matrix",
           "Risicomatrix per afhankelijke asset lezen en gebruiken")
footer(sl)

section_label(sl, "SCREENSHOT — RARM-sheet (rijen 1–20)", 0.3, 1.3, 8.5)
add_screenshot(sl, "04_rarm.png", 0.3, 1.65, 8.3, 4.0)

section_label(sl, "STRUCTUUR & INTERACTIE", 8.85, 1.3, 4.15)
add_rect(sl, 8.85, 1.65, 4.15, 5.1, fill=LIGHT_BLUE)

add_text(sl, "Rijstructuur", 9.0, 1.72, 3.8, 0.32, size=11, bold=True, color=NAVY)
struct = [
    "Rij 1 — Titelrij (samengevoegd)",
    "Rij 2 — Naam afhankelijke asset",
    "Rij 3 — Gekoppelde kwetsbaarheid",
    "Rij 4+ — 218 CyFun 2025 controls",
]
for i, s in enumerate(struct):
    add_text(sl, s, 9.05, 2.1 + i*0.36, 3.8, 0.32, size=10, color=NAVY)

add_text(sl, "Interactie", 9.0, 3.6, 3.8, 0.32, size=11, bold=True, color=NAVY)
interact = [
    "Dubbelklik op een cel → ✔ aan/uit zetten",
    "Navigeren naar sheet → DA-kolommen auto-sync",
    "Klik op DA-naam (rij 2) → VulnPicker popup",
    "VulnPicker: selecteer kwetsbaarheid + probabiliteit",
]
for i, s in enumerate(interact):
    add_text(sl, s, 9.05, 4.0 + i*0.38, 3.8, 0.32, size=10, color=NAVY)

add_tip(sl,
    "Importeer eerst Afhankelijke Assets (stap 3), daarna Kwetsbaarheden (stap 4), "
    "en tot slot Links DA/Kwetsbaarheden (stap 7) — in deze volgorde.",
    0.3, 5.75, 8.3, 0.55)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 10 — STAP 7: LINKS DA/KWETSBAARHEDEN
# ══════════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
slide_bg(sl)
header_bar(sl, 7, TOTAL, "Importeer Links DA / Kwetsbaarheden",
           "Vult de RARM en genereert het Controls 2023 - DA rapport")
footer(sl)

# Workflow pijlen
steps = [
    (NAVY,   "Access\nDatabase",      "LT - Selected\ncontrols to DA\n725 rijen"),
    (ACCENT, "DAID\nkoppeling",        "DAID → DAName\n→ RARM-kolom"),
    (ACCENT, "Control\nvertaling",     "ctrlRef → 2023-ID\n→ 2025-ID → rij"),
    (GREEN,  "RARM\nbijwerken",        "✔ in rij × kolom\nbij gevonden controls"),
    (ORANGE, "Rapport\naanmaken",      "Controls 2023 - DA\nniet-gematchte controls"),
]
for i, (col, title, sub) in enumerate(steps):
    lft = 0.3 + i * 2.55
    add_rect(sl, lft, 1.35, 2.3, 0.85, fill=col)
    add_text(sl, title, lft+0.1, 1.38, 2.1, 0.8, size=11, bold=True,
             color=WHITE, align=PP_ALIGN.CENTER, wrap=True)
    add_rect(sl, lft, 2.2, 2.3, 0.9, fill=LIGHT_BLUE)
    add_text(sl, sub,  lft+0.1, 2.24, 2.1, 0.84, size=10, color=NAVY,
             align=PP_ALIGN.CENTER, wrap=True)
    if i < 4:
        add_text(sl, "→", lft+2.3, 1.65, 0.25, 0.45,
                 size=20, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

# Wat verschijnt er nadien
add_rect(sl, 0.3, 3.35, 12.7, 0.38, fill=NAVY)
add_text(sl, "Resultaat na import", 0.5, 3.38, 12.2, 0.32,
         size=12, bold=True, color=WHITE)

add_rect(sl, 0.3, 3.73, 6.15, 3.0, fill=LIGHT_BLUE)
add_text(sl, "RARM — bijgewerkt", 0.45, 3.78, 5.8, 0.35, size=11, bold=True, color=NAVY)
rarm_r = [
    "Rij 3 per DA: kwetsbaarheid ingevuld",
    "Controls met ✔ in de juiste DA-kolom",
    "Bestaande handmatige ✔ blijven behouden",
    "Statistieken in MsgBox: X controls gekoppeld",
]
add_bullets(sl, rarm_r, 0.45, 4.2, 5.8, 2.4, size=11, color=NAVY)

add_rect(sl, 6.75, 3.73, 6.25, 3.0, fill=GREEN_LT)
add_text(sl, "Controls 2023 - DA — nieuw tabblad", 6.9, 3.78, 5.9, 0.35,
         size=11, bold=True, color=DARK_GREY)
rep_r = [
    "Kolom A: naam afhankelijke asset",
    "Kolom B: 2023 Control-ID",
    "Kolom C: reden (Enkel in CyFun 2023 / Niet gevonden)",
    "Wordt overschreven bij elke import",
    "Enkel unieke combinaties (geen duplicaten)",
]
add_bullets(sl, rep_r, 6.9, 4.2, 5.9, 2.4, size=11, color=DARK_GREY)

add_warning(sl,
    "Dit tabblad wordt aangemaakt door de macro — het bestaat nog niet "
    "in een vers gebouwd bestand.",
    0.3, 6.8, 12.7, 0.45)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 11 — STAP 8: CONTROLS 2023 - DA RAPPORT
# ══════════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
slide_bg(sl)
header_bar(sl, 8, TOTAL, "Controls 2023 - DA — het afwijkingsrapport",
           "Welke 2023-controls hebben geen equivalent in CyFun 2025?")
footer(sl)

# Linker uitleg
section_label(sl, "WAT STAAT ER IN?", 0.3, 1.3, 5.5)
add_rect(sl, 0.3, 1.65, 5.5, 5.1, fill=LIGHT_BLUE)
add_text(sl, "Het rapport bevat drie kolommen:", 0.5, 1.72, 5.1, 0.38,
         size=12, bold=True, color=NAVY)

cols_report = [
    ("A", "Dependent Asset",    "Naam van de technische asset waarvoor de control geselecteerd was"),
    ("B", "2023 Control-ID",    "Het CyFun 2023 ID dat niet gemapped kon worden (bv. 'pr.at-3.1')"),
    ("C", "Reden",              "'Enkel in CyFun 2023' — control bestaat niet in 2025\n"
                                "'Niet gevonden in 2025' — onbekend formaat (bv. 'pr.ds-2.1.a')"),
]
for i, (col, name, desc) in enumerate(cols_report):
    top = 2.2 + i * 1.4
    add_rect(sl, 0.4, top, 0.55, 0.55, fill=NAVY)
    add_text(sl, col,  0.4, top, 0.55, 0.55, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(sl, name, 1.05, top+0.04, 4.5, 0.3,  size=11, bold=True, color=NAVY)
    add_text(sl, desc, 1.05, top+0.36, 4.5, 0.85, size=10, color=DARK_GREY, wrap=True)

# Rechts: wat te doen met de afwijkingen
section_label(sl, "WAT TE DOEN MET AFWIJKINGEN?", 6.2, 1.3, 6.8)
add_rect(sl, 6.2, 1.65, 6.8, 2.3, fill=NAVY)
add_text(sl, "Enkel in CyFun 2023", 6.4, 1.7, 6.4, 0.38, size=12, bold=True, color=YELLOW)
actions_23 = [
    "De control bestaat niet meer in CyFun 2025",
    "Controleer of er een functioneel equivalent is in de 2025-set",
    "Documenteer de keuze: bewust niet overgenomen, of vervangen?",
    "Geen actie nodig in de tool — enkel organisatorische opvolging",
]
add_bullets(sl, actions_23, 6.4, 2.15, 6.4, 1.65, size=11, color=WHITE, bullet="→")

add_rect(sl, 6.2, 4.1, 6.8, 2.3, fill=LIGHT_BLUE)
add_text(sl, "Niet gevonden in 2025", 6.4, 4.15, 6.4, 0.38, size=12, bold=True, color=RED)
actions_nf = [
    "Subvarianten met .a/.b-suffix (bv. 'pr.ds-2.1.a')",
    "Niet opgenomen in de officiële 2025-controleset",
    "Melden aan de databeheerder voor correctie in de .accdb",
    "Verwijder of herstel de koppeling in de Access-database",
]
add_bullets(sl, actions_nf, 6.4, 4.56, 6.4, 1.7, size=11, color=NAVY, bullet="→")

add_tip(sl,
    "Dit rapport geeft directe input voor het updaten van de Access-database "
    "bij de overgang van CyFun 2023 naar CyFun 2025.",
    0.3, 6.85, 12.7, 0.52)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 12 — VALKUILEN & VEELGESTELDE VRAGEN
# ══════════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
slide_bg(sl)
header_bar(sl, 9, TOTAL, "Valkuilen & veelgestelde vragen",
           "De meest voorkomende problemen en hun oplossing")
footer(sl)

problemen = [
    (RED,    "Build mislukt — 'bestand vergrendeld'",
     "Excel staat nog open. Sluit Excel volledig:\nStop-Process -Name 'EXCEL' -Force"),
    (ORANGE, "Macro's werken niet",
     "Klik 'Inhoud inschakelen' in de gele Excel-balk.\nKijk of macro's zijn ingeschakeld in Trust Center."),
    (ORANGE, "'Tabel niet gevonden' bij import",
     "Verkeerde .accdb geselecteerd, of tabelnaam wijkt af.\nControleer: tabel heet exact [T - Vulnerabilities]."),
    (ORANGE, "ACE-driver fout (3706)",
     "Installeer MS Access Database Engine 2016,\nzelfde bitness (32/64-bit) als Excel."),
    (RED,    "RARM blijft leeg na import",
     "Importeer eerst Afhankelijke Assets — kolom B\nvan Dependent Assets-sheet moet gevuld zijn."),
    (MID_GREY,"Labels tonen [sleutelcode]",
     "Sleutel ontbreekt in _Lang-sheet.\nHerbouw via python scripts\\build_grc.py."),
    (MID_GREY,"Classificatielabels fout na taalwijziging",
     "Codes (1–5) zijn correct, labels zijn taalafhankelijk.\nHerselect via dropdown in de betreffende sheet."),
    (NAVY,   "Controls 2023 - DA bestaat niet",
     "Dit sheet wordt aangemaakt door de macro.\nVoer eerst 'Importeer Links DA/Kwets.' uit."),
    (NAVY,   "Weinig ✔ in Kwetsbaarheden-sheet",
     "Controleer of LT.Vulnerability koppelt aan\nT-Vulnerabilities.Reference (niet aan .ID)."),
]

for i, (col, vraag, antw) in enumerate(problemen):
    r, c = divmod(i, 3)
    lft = 0.3  + c * 4.35
    top = 1.35 + r * 1.9
    add_rect(sl, lft, top, 4.1, 0.38, fill=col)
    add_text(sl, vraag, lft+0.1, top+0.04, 3.9, 0.3, size=9.5, bold=True, color=WHITE, wrap=True)
    add_rect(sl, lft, top+0.38, 4.1, 1.38, fill=LIGHT_BLUE)
    add_text(sl, antw, lft+0.1, top+0.44, 3.9, 1.25, size=10, color=DARK_GREY, wrap=True)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 13 — IMPORTVOLGORDE SAMENVATTING
# ══════════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
slide_bg(sl)
header_bar(sl, 10, TOTAL, "Correcte importvolgorde — samenvatting",
           "Volg deze volgorde voor een correcte vulling van alle sheets")
footer(sl)

steps_full = [
    (NAVY,   "1", "python scripts\\build_grc.py",
     "Herbouw de .xlsm. Altijd uitvoeren na een wijziging aan het build-script.",
     "Terminal (PowerShell)"),
    (NAVY,   "2", "Open GRC_Tool.xlsm",
     "Klik 'Inhoud inschakelen'. Ga naar Config en stel taal/organisatie in.",
     "Excel"),
    (ACCENT, "3", "Importeer Informatieassets",
     "Selecteer de .accdb. Laadt T - Information Assets + koppelingen.",
     "Import & Export-sheet"),
    (ACCENT, "4", "Importeer Afhankelijke Assets",
     "Vereist vóór RARM-sync. Laadt T - Dependent assets.",
     "Import & Export-sheet"),
    (ACCENT, "5", "Importeer Processen",
     "Laadt T - Processes in scope + CIA-classificatie.",
     "Import & Export-sheet"),
    (GREEN,  "6", "Importeer Kwetsbaarheden",
     "Fase 1: CIA-impact. Fase 2: alle controls markeren. Genereert Afwijkingen-sheet.",
     "Import & Export-sheet"),
    (ORANGE, "7", "Importeer Links DA/Kwetsbaarheden",
     "Vult RARM. Genereert Controls 2023 - DA rapport. Uitvoeren ALS LAATSTE.",
     "Import & Export-sheet"),
]

for i, (col, nr, action, desc, where) in enumerate(steps_full):
    r = i % 4
    c = i // 4
    lft = 0.3 + c * 6.55
    top = 1.35 + r * 1.45

    add_rect(sl, lft, top, 0.62, 1.15, fill=col)
    add_text(sl, nr, lft, top, 0.62, 1.15, size=26, bold=True,
             color=WHITE, align=PP_ALIGN.CENTER)
    add_text(sl, action, lft+0.72, top+0.04, 5.6, 0.38,
             size=12, bold=True, color=DARK_GREY)
    add_text(sl, desc,   lft+0.72, top+0.42, 5.6, 0.48,
             size=10, color=MID_GREY, wrap=True)
    add_rect(sl, lft+0.72, top+0.9, 5.6, 0.22, fill=LIGHT_BLUE)
    add_text(sl, f"📍 {where}", lft+0.78, top+0.91, 5.4, 0.2,
             size=9, color=NAVY)

add_warning(sl,
    "Stap 7 altijd uitvoeren NADAT stap 4 (Afhankelijke Assets) "
    "én stap 6 (Kwetsbaarheden) zijn afgerond.",
    0.3, 7.0, 12.7, 0.42)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 14 — AFSLUITING
# ══════════════════════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(BLANK)
slide_bg(sl, NAVY)
add_rect(sl, 0, 3.15, 13.33, 0.08, fill=ACCENT)

add_text(sl, "Klaar voor gebruik!", 1.0, 1.0, 11.3, 1.1,
         size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(sl, "Bedankt voor het volgen van de training.",
         1.0, 2.2, 11.3, 0.6, size=20, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)
add_text(sl, "kurt.maekelberghe@gmail.com",
         1.0, 3.4, 11.3, 0.55, size=16, color=WHITE, align=PP_ALIGN.CENTER, italic=True)

# Snelle referentiekaart
add_rect(sl, 1.5, 4.15, 10.33, 2.6, fill=RGBColor(0x15, 0x2B, 0x55))
add_text(sl, "Snelle referentie", 1.7, 4.22, 9.9, 0.38,
         size=13, bold=True, color=LIGHT_BLUE)
refs = [
    "Build-commando  :  python scripts\\build_grc.py",
    "Excel sluiten   :  Stop-Process -Name 'EXCEL' -Force",
    "Importvolgorde  :  IA → DA → Processen → Kwetsbaarheden → Links DA",
    "Rapporten       :  Afwijkingen-sheet  +  Controls 2023 - DA",
]
for i, r in enumerate(refs):
    add_text(sl, r, 1.7, 4.66 + i*0.46, 9.9, 0.4, size=11, color=WHITE)

add_text(sl, "NIS2  ·  GDPR  ·  ISO 27001:2022  ·  BIO  ·  CyFun 2025",
         1.0, 6.9, 11.3, 0.4, size=11, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)


# ── Opslaan ──────────────────────────────────────────────────────────────────
prs.save(OUT)
print(f"Presentatie opgeslagen: {OUT}")
print(f"Aantal slides: {len(prs.slides)}")
