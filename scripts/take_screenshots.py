"""Take screenshots of every Streamlit page for the GRC presentation."""
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

OUT = Path("data/presentation_screenshots")
OUT.mkdir(parents=True, exist_ok=True)

BASE = "http://localhost:8502"

PAGES = [
    ("01_dashboard",            "🏠 Dashboard",             None),
    ("02_configuratie",         "⚙️ Configuratie",           None),
    ("03_processen_tab1",       "📋 Processen & Assets",     "Processen"),
    ("04_processen_tab2",       "📋 Processen & Assets",     "Informatieassets"),
    ("05_processen_tab3",       "📋 Processen & Assets",     "Afhankelijke Assets"),
    ("06_maatregelen",          "🎯 Maatregelen",            None),
    ("07_periodieke_controles", "🔍 Periodieke Controles",   None),
    ("08_acties",               "📋 Acties",                 None),
    ("09_kwetsbaarheden",       "🛡️ Kwetsbaarheden",         None),
    ("10_instellingen",         "🔧 Instellingen",           None),
]

def click_nav(page, label):
    """Click a sidebar radio button by label text."""
    page.locator(f"text={label}").first.click()
    time.sleep(2.5)

def click_tab(page, tab_label):
    page.get_by_role("tab", name=tab_label).click()
    time.sleep(1.5)

with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=True)
    ctx = browser.new_context(viewport={"width": 1440, "height": 900})
    p = ctx.new_page()

    print("Opening app…")
    p.goto(BASE, wait_until="networkidle", timeout=30000)
    time.sleep(3)

    for fname, nav_label, tab_label in PAGES:
        print(f"  > {fname}")
        click_nav(p, nav_label)
        if tab_label:
            click_tab(p, tab_label)
        # scroll to top
        p.evaluate("window.scrollTo(0,0)")
        time.sleep(0.5)
        p.screenshot(path=str(OUT / f"{fname}.png"), full_page=True)
        print(f"     saved {fname}.png")

    browser.close()

print(f"\nDone — {len(PAGES)} screenshots in {OUT}")
