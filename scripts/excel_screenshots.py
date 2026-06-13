"""Open GRC_Tool.xlsm in Excel and export key sheets as PNG via clipboard."""
import os, time, sys
from pathlib import Path
import win32com.client
import win32con
import win32gui
import win32api

OUT = Path("data/presentation_screenshots")
OUT.mkdir(parents=True, exist_ok=True)

XLSM = str(Path("data/Import/GRC_Tool.xlsm").resolve())

SHEETS = [
    ("xl_01_config",           "Config"),
    ("xl_02_processes",        "Processes"),
    ("xl_03_informatieassets", "Information Assets"),
    ("xl_04_dependent_assets", "Dependent Assets"),
    ("xl_05_rarm",             "RARM"),
    ("xl_06_kwetsbaarheden",   "Kwetsbaarheden"),
    ("xl_07_risicobeheer",     "Risicobeheer"),
    ("xl_08_controles",        "Controles"),
    ("xl_09_acties",           "Acties"),
    ("xl_10_import_export",    "Import & Export"),
]

xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False

try:
    wb = xl.Workbooks.Open(XLSM)
    time.sleep(2)

    for fname, sheet_name in SHEETS:
        try:
            ws = wb.Sheets(sheet_name)
            ws.Activate()
            xl.ActiveWindow.Zoom = 75
            time.sleep(1.2)

            # Export sheet to PNG via chart trick
            from PIL import ImageGrab
            import win32clipboard

            # Select used range and copy as picture
            used = ws.UsedRange
            used.CopyPicture(Appearance=1, Format=2)  # xlScreen, xlBitmap
            time.sleep(0.4)

            win32clipboard.OpenClipboard()
            try:
                img_data = win32clipboard.GetClipboardData(win32con.CF_DIB)
                win32clipboard.CloseClipboard()

                import io, struct
                from PIL import Image

                # DIB to PIL Image
                bfh = struct.pack('<2sIHHI', b'BM',
                                  len(img_data) + 14, 0, 0, 54)
                img = Image.open(io.BytesIO(bfh + img_data))
                out_path = OUT / f"{fname}.png"
                img.save(str(out_path))
                print(f"  saved {fname}.png  ({img.size})")
            except Exception as e:
                win32clipboard.CloseClipboard()
                # Fallback: screenshot of the Excel window
                print(f"  clipboard fallback for {fname}: {e}")
                hwnd = xl.ActiveWindow.Hwnd
                rect = win32gui.GetWindowRect(hwnd)
                img = ImageGrab.grab(bbox=rect)
                img.save(str(OUT / f"{fname}.png"))
                print(f"  saved {fname}.png (window grab)")
        except Exception as e:
            print(f"  SKIP {sheet_name}: {e}")

    wb.Close(SaveChanges=False)
finally:
    xl.Quit()

print(f"\nExcel screenshots done in {OUT}")
