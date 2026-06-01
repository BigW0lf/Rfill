"""
creer_fas.py v3
Ouvre le VLIDE dans AutoCAD et compile le LSP en FAS via pyautogui.
AutoCAD doit etre ouvert avec un dessin actif.
"""
import sys, time, os

LSP = r"C:\05_devtopo\plugin_autocad\CONTOUR-PIECE.lsp"
FAS = r"C:\05_devtopo\plugin_autocad\CONTOUR-PIECE.fas"

try:
    import win32com.client, win32gui, win32con
except ImportError:
    print("pip install pywin32")
    sys.exit(1)

try:
    import pyautogui
except ImportError:
    print("pip install pyautogui")
    sys.exit(1)

pyautogui.PAUSE = 0.3

# --- Connexion AutoCAD ---
app = None
for pid in ["AutoCAD.Application.23", "AutoCAD.Application",
            "AutoCAD.Application.25", "AutoCAD.Application.26"]:
    try:
        app = win32com.client.GetActiveObject(pid)
        print(f"Connecte : {pid}")
        break
    except Exception:
        pass

if not app:
    print("AutoCAD non trouve. Ouvre AutoCAD avec un dessin actif.")
    sys.exit(1)

doc = app.ActiveDocument
print(f"Document : {doc.Name}")

# --- Mettre AutoCAD au premier plan ---
hwnd = app.HWND
win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
win32gui.SetForegroundWindow(hwnd)
time.sleep(1)

# --- Ouvrir le VLIDE via la ligne de commande AutoCAD ---
print("Ouverture VLIDE...")
doc.SendCommand("VLIDE\n")
time.sleep(3)

# --- Trouver la fenetre VLIDE ---
vlide_hwnd = None
def enum_cb(hwnd, _):
    global vlide_hwnd
    title = win32gui.GetWindowText(hwnd)
    if "visual lisp" in title.lower() or "vlisp" in title.lower() or "vlide" in title.lower():
        vlide_hwnd = hwnd

win32gui.EnumWindows(enum_cb, None)

if vlide_hwnd:
    print(f"VLIDE trouve (hwnd={vlide_hwnd})")
    win32gui.SetForegroundWindow(vlide_hwnd)
    time.sleep(1)
else:
    print("VLIDE non trouve comme fenetre separee, on tente quand meme...")

# --- Taper la commande de compilation dans la console VLIDE ---
lsp_fwd = LSP.replace("\\", "/")
cmd = f'(vlisp-compile \'st "{lsp_fwd}")'
print(f"Commande : {cmd}")

# Cliquer dans la zone console VLIDE (bas de l'ecran) et taper
pyautogui.hotkey('alt', 'tab')
time.sleep(0.5)
pyautogui.typewrite(cmd, interval=0.03)
pyautogui.press('enter')
time.sleep(5)

# --- Verifier le resultat ---
if os.path.exists(FAS):
    sz = os.path.getsize(FAS)
    print(f"\nSUCCES : {FAS}  ({sz} octets)")
    print("\nDans GeoGex :")
    print("  APPLOAD -> selectionner CONTOUR-PIECE.fas")
    print("  Puis taper : CP")
else:
    print("\nFAS non cree.")
    print("\n--- ALTERNATIVE MANUELLE (30 secondes) ---")
    print("1. Dans AutoCAD, tape : VLIDE  puis Entree")
    print("2. Dans la console VLIDE (en bas), tape exactement :")
    print(f'   (vlisp-compile \'st "{lsp_fwd}")')
    print("3. Appuie sur Entree")
    print(f"4. Le fichier {FAS} sera cree")
