"""
detect_progid.py - cherche dans le registre les ProgIDs lies a Sogelink/GeoGex
"""
import winreg

mots = ["sogelink", "geogex", "cad", "icad", "intellicad"]

print("=== Recherche dans HKEY_CLASSES_ROOT ===")
found = []
try:
    root = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "")
    i = 0
    while True:
        try:
            name = winreg.EnumKey(root, i)
            nl = name.lower()
            if any(m in nl for m in mots):
                # Verifier si c'est un ProgID avec CLSID
                try:
                    sub = winreg.OpenKey(root, name + "\\CLSID")
                    clsid = winreg.QueryValue(sub, "")
                    print(f"  ProgID : {name}  -> CLSID {clsid}")
                    found.append(name)
                    winreg.CloseKey(sub)
                except:
                    pass
            i += 1
        except OSError:
            break
    winreg.CloseKey(root)
except Exception as e:
    print(f"Erreur registre : {e}")

if not found:
    print("Rien trouve avec ces mots-cles.")
    print("Essai avec juste 'cad'...")
    try:
        root = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "")
        i = 0
        while True:
            try:
                name = winreg.EnumKey(root, i)
                if "cad" in name.lower():
                    try:
                        sub = winreg.OpenKey(root, name + "\\CLSID")
                        clsid = winreg.QueryValue(sub, "")
                        print(f"  ProgID : {name}  -> CLSID {clsid}")
                        winreg.CloseKey(sub)
                    except:
                        pass
                i += 1
            except OSError:
                break
        winreg.CloseKey(root)
    except Exception as e:
        print(f"Erreur : {e}")
