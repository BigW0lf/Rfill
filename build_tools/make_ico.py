"""Convertit Rfill.png en Rfill.ico avec toutes les tailles standard Windows."""
from PIL import Image

img = Image.open("Rfill.png").convert("RGBA")
img.save(
    "Rfill.ico",
    format="ICO",
    sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)],
)
print("Rfill.ico généré.")
