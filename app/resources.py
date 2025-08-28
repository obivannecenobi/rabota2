import os
from PySide6.QtGui import QFontDatabase

FONTS_DIR = os.path.join(os.path.dirname(__file__), '..', 'assets', 'fonts')

def register_fonts():
    """Register all fonts located in the assets/fonts directory."""
    if not os.path.isdir(FONTS_DIR):
        return
    for fname in os.listdir(FONTS_DIR):
        if fname.lower().endswith(('.ttf', '.otf')):
            QFontDatabase.addApplicationFont(os.path.join(FONTS_DIR, fname))
