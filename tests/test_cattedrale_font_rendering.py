import os
import sys
import json
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtGui

import app.main as main


def test_cattedrale_font_applied(tmp_path):
    font_path = Path(__file__).resolve().parent.parent / "assets" / "fonts" / "Cattedrale[RUSbypenka220]-Regular.ttf"
    fid = QtGui.QFontDatabase.addApplicationFont(str(font_path))
    family = QtGui.QFontDatabase.applicationFontFamilies(fid)[0]

    main.CONFIG_PATH = str(tmp_path / "config.json")
    main.CONFIG["header_font"] = family
    main.CONFIG["sidebar_font"] = family
    with open(main.CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(main.CONFIG, f)

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()
    window.apply_settings()

    assert window.table.horizontalHeader().font().family() == family
    assert window.topbar.lbl_month.font().family() == family
    first_btn = window.sidebar.buttons[0]
    assert first_btn.font().family() == family

    window.close()
    app.quit()
