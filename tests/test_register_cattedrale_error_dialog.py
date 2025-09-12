import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets
import resources


def test_register_cattedrale_error_dialog(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    called = {}
    monkeypatch.setattr(resources, "tk", None)
    monkeypatch.setattr(resources, "ctk", None)
    monkeypatch.setattr(resources, "tkfont", None)
    monkeypatch.setattr(QtWidgets.QMessageBox, "critical", lambda *a, **k: called.setdefault("called", True))

    fam = resources.register_cattedrale("dummy.ttf")
    assert fam == "Exo 2"
    assert called.get("called")
    app.quit()
