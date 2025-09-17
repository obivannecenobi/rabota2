import os
import sys
import json
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import resources

resources.register_fonts = lambda: None

import app.main as main


def test_neon_preset_updates_config_and_filters(tmp_path):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    original_path = main.CONFIG_PATH
    original_config = dict(main.CONFIG)
    dlg = None
    try:
        main.CONFIG.update(
            {
                "neon_size": 10,
                "neon_thickness": 1,
                "neon_intensity": 255,
            }
        )
        tmp_config = tmp_path / "config.json"
        main.CONFIG_PATH = str(tmp_config)
        with open(tmp_config, "w", encoding="utf-8") as f:
            json.dump(main.CONFIG, f)

        dlg = main.SettingsDialog()
        preset_index = dlg.combo_neon_preset.findText("Яркий")
        assert preset_index >= 0

        dlg.combo_neon_preset.setCurrentIndex(preset_index)
        QtWidgets.QApplication.processEvents()

        expected = dlg.combo_neon_preset.itemData(preset_index)
        assert expected is not None
        assert (
            main.CONFIG["neon_size"],
            main.CONFIG["neon_thickness"],
            main.CONFIG["neon_intensity"],
        ) == expected

        neon_filter = getattr(dlg.sld_neon_size, "_neon_filter", None)
        assert neon_filter is not None
        assert neon_filter._config["neon_size"] == expected[0]
    finally:
        if dlg is not None:
            dlg.close()
        main.CONFIG.clear()
        main.CONFIG.update(original_config)
        main.CONFIG_PATH = original_path
        app.quit()
