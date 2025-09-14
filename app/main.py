# -*- coding: utf-8 -*-
import sys
import os
import json
import calendar
import logging
from datetime import datetime, date
from typing import Dict, List, Union, Iterable

from PySide6 import QtWidgets, QtGui, QtCore
import shiboken6
from dataclasses import dataclass, field

from widgets import StyledPushButton, StyledToolButton
from resources import register_fonts, load_icons, icon
import theme_manager
from effects import NeonEventFilter, neon_enabled, update_neon_filters, apply_neon_effect

logger = logging.getLogger(__name__)

ASSETS = os.path.join(os.path.dirname(__file__), "..", "assets")
DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")
CONFIG_PATH = os.path.join(DATA_DIR, "config.json")


def load_config():
    default = {
        "neon": False,
        "neon_size": 10,
        "neon_thickness": 1,
        "neon_intensity": 255,
        "accent_color": "#39ff14",
        "gradient_colors": ["#39ff14", "#2d7cdb"],
        "gradient_angle": 0,
        "font_family": "Exo 2",
        "header_font": "Cattedrale [RUS by penka220]",
        "text_font": "Exo 2",
        "sidebar_font": "Cattedrale [RUS by penka220]",
        "save_path": DATA_DIR,
        "day_rows": 6,
        "workspace_color": "#1e1e21",
        "sidebar_color": "#1f1f23",
        "sidebar_icon": os.path.join(ASSETS, "gpt_icon.png"),
        "app_icon": os.path.join(ASSETS, "gpt_icon.png"),
    }
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            save_path = data.get("save_path")
            if save_path:
                if not os.path.isabs(save_path):
                    save_path = os.path.abspath(
                        os.path.join(os.path.dirname(CONFIG_PATH), save_path)
                    )
                data["save_path"] = save_path
            default.update({k: v for k, v in data.items() if v is not None})
            for key in ("monochrome", "mono_saturation", "theme"):
                default.pop(key, None)
        except Exception:
            pass
    else:
        try:
            os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(default, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    return default


CONFIG = load_config()
BASE_SAVE_PATH = os.path.abspath(CONFIG.get("save_path", DATA_DIR))

def ensure_year_dirs(year):
    base = os.path.join(BASE_SAVE_PATH, str(year))
    for sub in ("stats", "release", "top", "year"):
        os.makedirs(os.path.join(base, sub), exist_ok=True)
    return base


def apply_neon_to_inputs(root: QtWidgets.QWidget) -> None:
    """Install neon event filters on editable input widgets under ``root``."""

    if not neon_enabled() or root is None or not shiboken6.isValid(root):
        return

    targets = (
        QtWidgets.QLineEdit,
        QtWidgets.QPlainTextEdit,
        QtWidgets.QTextEdit,
        QtWidgets.QAbstractSpinBox,
        QtWidgets.QComboBox,
        QtWidgets.QCheckBox,
        QtWidgets.QSlider,
    )

    widgets = [root] + root.findChildren(QtWidgets.QWidget)
    for w in widgets:
        if not isinstance(w, targets):
            continue
        if getattr(w, "_neon_filter", None) is not None:
            continue
        w.setAttribute(QtCore.Qt.WA_Hover, True)
        filt = NeonEventFilter(w)
        w.installEventFilter(filt)
        w._neon_filter = filt

def stats_dir(year):
    return os.path.join(ensure_year_dirs(year), "stats")

def release_dir(year):
    return os.path.join(ensure_year_dirs(year), "release")

def top_dir(year):
    return os.path.join(ensure_year_dirs(year), "top")

def year_dir(year):
    return os.path.join(ensure_year_dirs(year), "year")
ICON_TOGGLE = os.path.join(ASSETS, "gpt_icon.png")
ICON_TM   = os.path.join(ASSETS, "ic_tm.png")
ICON_TQ   = os.path.join(ASSETS, "ic_tq.png")
ICON_TP   = os.path.join(ASSETS, "ic_tp.png")
ICON_TG   = os.path.join(ASSETS, "ic_tg.png")
ICON_VYK  = os.path.join(ASSETS, "ic_vykl.png")
VERSION_FILE = os.path.join(os.path.dirname(__file__), "..", "VERSION")

RU_MONTHS = ["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"]


def ensure_font_registered(
    family: str, parent: QtWidgets.QWidget | None = None
) -> str:
    """Ensure *family* is available and return a usable family name.

    A missing font previously triggered a file-selection dialog which blocks in
    headless environments.  The function now silently falls back to a bundled
    alternative, preferring ``"Exo 2"`` then ``"Arial"``.  If neither is
    installed, the application's default font is used.  If Qt can resolve the
    requested *family* to an available font internally, the original family name
    is preserved so configuration values remain consistent.
    """
    if family in QtGui.QFontDatabase.families():
        return family
    app = QtWidgets.QApplication.instance()
    if app is not None:
        resolved = QtGui.QFontInfo(QtGui.QFont(family)).family()
        if resolved and resolved != app.font().family():
            logger.warning(
                "Using fallback font '%s' for missing '%s'", resolved, family
            )
            return family
    for fallback in ("Exo 2", "Arial"):
        if fallback in QtGui.QFontDatabase.families():
            logger.warning(
                "Unable to load font '%s'; falling back to '%s'", family, fallback
            )
            return fallback
    fallback = app.font().family() if app is not None else "Arial"
    logger.warning(
        "Unable to load font '%s'; using application font '%s'", family, fallback
    )
    return fallback

def resolve_font_config(parent: QtWidgets.QWidget | None = None) -> tuple[str, str]:
    """Resolve configured font families and persist any changes.

    Separate fonts are stored for headers, general text and the sidebar.  Each
    entry is validated via :func:`ensure_font_registered` so that missing font
    files gracefully fall back to bundled defaults.  Any adjustments are written
    back to the configuration file.  Only the header and text families are
    returned for backwards compatibility with existing call sites.
    """

    default = CONFIG.get("font_family", "Exo 2")
    header = ensure_font_registered(CONFIG.get("header_font", default), parent)
    text = ensure_font_registered(CONFIG.get("text_font", default), parent)
    sidebar = ensure_font_registered(
        CONFIG.get("sidebar_font", CONFIG.get("header_font", default)), parent
    )

    changed = False
    if header != CONFIG.get("header_font"):
        CONFIG["header_font"] = header
        changed = True
    if text != CONFIG.get("text_font"):
        CONFIG["text_font"] = text
        changed = True
    if sidebar != CONFIG.get("sidebar_font"):
        CONFIG["sidebar_font"] = sidebar
        changed = True

    if changed:
        os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(CONFIG, f, ensure_ascii=False, indent=2)

    return header, text


@dataclass
class MonthData:
    year: int
    month: int
    days: Dict[int, List[Dict[str, str]]] = field(default_factory=dict)

    @property
    def path(self) -> str:
        os.makedirs(DATA_DIR, exist_ok=True)
        return os.path.join(DATA_DIR, f"{self.year:04d}-{self.month:02d}.json")

    def save(self) -> None:
        days: Dict[str, List[Dict[str, str]]] = {}
        for day, rows in self.days.items():
            row_list: List[Dict[str, str]] = []
            for r in rows:
                row_list.append({
                    "work": r.get("work", ""),
                    "plan": r.get("plan", ""),
                    "done": r.get("done", ""),
                })
            if row_list:
                days[str(day)] = row_list
        data = {"year": self.year, "month": self.month, "days": days}
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls, year: int, month: int) -> "MonthData":
        path = os.path.join(DATA_DIR, f"{year:04d}-{month:02d}.json")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            days: Dict[int, List[Dict[str, str]]] = {}
            for k, v in data.get("days", {}).items():
                row_list: List[Dict[str, str]] = []
                for row in v:
                    if isinstance(row, dict):
                        row_list.append({
                            "work": row.get("work", ""),
                            "plan": row.get("plan", ""),
                            "done": row.get("done", ""),
                        })
                    elif isinstance(row, list):
                        row_list.append({
                            "work": row[0] if len(row) > 0 else "",
                            "plan": row[1] if len(row) > 1 else "",
                            "done": row[2] if len(row) > 2 else "",
                        })
                days[int(k)] = row_list
            return cls(year=data.get("year", year), month=data.get("month", month), days=days)
        return cls(year=year, month=month)


class ReleaseDialog(QtWidgets.QDialog):
    """Диалог для управления выкладкой.

    Структура представлена таблицей из четырёх колонок:
    день, работа, количество глав и время. Допускается несколько
    записей на один день."""

    def __init__(self, year, month, works, parent=None):
        super().__init__(parent)
        self.year = year
        self.month = month
        self.works = list(works)
        self.setWindowTitle("Выкладка")
        self.resize(600, 400)

        lay = QtWidgets.QVBoxLayout(self)
        self.days_in_month = calendar.monthrange(year, month)[1]

        self.table = QtWidgets.QTableWidget(0, 4, self)
        self.table.setHorizontalHeaderLabels(["День", "Работа", "Глав", "Время"])
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table.setStyleSheet(
            "QTableWidget{border:1px solid #555; border-radius:8px;} "
            "QHeaderView::section{padding:0 6px;}"
        )
        self.table.setRowCount(31)
        for row in range(31):
            item = QtWidgets.QTableWidgetItem(str(row + 1))
            item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.table.setItem(row, 0, item)

        app = QtWidgets.QApplication.instance()
        self.setFont(app.font())
        self.table.setFont(app.font())
        header_font = QtGui.QFont(CONFIG.get("header_font"))
        self.table.horizontalHeader().setFont(header_font)

        self.table.setAttribute(QtCore.Qt.WA_Hover, True)
        self.table.viewport().setAttribute(QtCore.Qt.WA_Hover, True)
        self._tbl_filter = None
        if neon_enabled():
            self._tbl_filter = NeonEventFilter(self.table)
            self.table.installEventFilter(self._tbl_filter)
            self.table.viewport().installEventFilter(self._tbl_filter)
            self.table._neon_filter = self._tbl_filter

        lay.addWidget(self.table)

        box = QtWidgets.QDialogButtonBox(self)
        btn_save = StyledPushButton("Сохранить", self)
        btn_save.setIcon(icon("save"))
        btn_save.setIconSize(QtCore.QSize(20, 20))
        btn_close = StyledPushButton("Закрыть", self)
        btn_close.setIcon(icon("x"))
        btn_close.setIconSize(QtCore.QSize(20, 20))
        box.addButton(btn_save, QtWidgets.QDialogButtonBox.AcceptRole)
        box.addButton(btn_close, QtWidgets.QDialogButtonBox.RejectRole)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        lay.addWidget(box)

        for b in (btn_save, btn_close):
            b.setFixedSize(b.sizeHint())

        self._settings = QtCore.QSettings("rabota2", "rabota2")
        geom = self._settings.value("ReleaseDialog/geometry", type=QtCore.QByteArray)
        if geom is not None:
            self.restoreGeometry(geom)
        sizes = self._settings.value("ReleaseDialog/columns", type=list)
        for i, w in enumerate(sizes or []):
            try:
                self.table.setColumnWidth(i, int(w))
            except (TypeError, ValueError):
                pass  # пропустить некорректное значение

        self.load()

    def closeEvent(self, event):
        self.save()
        self._settings.setValue("ReleaseDialog/geometry", self.saveGeometry())
        cols = [int(self.table.columnWidth(i)) for i in range(self.table.columnCount())]
        self._settings.setValue("ReleaseDialog/columns", cols)
        self._settings.sync()
        super().closeEvent(event)

    def file_path(self):
        return os.path.join(release_dir(self.year), f"{self.month:02d}.json")

    def load(self):
        path = self.file_path()
        data = {}
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except json.JSONDecodeError as exc:
                logger.error("Failed to parse release data from '%s': %s", path, exc)
                QtWidgets.QMessageBox.warning(
                    self,
                    "Ошибка",
                    "Данные повреждены или нечитаемы.",
                )
                data = {}

        for day_str, entries in data.get("days", {}).items():
            day = int(day_str)
            if 1 <= day <= self.days_in_month and entries:
                entry = entries[0]
                self.table.setItem(day - 1, 1, QtWidgets.QTableWidgetItem(entry.get("work", "")))
                self.table.setItem(day - 1, 2, QtWidgets.QTableWidgetItem(str(entry.get("chapters", ""))))
                self.table.setItem(day - 1, 3, QtWidgets.QTableWidgetItem(entry.get("time", "")))

    def save(self):
        days: Dict[str, List[Dict[str, str | int]]] = {}
        for row in range(self.days_in_month):
            work_item = self.table.item(row, 1)
            chapters_item = self.table.item(row, 2)
            time_item = self.table.item(row, 3)

            work_name = work_item.text().strip() if work_item else ""
            if not work_name:
                continue
            chapters_text = chapters_item.text().strip() if chapters_item else ""
            try:
                chapters = int(chapters_text) if chapters_text else 0
            except (TypeError, ValueError):
                chapters = 0
            time_text = time_item.text().strip() if time_item else ""
            day = row + 1
            entry = {
                "work": work_name,
                "chapters": chapters,
                "time": time_text,
            }
            days[str(day)] = [entry]

        works = sorted({e["work"] for entries in days.values() for e in entries})
        data = {"works": works, "days": days}

        try:
            os.makedirs(os.path.dirname(self.file_path()), exist_ok=True)
            with open(self.file_path(), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except OSError as exc:
            logger.warning("Failed to save release data: %s", exc)


class StatsEntryForm(QtWidgets.QWidget):
    """Форма для ввода/редактирования статистики месяца."""

    # ключ, заголовок, класс виджета
    INPUT_FIELDS = [
        ("work", "Работа", QtWidgets.QLineEdit),
        ("status", "Статус", QtWidgets.QLineEdit),
        ("adult", "18+", QtWidgets.QCheckBox),
        ("total_chapters", "Всего глав", QtWidgets.QSpinBox),
        ("chars_per_chapter", "Знаков глава", QtWidgets.QSpinBox),
        ("planned", "Запланировано", QtWidgets.QSpinBox),
        ("chapters", "Сделано глав", QtWidgets.QSpinBox),
        ("progress", "Прогресс перевода", QtWidgets.QDoubleSpinBox),
        ("release", "Выпуск", QtWidgets.QLineEdit),
        ("profit", "Профит", QtWidgets.QDoubleSpinBox),
        ("ads", "Затраты на рекламу", QtWidgets.QDoubleSpinBox),
        ("views", "Просмотры", QtWidgets.QSpinBox),
        ("likes", "Лайки", QtWidgets.QSpinBox),
        ("thanks", "Спасибо", QtWidgets.QSpinBox),
    ]

    # колонки таблицы: все поля + вычисляемое "chars"
    TABLE_COLUMNS = [(key, label) for key, label, _ in INPUT_FIELDS] + [
        ("chars", "Знаков"),
    ]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSizePolicy(
            QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding
        )
        form = QtWidgets.QFormLayout(self)
        form.setFieldGrowthPolicy(QtWidgets.QFormLayout.AllNonFixedFieldsGrow)
        self.widgets = {}
        for key, label, cls in self.INPUT_FIELDS:
            w = cls(self)
            if isinstance(w, QtWidgets.QSpinBox):
                w.setRange(0, 1_000_000_000)
                w.setFixedWidth(w.sizeHint().width() + 20)
            else:
                w.setMinimumWidth(w.sizeHint().width())
            if isinstance(w, QtWidgets.QDoubleSpinBox):
                w.setRange(0, 1_000_000_000)
                w.setDecimals(2)
            if key == "progress" and isinstance(w, QtWidgets.QDoubleSpinBox):
                w.setRange(0, 100)
                w.setSuffix("%")
            if isinstance(w, QtWidgets.QComboBox):
                w.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToContents)
            form.addRow(label, w)
            self.widgets[key] = w
            w.setAttribute(QtCore.Qt.WA_Hover, True)
            if key != "adult":  # avoid framing the entire row for checkbox
                if neon_enabled():
                    filt = NeonEventFilter(w)
                    w.installEventFilter(filt)
                    w._neon_filter = filt

    def get_record(self) -> Dict[str, int | float | str | bool]:
        record: Dict[str, int | float | str | bool] = {}
        for key, _, _ in self.INPUT_FIELDS:
            w = self.widgets[key]
            if isinstance(w, QtWidgets.QLineEdit):
                record[key] = w.text().strip()
            elif isinstance(w, QtWidgets.QCheckBox):
                record[key] = w.isChecked()
            else:  # spinboxes
                record[key] = w.value()
        record["chars"] = record["chapters"] * record["chars_per_chapter"]
        return record

    def set_record(self, record: Dict[str, int | float | str | bool]):
        for key, _, _ in self.INPUT_FIELDS:
            w = self.widgets[key]
            val = record.get(key)
            if isinstance(w, QtWidgets.QLineEdit):
                w.setText(str(val) if val is not None else "")
            elif isinstance(w, QtWidgets.QCheckBox):
                w.setChecked(bool(val))
            else:
                w.setValue(val if isinstance(val, (int, float)) else 0)

    def clear(self):
        for key, _, _ in self.INPUT_FIELDS:
            w = self.widgets[key]
            if isinstance(w, QtWidgets.QLineEdit):
                w.clear()
            elif isinstance(w, QtWidgets.QCheckBox):
                w.setChecked(False)
            else:
                w.setValue(0)


class StatsDialog(QtWidgets.QDialog):
    """Диалог для просмотра и редактирования статистики месяца."""

    def __init__(self, year: int, month: int, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Вводные данные")
        self.resize(800, 500)

        lay = QtWidgets.QVBoxLayout(self)
        self.table_stats = NeonTableWidget(0, len(StatsEntryForm.TABLE_COLUMNS), self)
        self.table_stats.setSelectionBehavior(QtWidgets.QTableView.SelectColumns)
        self.table_stats.setSelectionMode(QtWidgets.QTableView.SingleSelection)
        self.table_stats.setStyleSheet(
            "QTableWidget{border:1px solid #555; border-radius:8px;} "
            "QTableWidget::item{border:0;} "
            "QHeaderView::section{padding:0 8px;}"
        )
        header = self.table_stats.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        header.setTextElideMode(QtCore.Qt.ElideNone)
        self.table_stats.setHorizontalHeaderLabels(
            [h for _, h in StatsEntryForm.TABLE_COLUMNS]
        )
        self.table_stats.setSortingEnabled(True)
        self.table_stats.verticalHeader().setVisible(False)
        self.table_stats.verticalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.Stretch
        )
        self.table_stats.itemSelectionChanged.connect(self.on_table_selection)
        lay.addWidget(self.table_stats)

        self.form_stats = StatsEntryForm(self)
        lay.addWidget(self.form_stats)

        self.btn_box = QtWidgets.QDialogButtonBox(self)
        btn_save = StyledPushButton("Сохранить", self)
        btn_save.setIcon(icon("save"))
        btn_save.setIconSize(QtCore.QSize(20, 20))
        btn_close = StyledPushButton("Закрыть", self)
        btn_close.setIcon(icon("x"))
        btn_close.setIconSize(QtCore.QSize(20, 20))
        for btn in (btn_save, btn_close):
            btn.setFixedSize(btn.sizeHint())
            btn.setStyleSheet(btn.styleSheet() + "border:1px solid transparent;")
        self.btn_box.addButton(btn_save, QtWidgets.QDialogButtonBox.AcceptRole)
        self.btn_box.addButton(btn_close, QtWidgets.QDialogButtonBox.RejectRole)
        self.btn_box.accepted.connect(self.save_record)
        self.btn_box.rejected.connect(self.reject)
        lay.addWidget(self.btn_box)
        lay.setStretch(0, 2)
        lay.setStretch(1, 1)

        self.records: List[Dict[str, int | float | str | bool]] = []
        self.current_index = None
        self.year = year
        self.month = month
        self._settings = QtCore.QSettings("rabota2", "rabota2")
        geom = self._settings.value("StatsDialog/geometry", type=QtCore.QByteArray)
        if geom is not None:
            self.restoreGeometry(geom)
        self.load_stats(year, month)
        sizes = self._settings.value("StatsDialog/columns", type=list)
        for i, w in enumerate(sizes or []):
            try:
                self.table_stats.setColumnWidth(i, int(w))
            except (TypeError, ValueError):
                pass  # пропустить некорректное значение

    def resizeEvent(self, event):
        self.table_stats.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.Interactive
        )
        super().resizeEvent(event)

    def on_table_selection(self):
        sel = self.table_stats.selectionModel().selectedRows()
        if sel:
            row = sel[0].row()
            self.current_index = row
            self.form_stats.set_record(self.records[row])
        else:
            self.current_index = None
            self.form_stats.clear()

    def load_stats(self, year: int, month: int):
        self.year = year
        self.month = month
        path = os.path.join(stats_dir(year), f"{year}.json")
        data = {}
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except json.JSONDecodeError as exc:
                logger.error("Failed to parse stats data from '%s': %s", path, exc)
                QtWidgets.QMessageBox.warning(
                    self,
                    "Ошибка",
                    "Данные повреждены или нечитаемы.",
                )
                data = {}
        self.records = data.get(str(month), [])
        self.table_stats.setRowCount(len(self.records))
        for r, rec in enumerate(self.records):
            for c, (key, _) in enumerate(StatsEntryForm.TABLE_COLUMNS):
                val = rec.get(key, "")
                if isinstance(val, bool):
                    text = "✓" if val else ""
                    item = QtWidgets.QTableWidgetItem(text)
                    item.setData(QtCore.Qt.EditRole, int(val))
                elif isinstance(val, (int, float)):
                    item = QtWidgets.QTableWidgetItem()
                    item.setData(QtCore.Qt.EditRole, val)
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                else:
                    item = QtWidgets.QTableWidgetItem(str(val))
                self.table_stats.setItem(r, c, item)
        self.table_stats.resizeColumnsToContents()
        header = self.table_stats.horizontalHeader()
        header.setMinimumSectionSize(50)
        header.setTextElideMode(QtCore.Qt.ElideNone)
        self.table_stats.resizeRowsToContents()
        header.setStretchLastSection(True)
        total_width = sum(header.sectionSize(i) for i in range(header.count()))
        if total_width <= self.table_stats.viewport().width():
            self.table_stats.setHorizontalScrollBarPolicy(
                QtCore.Qt.ScrollBarAlwaysOff
            )
        else:
            self.table_stats.setHorizontalScrollBarPolicy(
                QtCore.Qt.ScrollBarAsNeeded
            )
        self.current_index = None
        self.form_stats.clear()

    def save_record(self):
        record = self.form_stats.get_record()
        if self.current_index is None:
            self.records.append(record)
        else:
            self.records[self.current_index] = record
        path = os.path.join(stats_dir(self.year), f"{self.year}.json")
        data = {}
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        data[str(self.month)] = self.records
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.load_stats(self.year, self.month)

    def closeEvent(self, event):
        if self.current_index is not None or any(
            self.form_stats.get_record().values()
        ):
            self.save_record()
        self._settings.setValue("StatsDialog/geometry", self.saveGeometry())
        cols = [
            int(self.table_stats.columnWidth(i))
            for i in range(self.table_stats.columnCount())
        ]
        self._settings.setValue("StatsDialog/columns", cols)
        self._settings.sync()
        super().closeEvent(event)


class AnalyticsDialog(QtWidgets.QDialog):
    """Годовая статистика: месяцы × показатели с колонкой "Итого за год"."""

    INDICATORS = [
        "Работ", "Завершенных", "Онгоингов", "Глав", "Знаков",
        "Просмотров", "Профит", "РК", "Чистыми", "Лайков", "Спасибо",
        "Камса", "Потрачено на софт",
    ]

    def __init__(self, year, parent=None):
        super().__init__(parent)
        self.year = year
        self.setWindowTitle("Аналитика")
        self.resize(900, 400)

        lay = QtWidgets.QVBoxLayout(self)
        top = QtWidgets.QHBoxLayout()
        top.addWidget(QtWidgets.QLabel("Год:"))
        self.spin_year = QtWidgets.QSpinBox(self)
        self.spin_year.setRange(2000, 2100)
        self.spin_year.setValue(year)
        self.spin_year.setFixedWidth(self.spin_year.sizeHint().width())
        self.spin_year.setButtonSymbols(QtWidgets.QAbstractSpinBox.UpDownArrows)
        self.spin_year.valueChanged.connect(self._year_changed)
        self.spin_year.setAttribute(QtCore.Qt.WA_Hover, True)
        top.addWidget(self.spin_year)
        if neon_enabled():
            filt = NeonEventFilter(self.spin_year)
            self.spin_year.installEventFilter(filt)
            self.spin_year._neon_filter = filt
        top.addStretch(1)
        lay.addLayout(top)

        cols = len(RU_MONTHS) + 1
        self.table = NeonTableWidget(len(self.INDICATORS), cols, self)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectItems)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.setStyleSheet(
            "QTableWidget{border:1px solid #555; border-radius:8px;} "
            "QTableWidget::item{border:0;} "
            "QHeaderView::section{padding:0 6px;}"
        )
        self.table.setHorizontalHeaderLabels(RU_MONTHS + ["Итого за год"])
        self.table.setVerticalHeaderLabels(self.INDICATORS)
        self.table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.Interactive
        )
        self.table.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        lay.addWidget(self.table, 1)

        box = QtWidgets.QDialogButtonBox(self)
        btn_save = StyledPushButton("Сохранить", self)
        btn_save.setIcon(icon("save"))
        btn_save.setIconSize(QtCore.QSize(20, 20))
        btn_close = StyledPushButton("Закрыть", self)
        btn_close.setIcon(icon("x"))
        btn_close.setIconSize(QtCore.QSize(20, 20))
        for btn in (btn_save, btn_close):
            btn.setFixedSize(btn.sizeHint())
            btn.setStyleSheet(btn.styleSheet() + "border:1px solid transparent;")
        box.addButton(btn_save, QtWidgets.QDialogButtonBox.AcceptRole)
        box.addButton(btn_close, QtWidgets.QDialogButtonBox.RejectRole)
        box.accepted.connect(self.save)
        box.rejected.connect(self.reject)
        lay.addWidget(box)

        # prepare items
        for r, name in enumerate(self.INDICATORS):
            for c in range(cols):
                it = QtWidgets.QTableWidgetItem("0")
                it.setTextAlignment(QtCore.Qt.AlignCenter)
                if name not in ("Камса", "Потрачено на софт") or c == cols - 1:
                    it.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                self.table.setItem(r, c, it)

        self.table.itemChanged.connect(self._item_changed)

        self._loading = False
        self._commissions = {str(m): 0.0 for m in range(1, 13)}
        self._software = {str(m): 0.0 for m in range(1, 13)}
        self._net = {str(m): 0.0 for m in range(1, 13)}
        self._settings = QtCore.QSettings("rabota2", "rabota2")
        geom = self._settings.value("AnalyticsDialog/geometry", type=QtCore.QByteArray)
        if geom is not None:
            self.restoreGeometry(geom)
        self.load(year)
        sizes = self._settings.value("AnalyticsDialog/columns", type=list)
        for i, w in enumerate(sizes or []):
            try:
                self.table.setColumnWidth(i, int(w))
            except (TypeError, ValueError):
                pass  # пропустить некорректное значение

    def resizeEvent(self, event):
        self.table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.Interactive
        )
        self.table.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        super().resizeEvent(event)

    def _year_changed(self, val):
        self.load(val)

    # --- data handling -------------------------------------------------
    def load(self, year):
        self._loading = True
        self.year = year
        self.spin_year.setValue(year)

        # load manual values
        self._commissions = {str(m): 0.0 for m in range(1, 13)}
        self._software = {str(m): 0.0 for m in range(1, 13)}
        self._net = {str(m): 0.0 for m in range(1, 13)}
        path = os.path.join(year_dir(year), f"{year}.json")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self._commissions.update({str(k): float(v) for k, v in data.get("commission", {}).items()})
            self._software.update({str(k): float(v) for k, v in data.get("software", {}).items()})
            self._net.update({str(k): float(v) for k, v in data.get("net", {}).items()})

        # fill table with monthly values
        for m in range(1, 13):
            stats = self._calc_month_stats(year, m)
            for ind, val in stats.items():
                row = self.INDICATORS.index(ind)
                self.table.item(row, m - 1).setText(str(val))
            self.table.item(self.INDICATORS.index("Камса"), m - 1).setText(str(self._commissions[str(m)]))
            self.table.item(self.INDICATORS.index("Потрачено на софт"), m - 1).setText(str(self._software[str(m)]))
            self.table.item(self.INDICATORS.index("Чистыми"), m - 1).setText(
                str(self._net.get(str(m), stats.get("Чистыми", 0)))
            )

        self._recalculate()
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setMinimumSectionSize(50)
        self.table.resizeRowsToContents()
        self.table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.Interactive
        )
        self.table.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self._loading = False

    def save(self, accept=True):
        path = os.path.join(year_dir(self.year), f"{self.year}.json")
        data = {
            "commission": self._commissions,
            "software": self._software,
            "net": self._net,
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        if accept:
            self.accept()

    def closeEvent(self, event):
        self.save(accept=False)
        self._settings.setValue("AnalyticsDialog/geometry", self.saveGeometry())
        cols = [int(self.table.columnWidth(i)) for i in range(self.table.columnCount())]
        self._settings.setValue("AnalyticsDialog/columns", cols)
        self._settings.sync()
        super().closeEvent(event)

    # --- helpers -------------------------------------------------------
    def _calc_month_stats(self, year, month):
        res = {k: 0 for k in self.INDICATORS if k not in ("Камса", "Потрачено на софт")}
        path = os.path.join(stats_dir(year), f"{year}.json")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            month_data = data.get(str(month), [])
            res["Работ"] = len(month_data)
            for rec in month_data:
                status = (rec.get("status", "") or "").lower()
                if "заверш" in status:
                    res["Завершенных"] += 1
                elif "онго" in status:
                    res["Онгоингов"] += 1
                res["Глав"] += int(rec.get("chapters", 0) or 0)
                res["Знаков"] += int(rec.get("chars", 0) or 0)
                res["Просмотров"] += int(rec.get("views", 0) or 0)
                res["Профит"] += float(rec.get("profit", 0) or 0)
                res["РК"] += float(rec.get("ads", 0) or 0)
                res["Лайков"] += int(rec.get("likes", 0) or 0)
                res["Спасибо"] += int(rec.get("thanks", 0) or 0)
        res["Чистыми"] = round(
            res["Профит"] - res["РК"] - self._software.get(str(month), 0.0), 2
        )
        return res

    def _item_changed(self, item):
        if self._loading:
            return
        row = item.row()
        col = item.column()
        ind = self.INDICATORS[row]
        if ind == "Камса":
            try:
                self._commissions[str(col + 1)] = float(item.text())
            except ValueError:
                self._commissions[str(col + 1)] = 0.0
        elif ind == "Потрачено на софт":
            try:
                self._software[str(col + 1)] = float(item.text())
            except ValueError:
                self._software[str(col + 1)] = 0.0
        self._recalculate()

    def _recalculate(self):
        cols = len(RU_MONTHS)
        r_profit = self.INDICATORS.index("Профит")
        r_rk = self.INDICATORS.index("РК")
        r_soft = self.INDICATORS.index("Потрачено на софт")
        r_net = self.INDICATORS.index("Чистыми")
        # recompute net values
        for c in range(cols):
            try:
                profit = float(self.table.item(r_profit, c).text())
            except ValueError:
                profit = 0.0
            try:
                rk = float(self.table.item(r_rk, c).text())
            except ValueError:
                rk = 0.0
            try:
                soft = float(self.table.item(r_soft, c).text())
            except ValueError:
                soft = 0.0
            net = round(profit - rk - soft, 2)
            self.table.item(r_net, c).setText(str(net))
            self._net[str(c + 1)] = net

        # totals
        for r in range(len(self.INDICATORS)):
            total = 0.0
            for c in range(cols):
                try:
                    total += float(self.table.item(r, c).text())
                except ValueError:
                    pass
            item = self.table.item(r, cols)
            item.setText(str(round(total, 2)))
            font = item.font()
            font.setBold(True)
            item.setFont(font)

class TopDialog(QtWidgets.QDialog):
    """Агрегирование и сохранение топов за период."""

    def __init__(self, year, parent=None):
        super().__init__(parent)
        self.year = year
        self.results = []
        self.setWindowTitle("Топы")
        self.resize(700, 400)

        lay = QtWidgets.QVBoxLayout(self)
        top = QtWidgets.QFormLayout()

        self.spin_year = QtWidgets.QSpinBox(self)
        self.spin_year.setRange(2000, 2100)
        self.spin_year.setValue(year)
        self.spin_year.setStyleSheet("padding:0 4px;")
        self.spin_year.setAttribute(QtCore.Qt.WA_Hover, True)
        if neon_enabled():
            filt = NeonEventFilter(self.spin_year)
            self.spin_year.installEventFilter(filt)
            self.spin_year._neon_filter = filt

        self.combo_mode = QtWidgets.QComboBox(self)
        self.combo_mode.addItem("Месяц", "month")
        self.combo_mode.addItem("Квартал", "quarter")
        self.combo_mode.addItem("Полугодие", "half")
        self.combo_mode.addItem("Год", "year")
        self.combo_mode.currentIndexChanged.connect(self._mode_changed)
        self.combo_mode.setAttribute(QtCore.Qt.WA_Hover, True)
        if neon_enabled():
            filt = NeonEventFilter(self.combo_mode)
            self.combo_mode.installEventFilter(filt)
            self.combo_mode._neon_filter = filt

        self.combo_period = QtWidgets.QComboBox(self)
        self.combo_period.setAttribute(QtCore.Qt.WA_Hover, True)
        if neon_enabled():
            filt = NeonEventFilter(self.combo_period)
            self.combo_period.installEventFilter(filt)
            self.combo_period._neon_filter = filt
        self._mode_changed()  # fill periods

        self.btn_calc = StyledPushButton("Сформировать", self)
        self.btn_calc.clicked.connect(self.calculate)

        row = QtWidgets.QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.addWidget(self.spin_year)
        row.addWidget(self.combo_mode)
        row.addWidget(self.combo_period)
        row.addWidget(self.btn_calc)

        top.addRow(QtWidgets.QLabel("Год:"), row)
        lay.addLayout(top)

        headers = [
            "№",
            "Работа",
            "Статус",
            "Всего глав",
            "Запланированно",
            "Сделано глав",
            "Прогресс перевода",
            "Выпуск",
            "Знаков",
            "Просмотров",
            "Профит",
            "РК",
            "Лайков",
            "Спасибо",
        ]
        self.table = NeonTableWidget(0, len(headers), self)
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setStyleSheet(
            "QTableWidget{border:1px solid #555; border-radius:8px;} "
            "QTableWidget::item{border:0;} "
            "QHeaderView::section{padding:0 8px;}"
        )
        self.table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.Interactive
        )
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.Stretch
        )
        self.table.setSortingEnabled(True)

        app = QtWidgets.QApplication.instance()
        self.setFont(app.font())
        self.table.setFont(app.font())
        header_font = QtGui.QFont(CONFIG.get("header_font"))
        self.table.horizontalHeader().setFont(header_font)

        lay.addWidget(self.table, 1)

        box = QtWidgets.QDialogButtonBox(self)
        btn_save = StyledPushButton("Сохранить", self)
        btn_save.setIcon(icon("save"))
        btn_save.setIconSize(QtCore.QSize(20, 20))
        btn_close = StyledPushButton("Закрыть", self)
        btn_close.setIcon(icon("x"))
        btn_close.setIconSize(QtCore.QSize(20, 20))
        box.addButton(btn_save, QtWidgets.QDialogButtonBox.AcceptRole)
        box.addButton(btn_close, QtWidgets.QDialogButtonBox.RejectRole)
        box.accepted.connect(self._save_and_accept)
        box.rejected.connect(self.reject)
        lay.addWidget(box)
        for b in (btn_save, btn_close):
            b.setFixedSize(b.sizeHint())

        self._settings = QtCore.QSettings("rabota2", "rabota2")
        geom = self._settings.value("TopDialog/geometry", type=QtCore.QByteArray)
        if geom is not None:
            self.restoreGeometry(geom)
        self.calculate()
        sizes = self._settings.value("TopDialog/columns", type=list)
        for i, w in enumerate(sizes or []):
            try:
                self.table.setColumnWidth(i, int(w))
            except (TypeError, ValueError):
                pass  # пропустить некорректное значение

    def resizeEvent(self, event):
        self.table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.Interactive
        )
        super().resizeEvent(event)

    def _mode_changed(self):
        mode = self.combo_mode.currentData()
        self.combo_period.clear()
        if mode == "month":
            for i, m in enumerate(RU_MONTHS, 1):
                self.combo_period.addItem(m, i)
            self.combo_period.setEnabled(True)
        elif mode == "quarter":
            for i in range(1, 5):
                self.combo_period.addItem(str(i), i)
            self.combo_period.setEnabled(True)
        elif mode == "half":
            for i in range(1, 3):
                self.combo_period.addItem(str(i), i)
            self.combo_period.setEnabled(True)
        else:
            self.combo_period.setEnabled(False)

    # --- helpers -------------------------------------------------------
    def _months_for_period(self):
        mode = self.combo_mode.currentData()
        if mode == "month" and self.combo_period.isEnabled():
            return [self.combo_period.currentData()]
        if mode == "quarter" and self.combo_period.isEnabled():
            q = self.combo_period.currentData()
            start = (q - 1) * 3 + 1
            return list(range(start, start + 3))
        if mode == "half" and self.combo_period.isEnabled():
            h = self.combo_period.currentData()
            start = (h - 1) * 6 + 1
            return list(range(start, start + 6))
        return list(range(1, 13))

    def calculate(self):
        year = self.spin_year.value()
        months = self._months_for_period()
        path = os.path.join(stats_dir(year), f"{year}.json")
        totals = {}
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            for m in months:
                for rec in data.get(str(m), []):
                    work = rec.get("work", "")
                    t = totals.setdefault(
                        work,
                        {
                            "status": "",
                            "total_chapters": 0,
                            "planned": 0,
                            "chapters": 0,
                            "progress": 0.0,
                            "release": "",
                            "chars": 0,
                            "views": 0,
                            "profit": 0.0,
                            "ads": 0.0,
                            "likes": 0,
                            "thanks": 0,
                            "done": 0,
                        },
                    )
                    t["status"] = rec.get("status", t["status"])
                    t["total_chapters"] = max(
                        t["total_chapters"], int(rec.get("total_chapters", 0) or 0)
                    )
                    t["planned"] += int(rec.get("planned", 0) or 0)
                    t["chapters"] += int(rec.get("chapters", 0) or 0)
                    prog = rec.get("progress")
                    if prog is not None:
                        t["progress"] = prog
                    rel = rec.get("release")
                    if rel:
                        t["release"] = rel
                    t["chars"] += int(rec.get("chars", 0) or 0)
                    t["views"] += int(rec.get("views", 0) or 0)
                    t["profit"] += float(rec.get("profit", 0) or 0)
                    t["ads"] += float(rec.get("ads", 0) or 0)
                    t["likes"] += int(rec.get("likes", 0) or 0)
                    t["thanks"] += int(rec.get("thanks", 0) or 0)
                    status = (rec.get("status", "") or "").lower()
                    if "заверш" in status:
                        t["done"] += 1
        results = sorted(totals.items(), key=lambda kv: kv[0])
        self.results = results
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        sums = {
            "total_chapters": 0,
            "planned": 0,
            "chapters": 0,
            "chars": 0,
            "views": 0,
            "profit": 0.0,
            "ads": 0.0,
            "likes": 0,
            "thanks": 0,
        }
        for idx, (work, vals) in enumerate(results, 1):
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(str(idx)))
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(work))
            self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(vals.get("status", "")))
            self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(str(vals["total_chapters"])))
            self.table.setItem(row, 4, QtWidgets.QTableWidgetItem(str(vals["planned"])))
            self.table.setItem(row, 5, QtWidgets.QTableWidgetItem(str(vals["chapters"])))
            self.table.setItem(
                row, 6, QtWidgets.QTableWidgetItem(f"{round(vals['progress'], 2)}%")
            )
            self.table.setItem(row, 7, QtWidgets.QTableWidgetItem(vals.get("release", "")))
            self.table.setItem(row, 8, QtWidgets.QTableWidgetItem(str(vals["chars"])))
            self.table.setItem(row, 9, QtWidgets.QTableWidgetItem(str(vals["views"])))
            self.table.setItem(row, 10, QtWidgets.QTableWidgetItem(str(round(vals["profit"], 2))))
            self.table.setItem(row, 11, QtWidgets.QTableWidgetItem(str(round(vals["ads"], 2))))
            self.table.setItem(row, 12, QtWidgets.QTableWidgetItem(str(vals["likes"])))
            self.table.setItem(row, 13, QtWidgets.QTableWidgetItem(str(vals["thanks"])))

            # accumulate sums
            sums["total_chapters"] += vals["total_chapters"]
            sums["planned"] += vals["planned"]
            sums["chapters"] += vals["chapters"]
            sums["chars"] += vals["chars"]
            sums["views"] += vals["views"]
            sums["profit"] += vals["profit"]
            sums["ads"] += vals["ads"]
            sums["likes"] += vals["likes"]
            sums["thanks"] += vals["thanks"]

        if results:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem("Итого"))
            self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(str(sums["total_chapters"])))
            self.table.setItem(row, 4, QtWidgets.QTableWidgetItem(str(sums["planned"])))
            self.table.setItem(row, 5, QtWidgets.QTableWidgetItem(str(sums["chapters"])))
            self.table.setItem(row, 8, QtWidgets.QTableWidgetItem(str(sums["chars"])))
            self.table.setItem(row, 9, QtWidgets.QTableWidgetItem(str(sums["views"])))
            self.table.setItem(row, 10, QtWidgets.QTableWidgetItem(str(round(sums["profit"], 2))))
            self.table.setItem(row, 11, QtWidgets.QTableWidgetItem(str(round(sums["ads"], 2))))
            self.table.setItem(row, 12, QtWidgets.QTableWidgetItem(str(sums["likes"])))
            self.table.setItem(row, 13, QtWidgets.QTableWidgetItem(str(sums["thanks"])))
        self.table.setSortingEnabled(True)
        self.table.resizeColumnsToContents()
        header = self.table.horizontalHeader()
        header.setMinimumSectionSize(50)
        header.setTextElideMode(QtCore.Qt.ElideNone)

    def _period_key(self):
        mode = self.combo_mode.currentData()
        if mode == "month" and self.combo_period.isEnabled():
            return f"M{self.combo_period.currentData():02d}"
        if mode == "quarter" and self.combo_period.isEnabled():
            return f"Q{self.combo_period.currentData()}"
        if mode == "half" and self.combo_period.isEnabled():
            return f"H{self.combo_period.currentData()}"
        return "Y"

    def save(self):
        if not self.results:
            self.calculate()
        year = self.spin_year.value()
        path = os.path.join(top_dir(year), f"{year}.json")
        data = {}
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        key = self._period_key()
        results = []
        for w, vals in self.results:
            results.append(
                {
                    "work": w,
                    "status": vals.get("status", ""),
                    "total_chapters": vals.get("total_chapters", 0),
                    "planned": vals.get("planned", 0),
                    "chapters": vals.get("chapters", 0),
                    "progress": vals.get("progress", 0.0),
                    "release": vals.get("release", ""),
                    "chars": vals.get("chars", 0),
                    "views": vals.get("views", 0),
                    "profit": vals.get("profit", 0.0),
                    "ads": vals.get("ads", 0.0),
                    "likes": vals.get("likes", 0),
                    "thanks": vals.get("thanks", 0),
                }
            )
        data[key] = {"results": results}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def _save_and_accept(self):
        self.save()
        self.accept()

    def closeEvent(self, event):
        self.save()
        self._settings.setValue("TopDialog/geometry", self.saveGeometry())
        cols = [int(self.table.columnWidth(i)) for i in range(self.table.columnCount())]
        self._settings.setValue("TopDialog/columns", cols)
        self._settings.sync()
        super().closeEvent(event)

class NeonTableWidget(QtWidgets.QTableWidget):
    """Вложенная таблица с neonовым подсвечиванием при наведении и фокусе."""

    def __init__(self, rows, cols, parent=None, use_neon=True):
        super().__init__(rows, cols, parent)
        self.setFocusPolicy(QtCore.Qt.StrongFocus)
        self.setAttribute(QtCore.Qt.WA_Hover, True)
        self.viewport().setAttribute(QtCore.Qt.WA_Hover, True)
        self._neon_filter = None
        self._active_editor: QtWidgets.QLineEdit | None = None
        if use_neon and neon_enabled():
            self._neon_filter = NeonEventFilter(self)
            self.installEventFilter(self._neon_filter)
            self.viewport().installEventFilter(self._neon_filter)
        self.setStyleSheet(
            "QTableWidget, QTableWidget::viewport{border:1px solid transparent;}\n"
            "QTableWidget::item:selected{background:transparent;}"
        )

    def focusOutEvent(self, e):
        super().focusOutEvent(e)
        self.clearSelection()

    def itemSelectionChanged(self):
        super().itemSelectionChanged()
        self.clearSelection()

    def edit(self, index, trigger, event):
        res = super().edit(index, trigger, event)
        if res and self._neon_filter is not None:
            editor = self.findChild(QtWidgets.QLineEdit)
            if editor and getattr(editor, "_neon_filter", None) is None:
                editor.setAttribute(QtCore.Qt.WA_Hover, True)
                editor.setStyleSheet(editor.styleSheet() + "border:1px solid transparent;")
                filt = NeonEventFilter(editor)
                editor.installEventFilter(filt)
                editor._neon_filter = filt
                apply_neon_effect(editor, True, shadow=False)
                if self._active_editor is not None and self._active_editor is not editor:
                    self._active_editor.removeEventFilter(self)
                    apply_neon_effect(self._active_editor, False)
                self._active_editor = editor
                editor.installEventFilter(self)
        return res

    def eventFilter(self, obj, event):  # noqa: D401 - Qt event filter signature
        if obj is self._active_editor and event.type() == QtCore.QEvent.FocusOut:
            apply_neon_effect(obj, False)
            obj.removeEventFilter(self)
            self._active_editor = None
        return super().eventFilter(obj, event)

class ExcelCalendarTable(QtWidgets.QTableWidget):
    """Таблица календаря месяца с вложенными таблицами по дням."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.verticalHeader().setVisible(False)
        self.setShowGrid(True)
        self.setWordWrap(True)
        day_names = ["ПН", "ВТ", "СР", "ЧТ", "ПТ", "СБ", "ВС"]
        self.setColumnCount(len(day_names))
        self.setHorizontalHeaderLabels(day_names)
        self.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.setFocusPolicy(QtCore.Qt.NoFocus)

        now = datetime.now()
        self.year = now.year
        self.month = now.month
        self.date_map: Dict[tuple[int, int], date] = {}
        self.cell_tables: Dict[tuple[int, int], QtWidgets.QTableWidget] = {}
        self.day_labels: Dict[tuple[int, int], QtWidgets.QLabel] = {}
        self.cell_containers: Dict[tuple[int, int], QtWidgets.QWidget] = {}
        self.cell_filters: Dict[tuple[int, int], NeonEventFilter] = {}

        self._col_widths: List[int] | None = None

        self._updating_rows = False

        # Timer for deferred row height updates
        self._row_timer = QtCore.QTimer(self)
        self._row_timer.setSingleShot(True)
        self._row_timer.timeout.connect(self._update_row_heights)
        self.destroyed.connect(lambda: self._row_timer.stop())

        self.load_month_data(self.year, self.month)

    def mousePressEvent(self, event):
        self.clearSelection()
        event.ignore()

    def _create_inner_table(self) -> QtWidgets.QTableWidget:
        tbl = NeonTableWidget(CONFIG.get("day_rows", 6), 3, self, use_neon=True)
        tbl.setHorizontalHeaderLabels(["Работа", "План", "Готово"])
        tbl.verticalHeader().setVisible(False)
        tbl.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        tbl.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked)
        tbl.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        tbl.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        if self._col_widths:
            for i, w in enumerate(self._col_widths):
                tbl.setColumnWidth(i, w)
        tbl.horizontalHeader().sectionResized.connect(self._sync_day_columns)
        return tbl

    def set_day_column_widths(self, widths: Iterable[int]):
        self._col_widths = [int(w) for w in widths]
        for tbl in self.cell_tables.values():
            header = tbl.horizontalHeader()
            blocker = QtCore.QSignalBlocker(header)
            for i, w in enumerate(self._col_widths):
                tbl.setColumnWidth(i, w)

    def _sync_day_columns(self, *_):
        header = self.sender()
        if not isinstance(header, QtWidgets.QHeaderView):
            return
        widths = [header.sectionSize(i) for i in range(header.count())]
        self.set_day_column_widths(widths)
        settings = QtCore.QSettings("rabota2", "rabota2")
        settings.setValue("MainWindow/columns", self._col_widths)
        settings.sync()

    def get_day_column_widths(self) -> List[int]:
        tbl = next(iter(self.cell_tables.values()), None)
        if tbl:
            return [tbl.columnWidth(i) for i in range(tbl.columnCount())]
        return []

    def update_day_rows(self):
        rows = CONFIG.get("day_rows", 6)
        for tbl in self.cell_tables.values():
            tbl.setRowCount(rows)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Schedule a deferred update of row heights
        self._row_timer.stop()
        self._row_timer.start(0)

    def _update_row_heights(self):
        if self._updating_rows:
            return
        if not shiboken6.isValid(self) or self.model() is None:
            return
        self._updating_rows = True
        try:
            rows = self.rowCount()
            if rows:
                height = self.viewport().height() // rows
                for r in range(rows):
                    self.setRowHeight(r, height)
        finally:
            self._updating_rows = False

    # ---------- Persistence ----------
    def save_current_month(self):
        md = MonthData(year=self.year, month=self.month)
        for (r, c), day in self.date_map.items():
            if day.month != self.month:
                continue
            inner = self.cell_tables.get((r, c))
            if not inner:
                continue
            rows = []
            for rr in range(inner.rowCount()):
                vals = []
                for cc in range(inner.columnCount()):
                    it = inner.item(rr, cc)
                    vals.append(it.text().strip() if it else "")
                if any(vals):
                    rows.append({"work": vals[0], "plan": vals[1], "done": vals[2]})
            if rows:
                md.days[day.day] = rows
        md.save()

    def load_month_data(self, year: int, month: int):
        self.year = year
        self.month = month
        md = MonthData.load(year, month)
        cal = calendar.Calendar()
        weeks = cal.monthdatescalendar(year, month)
        self.date_map.clear()
        self.cell_tables.clear()
        self.day_labels.clear()
        self.cell_containers.clear()
        self.cell_filters.clear()
        self.setRowCount(len(weeks))
        base_style = "border:1px solid transparent;"
        for r, week in enumerate(weeks):
            for c, day in enumerate(week):
                container = QtWidgets.QWidget()
                lay = QtWidgets.QVBoxLayout(container)
                lay.setContentsMargins(0, 0, 0, 0)
                lay.setSpacing(2)
                container.setStyleSheet(base_style)
                lbl = QtWidgets.QLabel(str(day.day), container)
                lbl.setFont(
                    QtGui.QFont(
                        CONFIG.get(
                            "header_font", CONFIG.get("font_family", "Exo 2")
                        )
                    )
                )
                lbl.setAlignment(QtCore.Qt.AlignCenter)
                lbl.setAttribute(QtCore.Qt.WA_Hover, True)
                # keep reference for later font updates
                self.day_labels[(r, c)] = lbl
                lay.addWidget(lbl, alignment=QtCore.Qt.AlignHCenter)
                inner = self._create_inner_table()
                lay.addWidget(inner)
                filt = None
                if neon_enabled():
                    container.setAttribute(QtCore.Qt.WA_Hover, True)
                    filt = NeonEventFilter(container)
                    container.installEventFilter(filt)
                    inner.installEventFilter(filt)
                    inner.viewport().installEventFilter(filt)
                    container._neon_filter = filt

                    lbl.setAttribute(QtCore.Qt.WA_Hover, True)
                    lbl_filt = NeonEventFilter(lbl)
                    lbl.installEventFilter(lbl_filt)
                    lbl._neon_filter = lbl_filt
                self.setCellWidget(r, c, container)
                self.date_map[(r, c)] = day
                self.cell_tables[(r, c)] = inner
                self.cell_containers[(r, c)] = container
                self.cell_filters[(r, c)] = filt
                if day.month != month:
                    container.setEnabled(False)
                    container.setStyleSheet(
                        base_style + "background-color:#2a2a2a; color:#777;"
                    )
                    lbl.setStyleSheet("color:#777;")
                else:
                    container.setEnabled(True)
                    container.setStyleSheet(base_style)
                    lbl.setStyleSheet("")
                rows = md.days.get(day.day, [])
                for rr, row in enumerate(rows):
                    if rr >= inner.rowCount():
                        break
                    for cc, key in enumerate(["work", "plan", "done"]):
                            item = QtWidgets.QTableWidgetItem(str(row.get(key, "")))
                            inner.setItem(rr, cc, item)
        self._update_row_heights()
        self.apply_fonts()
        return True

    def apply_fonts(self):
        """Apply current header font to calendar elements."""
        header_font = QtGui.QFont(
            CONFIG.get("header_font", CONFIG.get("font_family", "Exo 2"))
        )
        self.horizontalHeader().setFont(header_font)
        # Ensure weekday header items on the main calendar adopt the font
        for i in range(self.columnCount()):
            item = self.horizontalHeaderItem(i)
            if item:
                item.setFont(header_font)
        for tbl in self.cell_tables.values():
            tbl.horizontalHeader().setFont(header_font)
            # Ensure individual header items also adopt the selected font
            for i in range(tbl.columnCount()):
                item = tbl.horizontalHeaderItem(i)
                if item:
                    item.setFont(header_font)
        for lbl in self.day_labels.values():
            lbl.setFont(header_font)

    def apply_theme(self) -> None:
        """Apply palette-based colors to calendar widgets."""
        workspace = QtGui.QColor(CONFIG.get("workspace_color", "#1e1e21"))
        if CONFIG.get("monochrome", False):
            workspace = theme_manager.apply_monochrome(workspace)
        text_color = QtWidgets.QApplication.palette().color(QtGui.QPalette.WindowText).name()
        ws = workspace.name()
        style = (
            f"QTableWidget{{background-color:{ws};color:{text_color};}}"
            f"QTableWidget::item{{background-color:{ws};color:{text_color};}}"
            f"QTableWidget::item:selected{{background-color:{ws};color:{text_color};}}"
        )
        self.setStyleSheet(style)
        self.horizontalHeader().setStyleSheet(f"background-color:{ws};")
        for tbl in self.cell_tables.values():
            tbl.setStyleSheet(style)
            tbl.horizontalHeader().setStyleSheet(f"background-color:{ws};")
        for container in self.cell_containers.values():
            container.setStyleSheet(f"background-color:{ws};color:{text_color};")

    # ---------- Navigation ----------
    def go_prev_month(self):
        self.save_current_month()
        if self.month == 1:
            self.month = 12
            self.year -= 1
        else:
            self.month -= 1
        self.load_month_data(self.year, self.month)

    def go_next_month(self):
        self.save_current_month()
        if self.month == 12:
            self.month = 1
            self.year += 1
        else:
            self.month += 1
        self.load_month_data(self.year, self.month)

class CollapsibleSidebar(QtWidgets.QFrame):
    toggled = QtCore.Signal(bool)
    settings_clicked = QtCore.Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Sidebar")
        self.setStyleSheet(
            """
            #Sidebar { background-color: #1f1f23; }
            QLabel { color: #c7c7c7; }
            """
        )
        self.expanded_width=260; self.collapsed_width=64
        lay=QtWidgets.QVBoxLayout(self); lay.setContentsMargins(8,8,8,8); lay.setSpacing(6)

        # Toggle button — только иконка
        self.btn_toggle = StyledToolButton(self)
        self.btn_toggle.setIcon(QtGui.QIcon(CONFIG.get("sidebar_icon", ICON_TOGGLE)))
        self.btn_toggle.setIconSize(QtCore.QSize(28,28))
        self.btn_toggle.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly)
        self.btn_toggle.setCursor(QtCore.Qt.PointingHandCursor)
        self.btn_toggle.clicked.connect(self.toggle)
        lay.addWidget(self.btn_toggle)

        line = QtWidgets.QFrame(self); line.setFrameShape(QtWidgets.QFrame.HLine); line.setStyleSheet("color:#333;")
        lay.addWidget(line)

        items = [
            ("Вводные", ICON_TM),
            ("Выкладка", ICON_TQ),
            ("Аналитика", ICON_TG),
            ("Топы", ICON_TP),
        ]
        self.buttons = []
        self.btn_inputs = None
        self.btn_release = None
        self.btn_analytics = None
        self.btn_tops = None
        for label, icon_path in items:
            b = StyledToolButton(self)
            b.setIcon(QtGui.QIcon(icon_path))
            b.setIconSize(QtCore.QSize(22, 22))
            b.setText(" " + label)
            b.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
            b.setCursor(QtCore.Qt.PointingHandCursor)
            b.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
            b.setProperty("neon_selected", False)
            lay.addWidget(b)
            self.buttons.append(b)
            b.clicked.connect(lambda _, btn=b: self.activate_button(btn))
            if label == "Вводные":
                self.btn_inputs = b
            elif label == "Выкладка":
                self.btn_release = b
            elif label == "Аналитика":
                self.btn_analytics = b
            elif label == "Топы":
                self.btn_tops = b
        # Settings button at the bottom
        self.btn_settings = StyledToolButton(self)
        self.btn_settings.setIcon(icon("settings"))
        self.btn_settings.setIconSize(QtCore.QSize(22, 22))
        self.btn_settings.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly)
        self.btn_settings.clicked.connect(self.settings_clicked)
        # Push other buttons to the top when sidebar is taller than content
        lay.addStretch(1)
        lay.addWidget(self.btn_settings)

        self._collapsed = False
        self.anim = QtCore.QPropertyAnimation(self, b"maximumWidth", self)
        self.anim.setDuration(160)
        self.setMinimumWidth(self.collapsed_width)
        self.setMaximumWidth(self.expanded_width)
        self.update_icons()
        self.apply_fonts()

    def activate_button(self, btn: QtWidgets.QToolButton) -> None:
        for b in self.buttons:
            selected = b is btn
            b.setProperty("neon_selected", selected)
            b.apply_base_style()
            b._apply_hover(selected)
            apply_neon_effect(b, selected and neon_enabled())

        for b in self.buttons:
            b._neon_prev_style = b.styleSheet()

    def set_collapsed(self, collapsed: bool):
        self._collapsed = collapsed
        start = self.width()
        end = self.collapsed_width if collapsed else self.expanded_width
        self.anim.stop()
        self.anim.setStartValue(start)
        self.anim.setEndValue(end)
        self.anim.start()
        for b in self.buttons:
            b.setToolButtonStyle(
                QtCore.Qt.ToolButtonIconOnly if collapsed else QtCore.Qt.ToolButtonTextBesideIcon
            )
        self.toggled.emit(not collapsed)

    def toggle(self): self.set_collapsed(not self._collapsed)

    def apply_style(
        self,
        neon: bool,
        accent: QtGui.QColor,
        sidebar_color: Union[str, QtGui.QColor, None] = None,
    ):
        size = CONFIG.get("neon_size", 10)
        intensity = CONFIG.get("neon_intensity", 255)
        if sidebar_color is None:
            sidebar_color = CONFIG.get("sidebar_color", "#1f1f23")
        if isinstance(sidebar_color, QtGui.QColor):
            sidebar_color = sidebar_color.name()

        label_color = accent.name() if neon else "#c7c7c7"
        style = (
            f"#Sidebar {{ background-color: {sidebar_color}; }}\n"
            f"QLabel {{ color: {label_color}; }}\n"
        )
        self.setStyleSheet(style)

        widgets = [self.btn_toggle] + self.buttons + [self.btn_settings]
        for w in widgets:
            w.apply_base_style()
            w._apply_hover(w.property("neon_selected"))
            apply_neon_effect(
                w, w.property("neon_selected") and neon_enabled()
            )

        if neon:
            for w in widgets:
                eff = QtWidgets.QGraphicsDropShadowEffect(self)
                eff.setOffset(0, 0)
                eff.setBlurRadius(size)
                c = QtGui.QColor(accent)
                c.setAlpha(intensity)
                eff.setColor(c)
                w.setGraphicsEffect(eff)
        else:
            for w in widgets:
                if hasattr(w, "_neon_anim") and w._neon_anim:
                    w._neon_anim.stop()
                    w._neon_anim = None
                w.setGraphicsEffect(None)

    def update_icons(self) -> None:
        """Update sidebar icon from configuration."""
        self.btn_toggle.setIcon(QtGui.QIcon(CONFIG.get("sidebar_icon", ICON_TOGGLE)))
        self.btn_settings.setIcon(icon("settings"))

    def apply_fonts(self):
        """Apply configured sidebar font to heading widgets."""
        family = CONFIG.get("sidebar_font", CONFIG.get("header_font", "Exo 2"))
        font = QtGui.QFont(family)
        widgets = [self.btn_toggle] + self.buttons + [self.btn_settings]
        for w in widgets:
            w.setFont(font)


class SettingsDialog(QtWidgets.QDialog):
    """Окно настроек приложения."""

    settings_changed = QtCore.Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        # keep a reference to the main window
        self.main_window = parent if isinstance(parent, QtWidgets.QWidget) else None
        self.setWindowTitle("Настройки")
        self.resize(500, 400)
        main_lay = QtWidgets.QVBoxLayout(self)
        tabs = QtWidgets.QTabWidget(self)
        main_lay.addWidget(tabs)

        # --- Интерфейс ---
        tab_interface = QtWidgets.QWidget()
        form_interface = QtWidgets.QFormLayout(tab_interface)
        tabs.addTab(tab_interface, "Интерфейс")

        self._accent_color = QtGui.QColor(CONFIG.get("accent_color", "#39ff14"))
        self._preset_colors = [
            ("Зелёный", QtGui.QColor("#39ff14")),
            ("Красный", QtGui.QColor("#ff5555")),
            ("Синий", QtGui.QColor("#2d7cdb")),
            ("Жёлтый", QtGui.QColor("#ffd700")),
            ("Фиолетовый", QtGui.QColor("#8a2be2")),
        ]
        self.combo_accent = QtWidgets.QComboBox(self)
        for name, color in self._preset_colors:
            pix = QtGui.QPixmap(16, 16)
            pix.fill(color)
            self.combo_accent.addItem(QtGui.QIcon(pix), name)
        self.combo_accent.addItem("Другой…")
        other_index = self.combo_accent.count() - 1
        current = self._accent_color.name().lower()
        idx = next(
            (i for i, (_, c) in enumerate(self._preset_colors) if c.name().lower() == current),
            other_index,
        )
        self.combo_accent.blockSignals(True)
        self.combo_accent.setCurrentIndex(idx)
        self.combo_accent.blockSignals(False)
        if idx == other_index:
            pix = QtGui.QPixmap(16, 16)
            pix.fill(self._accent_color)
            self.combo_accent.setItemIcon(other_index, QtGui.QIcon(pix))
        self._accent_index = idx
        # connect after setting initial index to avoid unwanted color dialog
        # use activated so selecting "Другой" again reopens the color picker
        self.combo_accent.activated.connect(self._on_accent_changed)
        form_interface.addRow("Цвет подсветки", self.combo_accent)

        self._workspace_color = QtGui.QColor(CONFIG.get("workspace_color", "#1e1e21"))
        self.btn_workspace = QtWidgets.QPushButton(self)
        self.btn_workspace.setCursor(QtCore.Qt.PointingHandCursor)
        self._update_workspace_button()
        self.btn_workspace.clicked.connect(self.choose_workspace_color)
        form_interface.addRow("Цвет рабочей области", self.btn_workspace)

        self._sidebar_color = QtGui.QColor(CONFIG.get("sidebar_color", "#1f1f23"))
        self.btn_sidebar = QtWidgets.QPushButton(self)
        self.btn_sidebar.setCursor(QtCore.Qt.PointingHandCursor)
        self._update_sidebar_button()
        self.btn_sidebar.clicked.connect(self.choose_sidebar_color)
        form_interface.addRow("Цвет боковой панели", self.btn_sidebar)

        # gradient controls
        grad = CONFIG.get("gradient_colors", ["#39ff14", "#2d7cdb"])
        self._grad_color1 = QtGui.QColor(grad[0])
        self._grad_color2 = QtGui.QColor(grad[1])
        self.btn_grad1 = QtWidgets.QPushButton(self)
        self.btn_grad1.setCursor(QtCore.Qt.PointingHandCursor)
        self.btn_grad2 = QtWidgets.QPushButton(self)
        self.btn_grad2.setCursor(QtCore.Qt.PointingHandCursor)
        self._update_grad_buttons()
        self.btn_grad1.clicked.connect(lambda: self.choose_grad_color(1))
        self.btn_grad2.clicked.connect(lambda: self.choose_grad_color(2))
        lay_grad = QtWidgets.QHBoxLayout(); lay_grad.addWidget(self.btn_grad1); lay_grad.addWidget(self.btn_grad2)
        form_interface.addRow("Градиент", lay_grad)
        self.sld_grad_angle = QtWidgets.QSlider(QtCore.Qt.Horizontal, self)
        self.sld_grad_angle.setRange(0, 360)
        self.sld_grad_angle.setValue(int(CONFIG.get("gradient_angle", 0)))
        self.lbl_grad_angle = QtWidgets.QLabel(str(self.sld_grad_angle.value()))
        self.sld_grad_angle.valueChanged.connect(lambda v: (self.lbl_grad_angle.setText(str(v)), self._save_config()))
        lay_angle = QtWidgets.QHBoxLayout(); lay_angle.addWidget(self.sld_grad_angle,1); lay_angle.addWidget(self.lbl_grad_angle)
        form_interface.addRow("Угол градиента", lay_angle)
        # neon controls
        grp_neon = QtWidgets.QGroupBox("Неон", self)
        lay_neon = QtWidgets.QFormLayout(grp_neon)
        self.chk_neon = QtWidgets.QCheckBox(self)
        self.chk_neon.setChecked(CONFIG.get("neon", False))
        self.chk_neon.toggled.connect(self._on_neon_changed)
        lay_neon.addRow("Включен", self.chk_neon)
        self.spin_neon_size = QtWidgets.QSpinBox(self)
        self.spin_neon_size.setRange(0, 200)
        self.spin_neon_size.setValue(CONFIG.get("neon_size", 10))
        self.spin_neon_size.valueChanged.connect(self._on_neon_changed)
        lay_neon.addRow("Размер", self.spin_neon_size)
        self.spin_neon_thickness = QtWidgets.QSpinBox(self)
        self.spin_neon_thickness.setRange(0, 10)
        self.spin_neon_thickness.setValue(CONFIG.get("neon_thickness", 1))
        self.spin_neon_thickness.valueChanged.connect(self._on_neon_changed)
        lay_neon.addRow("Толщина", self.spin_neon_thickness)
        self.spin_neon_intensity = QtWidgets.QSpinBox(self)
        self.spin_neon_intensity.setRange(0, 255)
        self.spin_neon_intensity.setValue(CONFIG.get("neon_intensity", 255))
        self.spin_neon_intensity.valueChanged.connect(self._on_neon_changed)
        lay_neon.addRow("Интенсивность", self.spin_neon_intensity)
        form_interface.addRow(grp_neon)


        self.font_header = QtWidgets.QFontComboBox(self)
        self.font_header.setCurrentFont(
            QtGui.QFont(CONFIG.get("header_font", CONFIG.get("font_family", "Exo 2")))
        )
        self.font_header.currentFontChanged.connect(lambda _: self.apply_fonts())
        form_interface.addRow("Шрифт заголовков", self.font_header)

        self.font_text = QtWidgets.QFontComboBox(self)
        self.font_text.setCurrentFont(
            QtGui.QFont(CONFIG.get("text_font", CONFIG.get("font_family", "Exo 2")))
        )
        self.font_text.currentFontChanged.connect(lambda _: self._save_config())
        form_interface.addRow("Шрифт текста", self.font_text)

        self.font_sidebar = QtWidgets.QFontComboBox(self)
        self.font_sidebar.setCurrentFont(QtGui.QFont(CONFIG.get("sidebar_font", "Exo 2")))
        self.font_sidebar.currentFontChanged.connect(lambda _: self._on_sidebar_font_changed())
        form_interface.addRow("Шрифт боковой панели", self.font_sidebar)

        path_lay = QtWidgets.QHBoxLayout()
        self.edit_path = QtWidgets.QLineEdit(CONFIG.get("save_path", DATA_DIR), self)
        self.edit_path.editingFinished.connect(self._save_config)
        btn_browse = StyledPushButton("...", self)
        btn_browse.clicked.connect(self.browse_path)
        path_lay.addWidget(self.edit_path, 1)
        path_lay.addWidget(btn_browse)
        form_interface.addRow("Путь сохранения", path_lay)

        # --- Иконки ---
        tab_icons = QtWidgets.QWidget()
        form_icons = QtWidgets.QFormLayout(tab_icons)
        icon_files = [
            f for f in os.listdir(ASSETS) if f.lower().endswith((".png", ".ico"))
        ]

        # sidebar icon
        self.combo_sidebar_icon = QtWidgets.QComboBox(self)
        for f in icon_files:
            path = os.path.join(ASSETS, f)
            self.combo_sidebar_icon.addItem(QtGui.QIcon(path), f, path)
        current_sidebar = CONFIG.get("sidebar_icon", ICON_TOGGLE)
        idx = self.combo_sidebar_icon.findData(current_sidebar)
        if idx < 0 and os.path.isfile(current_sidebar):
            self.combo_sidebar_icon.addItem(
                QtGui.QIcon(current_sidebar), os.path.basename(current_sidebar), current_sidebar
            )
            idx = self.combo_sidebar_icon.count() - 1
        if idx >= 0:
            self.combo_sidebar_icon.setCurrentIndex(idx)
        self.combo_sidebar_icon.currentIndexChanged.connect(lambda _: self._save_config())
        btn_sidebar_browse = StyledPushButton("Обзор…", self)
        btn_sidebar_browse.clicked.connect(
            lambda: self.browse_icon(self.combo_sidebar_icon)
        )
        lay_sidebar = QtWidgets.QHBoxLayout()
        lay_sidebar.addWidget(self.combo_sidebar_icon, 1)
        lay_sidebar.addWidget(btn_sidebar_browse)
        form_icons.addRow("Иконка боковой панели", lay_sidebar)

        # application icon
        self.combo_app_icon = QtWidgets.QComboBox(self)
        for f in icon_files:
            path = os.path.join(ASSETS, f)
            self.combo_app_icon.addItem(QtGui.QIcon(path), f, path)
        current_app = CONFIG.get("app_icon", ICON_TOGGLE)
        idx = self.combo_app_icon.findData(current_app)
        if idx < 0 and os.path.isfile(current_app):
            self.combo_app_icon.addItem(
                QtGui.QIcon(current_app), os.path.basename(current_app), current_app
            )
            idx = self.combo_app_icon.count() - 1
        if idx >= 0:
            self.combo_app_icon.setCurrentIndex(idx)
        self.combo_app_icon.currentIndexChanged.connect(lambda _: self._save_config())
        btn_app_browse = StyledPushButton("Обзор…", self)
        btn_app_browse.clicked.connect(
            lambda: self.browse_icon(self.combo_app_icon)
        )
        lay_app = QtWidgets.QHBoxLayout()
        lay_app.addWidget(self.combo_app_icon, 1)
        lay_app.addWidget(btn_app_browse)
        form_icons.addRow("Иконка приложения", lay_app)

        hint = QtWidgets.QLabel("Размер иконок: 20×20 px", self)
        hint.setAlignment(QtCore.Qt.AlignCenter)
        form_icons.addRow(hint)
        tabs.addTab(tab_icons, "Иконки")

        # General options below tabs
        form_gen = QtWidgets.QFormLayout()
        main_lay.addLayout(form_gen)
        self.spin_day_rows = QtWidgets.QSpinBox(self)
        self.spin_day_rows.setRange(1, 20)
        self.spin_day_rows.setValue(CONFIG.get("day_rows", 6))
        self.spin_day_rows.setFixedWidth(self.spin_day_rows.sizeHint().width() + 20)
        self.spin_day_rows.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
        self.spin_day_rows.valueChanged.connect(lambda _: self._save_config())
        form_gen.addRow("Строк на день", self.spin_day_rows)

        box = QtWidgets.QDialogButtonBox(self)
        btn_save = StyledPushButton("Сохранить", self)
        btn_save.setIcon(icon("save"))
        btn_save.setIconSize(QtCore.QSize(20, 20))
        btn_save.setFixedSize(btn_save.sizeHint())
        btn_cancel = StyledPushButton("Отмена", self)
        btn_cancel.setIcon(icon("x"))
        btn_cancel.setIconSize(QtCore.QSize(20, 20))
        btn_cancel.setFixedSize(btn_cancel.sizeHint())
        box.addButton(btn_save, QtWidgets.QDialogButtonBox.AcceptRole)
        box.addButton(btn_cancel, QtWidgets.QDialogButtonBox.RejectRole)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        main_lay.addWidget(box)

        self._settings = QtCore.QSettings("rabota2", "rabota2")
        geom = self._settings.value("SettingsDialog/geometry")
        if geom is not None:
            self.restoreGeometry(geom)

        apply_neon_to_inputs(self)

    def closeEvent(self, event):
        self._settings.setValue("SettingsDialog/geometry", self.saveGeometry())
        super().closeEvent(event)

    def browse_path(self):
        path = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Выбрать папку", self.edit_path.text()
        )
        if path:
            self.edit_path.setText(path)
            self._save_config()

    def browse_icon(self, combo: QtWidgets.QComboBox):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Выберите иконку",
            "",
            "Image Files (*.png *.ico)",
        )
        if path:
            idx = combo.findData(path)
            if idx < 0:
                combo.addItem(QtGui.QIcon(path), os.path.basename(path), path)
                idx = combo.count() - 1
            combo.setCurrentIndex(idx)
            self._save_config()

    def _collect_config(self):
        return {
            "neon": self.chk_neon.isChecked(),
            "neon_size": self.spin_neon_size.value(),
            "neon_thickness": self.spin_neon_thickness.value(),
            "neon_intensity": self.spin_neon_intensity.value(),
            "accent_color": self._accent_color.name(),
            "gradient_colors": [self._grad_color1.name(), self._grad_color2.name()],
            "gradient_angle": self.sld_grad_angle.value(),
            "workspace_color": self._workspace_color.name(),
            "sidebar_color": self._sidebar_color.name(),
            "font_family": self.font_text.currentFont().family(),
            "header_font": self.font_header.currentFont().family(),
            "text_font": self.font_text.currentFont().family(),
            "sidebar_font": self.font_sidebar.currentFont().family(),
            "sidebar_icon": self.combo_sidebar_icon.currentData(),
            "app_icon": self.combo_app_icon.currentData(),
            "day_rows": self.spin_day_rows.value(),
            "save_path": self.edit_path.text().strip() or DATA_DIR,
        }

    def _on_sidebar_font_changed(self):
        self._save_config()
        if self.main_window and hasattr(self.main_window, "sidebar"):
            self.main_window.sidebar.apply_fonts()

    def _on_neon_changed(self):
        self._save_config()
        parent = self.parent()
        if parent is not None and hasattr(parent, "apply_style"):
            parent.apply_style(self.chk_neon.isChecked())
            update_neon_filters(parent)
        update_neon_filters(self)

    def _save_config(self):
        config = self._collect_config()
        CONFIG.update(config)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(CONFIG, f, ensure_ascii=False, indent=2)
        self.settings_changed.emit()

    def save(self) -> None:
        """Public method to persist configuration changes."""
        self._save_config()

    def apply_fonts(self):
        """Save and apply new header font immediately."""
        self._save_config()
        theme_manager.set_header_font(self.font_header.currentFont().family())
        parent = self.parent()
        if parent and hasattr(parent, "table"):
            parent.table.apply_fonts()
            if hasattr(parent, "topbar"):
                parent.topbar.update_labels()

    def accept(self):
        self._save_config()
        theme_manager.set_text_font(self.font_text.currentFont().family())
        theme_manager.set_header_font(self.font_header.currentFont().family())
        parent = self.parent()
        if parent and hasattr(parent, "table"):
            parent.table.apply_fonts()
            if hasattr(parent, "topbar"):
                parent.topbar.update_labels()
        super().accept()

    def _on_accent_changed(self, idx):
        other_index = self.combo_accent.count() - 1
        if idx == other_index:
            color = QtWidgets.QColorDialog.getColor(
                self._accent_color, self, "Цвет"
            )
            if color.isValid():
                self._accent_color = color
                pix = QtGui.QPixmap(16, 16)
                pix.fill(color)
                self.combo_accent.setItemIcon(idx, QtGui.QIcon(pix))
                self._accent_index = idx
            else:
                self.combo_accent.blockSignals(True)
                self.combo_accent.setCurrentIndex(self._accent_index)
                self.combo_accent.blockSignals(False)
        else:
            self._accent_color = self._preset_colors[idx][1]
            self._accent_index = idx
        self._save_config()

    def _update_workspace_button(self):
        color = self._workspace_color.name()
        self.btn_workspace.setStyleSheet(
            f"QPushButton{{background:{color}; border:1px solid #555;}} QPushButton:hover{{background:{color};}}"
        )

    def choose_workspace_color(self):
        color = QtWidgets.QColorDialog.getColor(
            self._workspace_color, self, "Цвет"
        )
        if color.isValid():
            self._workspace_color = color
            self._update_workspace_button()
            self._save_config()
            parent = self.parent()
            if parent is not None:
                if hasattr(parent, "table"):
                    parent.table.apply_theme()
                if hasattr(parent, "topbar"):
                    parent.topbar.apply_background(self._workspace_color)

    def _update_sidebar_button(self):
        color = self._sidebar_color.name()
        self.btn_sidebar.setStyleSheet(
            f"QPushButton{{background:{color}; border:1px solid #555;}} QPushButton:hover{{background:{color};}}"
        )

    def choose_sidebar_color(self):
        color = QtWidgets.QColorDialog.getColor(
            self._sidebar_color, self, "Цвет"
        )
        if color.isValid():
            self._sidebar_color = color
            self._update_sidebar_button()
            self._save_config()

    def _update_grad_buttons(self):
        c1 = self._grad_color1.name()
        c2 = self._grad_color2.name()
        self.btn_grad1.setStyleSheet(
            f"QPushButton{{background:{c1}; border:1px solid #555;}} QPushButton:hover{{background:{c1};}}"
        )
        self.btn_grad2.setStyleSheet(
            f"QPushButton{{background:{c2}; border:1px solid #555;}} QPushButton:hover{{background:{c2};}}"
        )

    def choose_grad_color(self, idx):
        color = QtWidgets.QColorDialog.getColor(
            self._grad_color1 if idx == 1 else self._grad_color2,
            self,
            "Цвет",
        )
        if color.isValid():
            if idx == 1:
                self._grad_color1 = color
            else:
                self._grad_color2 = color
            self._update_grad_buttons()
            self._save_config()
            parent = self.parent()
            if parent is not None and hasattr(parent, "apply_settings"):
                parent.apply_settings()


class TopBar(QtWidgets.QWidget):
    prev_clicked = QtCore.Signal()
    next_clicked = QtCore.Signal()
    year_changed = QtCore.Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QtWidgets.QHBoxLayout(self)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(8)
        self.btn_prev = StyledToolButton(self)
        self.btn_prev.setCursor(QtCore.Qt.PointingHandCursor)
        self.btn_prev.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly)
        self.btn_prev.clicked.connect(self.prev_clicked)
        lay.addWidget(self.btn_prev)
        lay.addStretch(1)
        self.lbl_month = QtWidgets.QLabel("Месяц")
        self.lbl_month.setAlignment(QtCore.Qt.AlignCenter)
        self.lbl_month.setContentsMargins(8, 0, 8, 0)
        lay.addWidget(self.lbl_month)
        self.spin_year = QtWidgets.QSpinBox(self)
        self.spin_year.setRange(2000, 2100)
        self.spin_year.setValue(datetime.now().year)
        self.spin_year.setStyleSheet("padding:0 4px;")
        self.spin_year.setFixedWidth(self.spin_year.sizeHint().width())
        self.spin_year.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
        self.spin_year.valueChanged.connect(self.year_changed.emit)
        lay.addWidget(self.spin_year)
        self.spin_year.setAttribute(QtCore.Qt.WA_Hover, True)
        if neon_enabled():
            filt = NeonEventFilter(self.spin_year)
            self.spin_year.installEventFilter(filt)
            self.spin_year._neon_filter = filt
        lay.addStretch(1)
        self.btn_next = StyledToolButton(self)
        self.btn_next.setCursor(QtCore.Qt.PointingHandCursor)
        self.btn_next.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly)
        self.btn_next.clicked.connect(self.next_clicked)
        lay.addWidget(self.btn_next)
        self.update_icons()
        # Base stylesheet used when neon is disabled in dark theme
        # Explicitly remove borders for QLabel to avoid unwanted frames
        self.default_style = (
            "QLabel{color:#e5e5e5; border:none;} "
            "QToolButton{color:#e5e5e5; border:1px solid #555; border-radius:6px; padding:4px 10px;}"
        )
        self.setAutoFillBackground(True)
        color = QtGui.QColor(CONFIG.get("workspace_color", "#1e1e21"))
        if CONFIG.get("monochrome", False):
            color = theme_manager.apply_monochrome(color)
        self.apply_background(color)
        self.apply_style(CONFIG.get("neon", False))
        self.apply_fonts()

    def apply_fonts(self):
        font = QtGui.QFont(
            CONFIG.get("header_font", CONFIG.get("font_family", "Exo 2"))
        )
        font.setBold(True)
        self.lbl_month.setFont(font)

    def update_labels(self):
        """Refresh label fonts based on current configuration."""
        self.apply_fonts()

    def update_icons(self) -> None:
        self.btn_prev.setIcon(icon("chevron-left"))
        self.btn_prev.setIconSize(QtCore.QSize(22, 22))
        self.btn_next.setIcon(icon("chevron-right"))
        self.btn_next.setIconSize(QtCore.QSize(22, 22))

    def apply_background(self, color: Union[str, QtGui.QColor]) -> None:
        qcolor = QtGui.QColor(color)
        for w in (self, self.spin_year):
            pal = w.palette()
            pal.setColor(QtGui.QPalette.Window, qcolor)
            w.setAutoFillBackground(True)
            w.setPalette(pal)
        pal = self.lbl_month.palette()
        pal.setColor(QtGui.QPalette.Window, qcolor)
        self.lbl_month.setAutoFillBackground(False)
        self.lbl_month.setPalette(pal)
        self.lbl_month.setStyleSheet("background:transparent; border:none;")
        # Ensure rounded spin box with a semi-transparent border for contrast
        self.spin_year.setStyleSheet(
            f"background-color:{qcolor.name()}; padding:0 4px;"
            " border-radius:6px; border:1px solid rgba(255,255,255,0.2);"
        )
        self.apply_fonts()

    def update_background(self, color: Union[str, QtGui.QColor]) -> None:
        qcolor = QtGui.QColor(color)
        self.setStyleSheet(
            f"background-color:{qcolor.name()};" + self.styleSheet()
        )
        self.apply_fonts()

    def apply_style(self, neon):
        accent = QtGui.QColor(CONFIG.get("accent_color", "#39ff14"))
        if CONFIG.get("monochrome", False):
            h, s, v, _ = accent.getHsv()
            s = int(CONFIG.get("mono_saturation", 100))
            accent.setHsv(h, s, v)
        thickness = CONFIG.get("neon_thickness", 1)
        # ``neon_size`` and ``neon_intensity`` are no longer used directly but kept
        # for compatibility with existing configuration files.
        CONFIG.get("neon_size", 10)
        CONFIG.get("neon_intensity", 255)
        if neon:
            style = (
                "QLabel{color:#e5e5e5; border:none;} "
                f"QToolButton{{color:{accent.name()}; border:{thickness}px solid {accent.name()}; border-radius:6px; padding:4px 10px;}}"
            )
            self.setStyleSheet(style)
            self.apply_background(self.palette().color(QtGui.QPalette.Window))
            apply_neon_effect(self.lbl_month, True)
        else:
            if CONFIG.get("theme", "dark") == "dark":
                style = self.default_style
            else:
                style = (
                    "QLabel{color:#000; border:none;} "
                    "QToolButton{color:#000; border:1px solid #999; border-radius:6px; padding:4px 10px;}"
                )
            self.setStyleSheet(style)
            self.apply_background(self.palette().color(QtGui.QPalette.Window))
            apply_neon_effect(self.lbl_month, False)
        apply_neon_effect(self.btn_prev, neon)
        apply_neon_effect(self.btn_next, neon)
        self.apply_fonts()
        update_neon_filters(self)


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("План-график")
        self.setWindowIcon(QtGui.QIcon(CONFIG.get("app_icon", ICON_TOGGLE)))
        central = QtWidgets.QWidget(self)
        h = QtWidgets.QHBoxLayout(central); h.setContentsMargins(0,0,0,0); h.setSpacing(0)

        # left sidebar
        self.sidebar=CollapsibleSidebar(self); h.addWidget(self.sidebar)
        self.sidebar.btn_tops.clicked.connect(self.open_top_dialog)

        # right: vbox with topbar + table
        right = QtWidgets.QWidget(self); v = QtWidgets.QVBoxLayout(right); v.setContentsMargins(0,0,0,0); v.setSpacing(0)
        self.topbar = TopBar(self); v.addWidget(self.topbar)
        self.table = ExcelCalendarTable(self); v.addWidget(self.table, 1)
        h.addWidget(right, 1)

        self.setCentralWidget(central)

        self._settings = QtCore.QSettings("rabota2", "rabota2")
        cols = self._settings.value("MainWindow/columns")
        if cols:
            self.table.set_day_column_widths([int(w) for w in cols])

        # Connect topbar
        self.topbar.prev_clicked.connect(self.prev_month)
        self.topbar.next_clicked.connect(self.next_month)
        self.topbar.year_changed.connect(self.change_year)
        if self.sidebar.btn_inputs is not None:
            self.sidebar.btn_inputs.clicked.connect(self.open_input_dialog)
        else:
            logger.error("sidebar.btn_inputs is missing; input dialog feature disabled")
            def _disabled_input_dialog(*_, **__):
                logger.error("Input dialog feature is disabled because btn_inputs is missing")
            self.open_input_dialog = _disabled_input_dialog  # type: ignore[assignment]

        if self.sidebar.btn_release is not None:
            self.sidebar.btn_release.clicked.connect(self.open_release_dialog)
        else:
            logger.error("sidebar.btn_release is missing; release dialog feature disabled")
            def _disabled_release_dialog(*_, **__):
                logger.error("Release dialog feature is disabled because btn_release is missing")
            self.open_release_dialog = _disabled_release_dialog  # type: ignore[assignment]
        self.sidebar.btn_analytics.clicked.connect(self.open_analytics_dialog)
        self.sidebar.settings_clicked.connect(self.open_settings_dialog)
        self._update_month_label()

        # status bar with timer and version info
        self._start_dt = QtCore.QDateTime.currentDateTime()
        bar = QtWidgets.QStatusBar(self)
        self.setStatusBar(bar)
        self.lbl_timer = QtWidgets.QLabel("000:00:00:00", self)
        bar.addWidget(self.lbl_timer)
        self.lbl_version = QtWidgets.QLabel("", self)
        bar.addPermanentWidget(self.lbl_version)
        self._timer = QtCore.QTimer(self)
        self._timer.timeout.connect(self._update_timer)
        self._timer.start(1000)
        self._update_timer()
        self._update_version()

    def _update_month_label(self):
        self.topbar.lbl_month.setText(RU_MONTHS[self.table.month-1])
        self.topbar.spin_year.blockSignals(True)
        self.topbar.spin_year.setValue(self.table.year)
        self.topbar.spin_year.blockSignals(False)
        self.topbar.update_labels()

    def _update_timer(self):
        secs = self._start_dt.secsTo(QtCore.QDateTime.currentDateTime())
        days = secs // 86400
        hours = (secs // 3600) % 24
        minutes = (secs // 60) % 60
        seconds = secs % 60
        self.lbl_timer.setText(f"{days:02d}:{hours:02d}:{minutes:02d}:{seconds:02d}")

    def _update_version(self):
        version = "unknown"
        try:
            with open(VERSION_FILE, "r", encoding="utf-8") as f:
                version = f.read().strip()
        except Exception:
            pass
        self.lbl_version.setText(f"v{version}")

    def prev_month(self):
        self.table.go_prev_month(); self._update_month_label()

    def next_month(self):
        self.table.go_next_month(); self._update_month_label()

    def change_year(self, year):
        self.table.save_current_month()
        self.table.year = year
        self.table.load_month_data(year, self.table.month)
        self._update_month_label()

    def open_input_dialog(self):
        dlg = StatsDialog(self.table.year, self.table.month, self)
        dlg.exec()
        self.sidebar.activate_button(None)
        for b in self.sidebar.buttons:
            apply_neon_effect(b, False)

    def open_analytics_dialog(self):
        dlg = AnalyticsDialog(self.table.year, self)
        dlg.exec()
        self.sidebar.activate_button(None)
        for b in self.sidebar.buttons:
            apply_neon_effect(b, False)

    def _collect_work_names(self) -> List[str]:
        names = set()
        for (r, c), day in self.table.date_map.items():
            if day.month != self.table.month:
                continue
            inner = self.table.cell_tables.get((r, c))
            if not inner:
                continue
            for row in range(inner.rowCount()):
                it = inner.item(row, 0)
                if it:
                    name = it.text().strip()
                    if name:
                        names.add(name)
        return sorted(names)

    def open_release_dialog(self):
        works = self._collect_work_names()
        dlg = ReleaseDialog(self.table.year, self.table.month, works, self)
        dlg.exec()
        self.sidebar.activate_button(None)
        for b in self.sidebar.buttons:
            apply_neon_effect(b, False)

    def open_top_dialog(self):
        dlg = TopDialog(self.table.year, self)
        dlg.exec()
        self.sidebar.activate_button(None)
        for b in self.sidebar.buttons:
            apply_neon_effect(b, False)

    def open_settings_dialog(self):
        dlg = SettingsDialog(self)
        dlg.settings_changed.connect(self._on_settings_changed)
        dlg.exec()
        self.sidebar.activate_button(None)
        for b in self.sidebar.buttons:
            apply_neon_effect(b, False)

    def _on_settings_changed(self):
        global CONFIG, BASE_SAVE_PATH
        CONFIG = load_config()
        if not isinstance(CONFIG.get("gradient_colors"), list):
            CONFIG["gradient_colors"] = ["#39ff14", "#2d7cdb"]
        BASE_SAVE_PATH = os.path.abspath(CONFIG.get("save_path", DATA_DIR))
        self.apply_settings()
        workspace = QtGui.QColor(CONFIG.get("workspace_color", "#1e1e21"))
        self.topbar.apply_background(workspace)

    def apply_fonts(self):
        header_family, text_family = resolve_font_config(self)
        theme_manager.set_text_font(text_family)
        theme_manager.set_header_font(header_family)
        app = QtWidgets.QApplication.instance()
        for w in app.allWidgets():
            if w is self.sidebar or self.sidebar.isAncestorOf(w):
                continue
            w.setFont(app.font())
        self.sidebar.apply_fonts()
        header_font = QtGui.QFont(header_family)
        self.topbar.lbl_month.setFont(header_font)
        self.table.setFont(app.font())
        self.table.horizontalHeader().setFont(header_font)
        for tbl in self.table.cell_tables.values():
            tbl.setFont(app.font())
            tbl.horizontalHeader().setFont(header_font)
        for lbl in self.table.day_labels.values():
            lbl.setFont(header_font)
        # Ensure weekday headers and day labels adopt the selected font
        # after direct font assignments above.
        self.table.apply_fonts()
        self.table.update_day_rows()
        for dlg in app.topLevelWidgets():
            if isinstance(dlg, QtWidgets.QDialog):
                for tbl in dlg.findChildren(QtWidgets.QTableWidget):
                    tbl.setFont(app.font())
                    tbl.horizontalHeader().setFont(header_font)

    def apply_palette(self):
        load_icons(CONFIG.get("theme", "dark"))
        self.topbar.update_icons()
        self.sidebar.update_icons()
        self.setWindowIcon(QtGui.QIcon(CONFIG.get("app_icon", ICON_TOGGLE)))
        workspace = QtGui.QColor(CONFIG.get("workspace_color", "#1e1e21"))
        if CONFIG.get("monochrome", False):
            workspace = theme_manager.apply_monochrome(workspace)
        self.statusBar().setStyleSheet(
            f"background-color:{workspace.name()};"
        )
        self.topbar.apply_background(workspace)
        self.table.horizontalHeader().setStyleSheet(
            f"background-color:{workspace.name()};"
        )
        accent = QtGui.QColor(CONFIG.get("accent_color", "#39ff14"))
        if CONFIG.get("monochrome", False):
            h, s, v, _ = accent.getHsv()
            s = int(CONFIG.get("mono_saturation", 100))
            accent.setHsv(h, s, v)
        # ensure widgets use the latest accent color for borders and neon effects
        app = QtWidgets.QApplication.instance()
        pal = app.palette()
        pal.setColor(QtGui.QPalette.Highlight, accent)
        app.setPalette(pal)
        sidebar_color = CONFIG.get("sidebar_color", "#1f1f23")
        self.topbar.apply_style(CONFIG.get("neon", False))
        self.sidebar.apply_style(CONFIG.get("neon", False), accent, sidebar_color)
        self.sidebar.anim.setDuration(160)
        update_neon_filters(self)
        # Reapply background so the spin box border matches the current theme
        self.topbar.apply_background(workspace)
        self.topbar.update_labels()

    def apply_settings(self):
        self.apply_fonts()
        self.apply_palette()
        self.apply_theme()
        accent = QtGui.QColor(CONFIG.get("accent_color", "#39ff14"))
        sidebar_color = CONFIG.get("sidebar_color", "#1f1f23")
        self.sidebar.apply_style(
            CONFIG.get("neon", False), accent, sidebar_color
        )
        app = QtWidgets.QApplication.instance()
        for w in app.allWidgets():
            if isinstance(w, StyledToolButton):
                w.apply_base_style()
                w._neon_prev_style = w.styleSheet()
                apply_neon_effect(
                    w, w.property("neon_selected") and neon_enabled()
                )



    def apply_style(self, neon: bool):
        accent = QtGui.QColor(CONFIG.get("accent_color", "#39ff14"))
        workspace = QtGui.QColor(CONFIG.get("workspace_color", "#1e1e21"))
        if CONFIG.get("monochrome", False):
            h, s, v, _ = accent.getHsv()
            s = int(CONFIG.get("mono_saturation", 100))
            accent.setHsv(h, s, v)
            workspace = theme_manager.apply_monochrome(workspace)
        sidebar = CONFIG.get("sidebar_color", "#1f1f23")
        border = accent.name() if neon else "#555"
        style = (
            f"background:{workspace.name()};" f"border:1px solid {border};" f"border-radius:8px;"
        )
        for cls in (
            QtWidgets.QLineEdit,
            QtWidgets.QComboBox,
            QtWidgets.QSpinBox,
            QtWidgets.QTimeEdit,
        ):
            for w in self.findChildren(cls):
                w.setStyleSheet(style)
        app = QtWidgets.QApplication.instance()
        pal = app.palette()
        pal.setColor(QtGui.QPalette.Highlight, accent)
        app.setPalette(pal)
        self.topbar.apply_style(neon)
        self.sidebar.apply_style(neon, accent, sidebar)
        update_neon_filters(self)

    def apply_theme(self):
        accent = QtGui.QColor(CONFIG.get("accent_color", "#39ff14"))
        if CONFIG.get("monochrome", False):
            accent = theme_manager.apply_monochrome(accent)
        app = QtWidgets.QApplication.instance()
        def mc(c):
            return theme_manager.apply_monochrome(QtGui.QColor(c)) if CONFIG.get("monochrome", False) else QtGui.QColor(c)
        if CONFIG.get("theme", "dark") == "dark":
            pal = QtGui.QPalette()
            pal.setColor(QtGui.QPalette.Window, mc(QtGui.QColor(30, 30, 33)))
            pal.setColor(QtGui.QPalette.WindowText, mc("#e5e5e5"))
            pal.setColor(QtGui.QPalette.Base, mc(QtGui.QColor(30, 30, 33)))
            pal.setColor(QtGui.QPalette.AlternateBase, mc(QtGui.QColor(45, 45, 48)))
            pal.setColor(QtGui.QPalette.ToolTipBase, mc("#e5e5e5"))
            pal.setColor(QtGui.QPalette.ToolTipText, mc("#e5e5e5"))
            pal.setColor(QtGui.QPalette.Text, mc("#e5e5e5"))
            pal.setColor(QtGui.QPalette.Button, mc(QtGui.QColor(45, 45, 48)))
            pal.setColor(QtGui.QPalette.ButtonText, mc("#e5e5e5"))
            pal.setColor(QtGui.QPalette.Highlight, accent)
            pal.setColor(QtGui.QPalette.HighlightedText, mc(QtGui.QColor(0, 0, 0)))
            app.setPalette(pal)
        else:
            app.setPalette(app.style().standardPalette())
        base, workspace = theme_manager.apply_gradient(CONFIG)
        workspace = QtGui.QColor(workspace)
        self.statusBar().setStyleSheet(
            f"background-color:{workspace.name()};"
        )
        flat_cfg = CONFIG.copy()
        flat_cfg.pop("gradient_colors", None)
        flat_base, _ = theme_manager.apply_gradient(flat_cfg)
        self.topbar.setStyleSheet(
            f"background-color:{workspace.name()};" + self.topbar.styleSheet()
        )
        self.topbar.lbl_month.setStyleSheet(
            "background:transparent; border:none;"
        )
        self.topbar.spin_year.setStyleSheet(
            f"background-color:{workspace.name()}; padding:0 4px; border:1px solid transparent; border-radius:6px;"
        )
        theme = CONFIG.get("theme", "dark")
        self.setStyleSheet(
            "QPushButton,QToolButton{" + base + "}"
            "QPushButton:hover,QToolButton:hover{" + flat_base + "}"
            "QSpinBox,QDoubleSpinBox,QTimeEdit,QComboBox,QLineEdit{" + flat_base + "}"
            f"""
            QSpinBox::up-button,QDoubleSpinBox::up-button{{
                subcontrol-origin:border;
                subcontrol-position:right top;
                margin-left:2px;
                border:1px solid transparent;
                border-radius:8px;
                width:16px;height:16px;
            }}
            QSpinBox::down-button,QDoubleSpinBox::down-button{{
                subcontrol-origin:border;
                subcontrol-position:right bottom;
                margin-left:2px;
                border:1px solid transparent;
                border-radius:8px;
                width:16px;height:16px;
            }}
            QSpinBox::up-arrow,QSpinBox::down-arrow,QDoubleSpinBox::up-arrow,QDoubleSpinBox::down-arrow{{
                width:10px;height:10px;
            }}
            QSpinBox::up-arrow,QDoubleSpinBox::up-arrow{{ image:url(assets/icons/{theme}/chevron-up.svg); }}
            QSpinBox::down-arrow,QDoubleSpinBox::down-arrow{{ image:url(assets/icons/{theme}/chevron-down.svg); }}
            QComboBox::down-arrow{{
                image:url(assets/icons/{theme}/chevron-down.svg);
                width:10px;height:10px;
            }}
            QComboBox::drop-down{{
                subcontrol-origin:padding;
                subcontrol-position:right center;
                width:16px;
            }}
            """
        )
        self.table.apply_theme()
        sidebar = CONFIG.get("sidebar_color", "#1f1f23")
        if CONFIG.get("monochrome", False):
            sidebar = theme_manager.apply_monochrome(QtGui.QColor(sidebar)).name()
            accent = theme_manager.apply_monochrome(accent)
        self.sidebar.apply_style(CONFIG.get("neon", False), accent, sidebar)
        # Reapply topbar background so spin box picks up workspace color
        self.topbar.apply_background(workspace)

    def resizeEvent(self, event):
        super().resizeEvent(event)

    def closeEvent(self, event):
        self.table.save_current_month()
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(CONFIG, f, ensure_ascii=False, indent=2)
        cols = self.table.get_day_column_widths()
        self._settings.setValue("MainWindow/columns", cols)
        super().closeEvent(event)


def main():
    load_icons(CONFIG.get("theme", "dark"))

    # Ensure the configured base font is available before applying it globally
    base_family = ensure_font_registered(CONFIG.get("font_family", "Exo 2"))
    CONFIG["font_family"] = base_family

    header_family, text_family = resolve_font_config()
    theme_manager.set_header_font(header_family)
    theme_manager.set_text_font(text_family)

    w = MainWindow()
    w.apply_settings()
    w.show()
    w.showMaximized()
    return w

if __name__ == "__main__":
    QtCore.QLocale.setDefault(QtCore.QLocale("ru_RU"))
    app = QtWidgets.QApplication(sys.argv)
    try:
        register_fonts()
        window = main()
        exit_code = app.exec()
        sys.exit(exit_code)
    except Exception as exc:
        logging.exception("Unhandled exception in application")
        QtWidgets.QMessageBox.critical(
            None, "Ошибка", f"{type(exc).__name__}: {exc}"
        )
        sys.exit(1)
