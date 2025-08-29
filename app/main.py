# -*- coding: utf-8 -*-
import sys, os, json, calendar, re
from datetime import datetime, date
from typing import Dict, List, Union

from PySide6 import QtWidgets, QtGui, QtCore
from dataclasses import dataclass, field

from widgets import StyledPushButton, StyledToolButton
from resources import register_fonts, load_icons, icon
import theme_manager
from effects import set_neon

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
        "monochrome": False,
        "mono_saturation": 100,
        "glass_effect": "Acrylic",
        "glass_opacity": 0.5,
        "glass_blur": 10,
        "animation_speed": 1.0,
        "neon_motion": False,
        "theme": "dark",
        "header_font": "Exo 2",
        "text_font": "Inter",
        "save_path": DATA_DIR,
        "day_rows": 6,
        "workspace_color": "#1e1e21",
        "sidebar_color": "#1f1f23",
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


class NeonEventFilter(QtCore.QObject):
    """Event filter enabling animated neon highlight on hover and focus."""

    def __init__(self, widget: QtWidgets.QWidget):
        super().__init__(widget)
        self._widget = widget
        self._anim = None
        self._motion = None
        self._orig_style = widget.styleSheet()

    def _start(self) -> None:
        if not CONFIG.get("neon", False):
            return
        accent = QtGui.QColor(CONFIG.get("accent_color", "#39ff14"))
        self._orig_style = self._widget.styleSheet()
        style = re.sub(r"border-color:[^;]+;?", "", self._orig_style)
        self._widget.setStyleSheet(style + f"border-color:{accent.name()};")
        if isinstance(self._widget, QtWidgets.QAbstractButton):
            eff, anim, motion = set_neon(
                self._widget,
                accent,
                intensity=CONFIG.get("neon_intensity", 255),
                pulse=CONFIG.get("neon_motion", False),
                motion_speed=CONFIG.get("animation_speed", 1.0),
                mode="inner",
            )
            self._anim = anim
            self._motion = motion
            self._effect = eff
        else:
            self._widget.setGraphicsEffect(None)
            self._anim = None
            self._motion = None
            self._effect = None

    def _stop(self) -> None:
        self._widget.setStyleSheet(self._orig_style)
        self._widget.setGraphicsEffect(None)
        if self._anim:
            self._anim.stop()
            self._anim = None
        if self._motion:
            self._motion.stop()
            self._motion = None

    def eventFilter(self, obj, ev):
        if ev.type() in (QtCore.QEvent.HoverEnter, QtCore.QEvent.FocusIn):
            self._start()
        elif isinstance(self._widget, QtWidgets.QToolButton):
            if ev.type() in (QtCore.QEvent.Leave, QtCore.QEvent.HoverLeave):
                if self._widget.property("neon_selected"):
                    self._widget.setStyleSheet(self._orig_style)
                else:
                    self._stop()
            elif ev.type() == QtCore.QEvent.FocusOut:
                if not self._widget.property("neon_selected"):
                    self._stop()
        elif ev.type() in (
            QtCore.QEvent.HoverLeave,
            QtCore.QEvent.Leave,
            QtCore.QEvent.FocusOut,
        ):
            self._stop()
        return False


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
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table.setStyleSheet(
            "QTableWidget{border:1px solid #555; border-radius:8px;} "
            "QTableWidget::item{border:0;}"
        )
        self.table.setRowCount(self.days_in_month)

        app = QtWidgets.QApplication.instance()
        self.setFont(app.font())
        self.table.setFont(app.font())
        header_font = QtGui.QFont(CONFIG.get("header_font"))
        self.table.horizontalHeader().setFont(header_font)

        for row in range(self.days_in_month):
            sp_day = QtWidgets.QSpinBox(self.table)
            sp_day.setRange(1, self.days_in_month)
            sp_day.setValue(row + 1)
            sp_day.setReadOnly(True)
            sp_day.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
            sp_day.setMinimumWidth(sp_day.sizeHint().width())
            self.table.setCellWidget(row, 0, sp_day)
            sp_day.setAttribute(QtCore.Qt.WA_Hover, True)
            sp_day.installEventFilter(NeonEventFilter(sp_day))

            cb_work = QtWidgets.QComboBox(self.table)
            cb_work.setEditable(True)
            cb_work.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToContents)
            cb_work.setMinimumWidth(cb_work.sizeHint().width())
            self.table.setCellWidget(row, 1, cb_work)
            cb_work.setAttribute(QtCore.Qt.WA_Hover, True)
            cb_work.installEventFilter(NeonEventFilter(cb_work))

            sp_ch = QtWidgets.QSpinBox(self.table)
            sp_ch.setRange(0, 9999)
            sp_ch.setMinimumWidth(sp_ch.sizeHint().width())
            self.table.setCellWidget(row, 2, sp_ch)
            sp_ch.setAttribute(QtCore.Qt.WA_Hover, True)
            sp_ch.installEventFilter(NeonEventFilter(sp_ch))

            te_time = QtWidgets.QTimeEdit(self.table)
            te_time.setDisplayFormat("HH:mm")
            te_time.setTime(QtCore.QTime.currentTime())
            te_time.setMinimumWidth(te_time.sizeHint().width())
            self.table.setCellWidget(row, 3, te_time)
            te_time.setAttribute(QtCore.Qt.WA_Hover, True)
            te_time.installEventFilter(NeonEventFilter(te_time))

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
        box.accepted.connect(self.save)
        box.rejected.connect(self.reject)
        lay.addWidget(box)

        for b in (btn_save, btn_close):
            b.setAttribute(QtCore.Qt.WA_Hover, True)
            b.installEventFilter(NeonEventFilter(b))

        self.load()

    def file_path(self):
        return os.path.join(release_dir(self.year), f"{self.month:02d}.json")

    def load(self):
        path = self.file_path()
        data = {}
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)

        file_works = data.get("works", [])
        for w in file_works:
            if w not in self.works:
                self.works.append(w)

        for row in range(self.days_in_month):
            cb = self.table.cellWidget(row, 1)
            if isinstance(cb, QtWidgets.QComboBox):
                cb.clear()
                cb.addItems(self.works)
                cb.setCurrentIndex(-1)
                cb.setMinimumWidth(cb.sizeHint().width())

        for day_str, entries in data.get("days", {}).items():
            day = int(day_str)
            if 1 <= day <= self.days_in_month and entries:
                entry = entries[0]
                cb_work = self.table.cellWidget(day - 1, 1)
                sp_ch = self.table.cellWidget(day - 1, 2)
                te_time = self.table.cellWidget(day - 1, 3)
                work = entry.get("work", "")
                if isinstance(cb_work, QtWidgets.QComboBox):
                    if work and cb_work.findText(work) < 0:
                        cb_work.addItem(work)
                    cb_work.setCurrentText(work)
                    cb_work.setMinimumWidth(cb_work.sizeHint().width())
                if isinstance(sp_ch, QtWidgets.QSpinBox):
                    sp_ch.setValue(entry.get("chapters", 0))
                if isinstance(te_time, QtWidgets.QTimeEdit):
                    qt = QtCore.QTime.fromString(entry.get("time", "00:00"), "HH:mm")
                    if qt.isValid():
                        te_time.setTime(qt)

    def save(self):
        days: Dict[str, List[Dict[str, str | int]]] = {}
        for row in range(self.days_in_month):
            cb_work = self.table.cellWidget(row, 1)
            sp_ch = self.table.cellWidget(row, 2)
            te_time = self.table.cellWidget(row, 3)
            if not (cb_work and sp_ch and te_time):
                continue
            work_name = cb_work.currentText().strip()
            if not work_name:
                continue
            day = row + 1
            entry = {
                "work": work_name,
                "chapters": sp_ch.value(),
                "time": te_time.time().toString("HH:mm"),
            }
            days[str(day)] = [entry]

        works = sorted({e["work"] for entries in days.values() for e in entries})
        data = {"works": works, "days": days}

        os.makedirs(os.path.dirname(self.file_path()), exist_ok=True)
        with open(self.file_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.accept()


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
            if isinstance(w, QtWidgets.QDoubleSpinBox):
                w.setRange(0, 1_000_000_000)
                w.setDecimals(2)
            if key == "progress" and isinstance(w, QtWidgets.QDoubleSpinBox):
                w.setRange(0, 100)
                w.setSuffix("%")
            if isinstance(w, QtWidgets.QComboBox):
                w.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToContents)
            w.setMinimumWidth(w.sizeHint().width())
            form.addRow(label, w)
            self.widgets[key] = w
            w.setAttribute(QtCore.Qt.WA_Hover, True)
            w.installEventFilter(NeonEventFilter(w))

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
            "QHeaderView::section{padding:0 6px;}"
        )
        header = self.table_stats.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        header.setTextElideMode(QtCore.Qt.ElideNone)
        self.table_stats.setHorizontalHeaderLabels(
            [h for _, h in StatsEntryForm.TABLE_COLUMNS]
        )
        self.table_stats.setSortingEnabled(True)
        self.table_stats.verticalHeader().setVisible(False)
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
        btn_save.installEventFilter(NeonEventFilter(btn_save))
        btn_close.installEventFilter(NeonEventFilter(btn_close))
        lay.setStretch(0, 2)
        lay.setStretch(1, 1)

        self.records: List[Dict[str, int | float | str | bool]] = []
        self.current_index = None
        self.year = year
        self.month = month
        self.load_stats(year, month)

    def resizeEvent(self, event):
        self.table_stats.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.Stretch
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
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
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
        self.table_stats.resizeRowsToContents()
        header = self.table_stats.horizontalHeader()
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
        self.spin_year.valueChanged.connect(self._year_changed)
        top.addWidget(self.spin_year)
        self.spin_year.installEventFilter(NeonEventFilter(self.spin_year))
        top.addStretch(1)
        lay.addLayout(top)

        cols = len(RU_MONTHS) + 1
        self.table = NeonTableWidget(len(self.INDICATORS), cols, self)
        self.table.setSelectionBehavior(QtWidgets.QTableView.SelectColumns)
        self.table.setSelectionMode(QtWidgets.QTableView.SingleSelection)
        self.table.setStyleSheet(
            "QTableWidget{border:1px solid #555; border-radius:8px;} "
            "QTableWidget::item{border:0;} "
            "QHeaderView::section{padding:0 6px;}"
        )
        self.table.setHorizontalHeaderLabels(RU_MONTHS + ["Итого за год"])
        self.table.setVerticalHeaderLabels(self.INDICATORS)
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
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
        btn_save.installEventFilter(NeonEventFilter(btn_save))
        btn_close.installEventFilter(NeonEventFilter(btn_close))

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
        self.load(year)

    def resizeEvent(self, event):
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
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
        self.table.resizeRowsToContents()
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self._loading = False

    def save(self):
        path = os.path.join(year_dir(self.year), f"{self.year}.json")
        data = {
            "commission": self._commissions,
            "software": self._software,
            "net": self._net,
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.accept()

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
        top = QtWidgets.QHBoxLayout()
        top.addWidget(QtWidgets.QLabel("Год:"))
        self.spin_year = QtWidgets.QSpinBox(self)
        self.spin_year.setRange(2000, 2100)
        self.spin_year.setValue(year)
        top.addWidget(self.spin_year)
        self.spin_year.installEventFilter(NeonEventFilter(self.spin_year))

        self.combo_mode = QtWidgets.QComboBox(self)
        self.combo_mode.addItem("Месяц", "month")
        self.combo_mode.addItem("Квартал", "quarter")
        self.combo_mode.addItem("Полугодие", "half")
        self.combo_mode.addItem("Год", "year")
        self.combo_mode.currentIndexChanged.connect(self._mode_changed)
        top.addWidget(self.combo_mode)
        self.combo_mode.installEventFilter(NeonEventFilter(self.combo_mode))

        self.combo_period = QtWidgets.QComboBox(self)
        top.addWidget(self.combo_period)
        self.combo_period.installEventFilter(NeonEventFilter(self.combo_period))
        self._mode_changed()  # fill periods

        self.btn_calc = StyledPushButton("Сформировать", self)
        self.btn_calc.clicked.connect(self.calculate)
        top.addWidget(self.btn_calc)
        self.btn_calc.installEventFilter(NeonEventFilter(self.btn_calc))
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
        self.table.horizontalHeader().setStretchLastSection(True)
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
        box.accepted.connect(self.save)
        box.rejected.connect(self.reject)
        lay.addWidget(box)
        btn_save.installEventFilter(NeonEventFilter(btn_save))
        btn_close.installEventFilter(NeonEventFilter(btn_close))

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
        self.accept()

class NeonTableWidget(QtWidgets.QTableWidget):
    """Вложенная таблица с неоновым подсвечиванием при наведении и фокусе."""

    def __init__(self, rows, cols, parent=None):
        super().__init__(rows, cols, parent)
        self.setFocusPolicy(QtCore.Qt.StrongFocus)
        self.setAttribute(QtCore.Qt.WA_Hover, True)
        self.viewport().setAttribute(QtCore.Qt.WA_Hover, True)
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

        self.load_month_data(self.year, self.month)

    def mousePressEvent(self, event):
        self.clearSelection()
        event.ignore()

    def _create_inner_table(self) -> QtWidgets.QTableWidget:
        tbl = NeonTableWidget(CONFIG.get("day_rows", 6), 3, self)
        tbl.setHorizontalHeaderLabels(["Работа", "План", "Готово"])
        tbl.verticalHeader().setVisible(False)
        tbl.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        tbl.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked)
        return tbl

    def update_day_rows(self):
        rows = CONFIG.get("day_rows", 6)
        for tbl in self.cell_tables.values():
            tbl.setRowCount(rows)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_row_heights()

    def _update_row_heights(self):
        rows = self.rowCount()
        if rows:
            height = self.viewport().height() // rows
            for r in range(rows):
                self.setRowHeight(r, height)

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
        self.setRowCount(len(weeks))
        for r, week in enumerate(weeks):
            for c, day in enumerate(week):
                container = QtWidgets.QWidget()
                lay = QtWidgets.QVBoxLayout(container)
                lay.setContentsMargins(0, 0, 0, 0)
                lay.setSpacing(2)
                lbl = QtWidgets.QLabel(str(day.day), container)
                lbl.setFont(QtGui.QFont(CONFIG.get("header_font", "Exo 2")))
                lbl.setAlignment(QtCore.Qt.AlignCenter)
                lay.addWidget(lbl, alignment=QtCore.Qt.AlignHCenter)
                inner = self._create_inner_table()
                lay.addWidget(inner)
                self.setCellWidget(r, c, container)
                self.date_map[(r, c)] = day
                self.cell_tables[(r, c)] = inner
                self.day_labels[(r, c)] = lbl
                self.cell_containers[(r, c)] = container
                if day.month != month:
                    container.setEnabled(False)
                    container.setStyleSheet("background-color:#2a2a2a; color:#777;")
                    lbl.setStyleSheet("color:#777;")
                else:
                    container.setEnabled(True)
                    container.setStyleSheet("")
                    lbl.setStyleSheet("")
                    rows = md.days.get(day.day, [])
                    for rr, row in enumerate(rows):
                        if rr >= inner.rowCount():
                            break
                        for cc, key in enumerate(["work", "plan", "done"]):
                            item = QtWidgets.QTableWidgetItem(str(row.get(key, "")))
                            inner.setItem(rr, cc, item)
        self._update_row_heights()
        return True

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
        self.setStyleSheet("""
            #Sidebar { background-color: #1f1f23; }
            QToolButton { color: white; border: none; padding: 10px; border-radius: 8px; }
            QToolButton:hover { background-color: rgba(255,255,255,0.08); }
            QLabel { color: #c7c7c7; }
        """)
        self.expanded_width=260; self.collapsed_width=64
        lay=QtWidgets.QVBoxLayout(self); lay.setContentsMargins(8,8,8,8); lay.setSpacing(6)

        # Toggle button — только иконка
        self.btn_toggle = StyledToolButton(self)
        self.btn_toggle.setIcon(QtGui.QIcon(ICON_TOGGLE))
        self.btn_toggle.setIconSize(QtCore.QSize(28,28))
        self.btn_toggle.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly)
        self.btn_toggle.setCursor(QtCore.Qt.PointingHandCursor)
        self.btn_toggle.clicked.connect(self.toggle)
        self.btn_toggle.setAttribute(QtCore.Qt.WA_Hover, True)
        lay.addWidget(self.btn_toggle)
        self.btn_toggle.installEventFilter(NeonEventFilter(self.btn_toggle))

        line = QtWidgets.QFrame(self); line.setFrameShape(QtWidgets.QFrame.HLine); line.setStyleSheet("color:#333;")
        lay.addWidget(line)

        items = [
            ("Вводные", ICON_TM),
            ("Выкладка", ICON_TQ),
            ("Аналитика", ICON_TG),
            ("Топы", ICON_TP),
        ]
        self.buttons = []
        self._filters = {}
        self.btn_inputs = None
        self.btn_release = None
        self.btn_analytics = None
        self.btn_tops = None
        for label, icon in items:
            b = StyledToolButton(self)
            b.setIcon(QtGui.QIcon(icon))
            b.setIconSize(QtCore.QSize(22, 22))
            b.setText(" " + label)
            b.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
            b.setCursor(QtCore.Qt.PointingHandCursor)
            b.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
            b.setAttribute(QtCore.Qt.WA_Hover, True)
            b.setProperty("neon_selected", False)
            lay.addWidget(b)
            self.buttons.append(b)
            filt = NeonEventFilter(b)
            b.installEventFilter(filt)
            self._filters[b] = filt
            b.clicked.connect(lambda _, btn=b: self.activate_button(btn))
            if label == "Вводные":
                self.btn_inputs = b
            elif label == "Выкладка":
                self.btn_release = b
            elif label == "Аналитика":
                self.btn_analytics = b
            elif label == "Топы":
                self.btn_tops = b
        # Stretch removed since there is no settings button at the bottom

        self._collapsed = False
        self.anim = QtCore.QPropertyAnimation(self, b"maximumWidth", self)
        self.anim.setDuration(160)
        self.setMinimumWidth(self.collapsed_width)
        self.setMaximumWidth(self.expanded_width)
        self.update_icons()

    def activate_button(self, btn: QtWidgets.QToolButton) -> None:
        for b, filt in self._filters.items():
            if b is btn:
                b.setProperty("neon_selected", True)
                filt._start()
            else:
                b.setProperty("neon_selected", False)
                filt._stop()

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

    def apply_style(self, neon: bool, accent: QtGui.QColor, sidebar_color: Union[str, QtGui.QColor, None] = None):
        thickness = CONFIG.get("neon_thickness", 1)
        size = CONFIG.get("neon_size", 10)
        intensity = CONFIG.get("neon_intensity", 255)
        if sidebar_color is None:
            sidebar_color = CONFIG.get("sidebar_color", "#1f1f23")
        if isinstance(sidebar_color, QtGui.QColor):
            sidebar_color = sidebar_color.name()
        speed = CONFIG.get("animation_speed", 1.0)
        if neon:
            style = (
                f"#Sidebar {{ background-color: {sidebar_color}; }}"
                f"QToolButton {{ color: {accent.name()}; border:{thickness}px solid {accent.name()}; padding: 10px; border-radius: 8px; }}"
                "QToolButton:hover { background-color: rgba(255,255,255,0.08); }"
                f"QLabel {{ color: {accent.name()}; }}"
            )
            self.setStyleSheet(style)
            widgets = [self.btn_toggle] + self.buttons
            for w in widgets:
                eff = QtWidgets.QGraphicsDropShadowEffect(self)
                eff.setOffset(0, 0)
                eff.setBlurRadius(size)
                c = QtGui.QColor(accent)
                c.setAlpha(intensity)
                eff.setColor(c)
                w.setGraphicsEffect(eff)
                if CONFIG.get("neon_motion", False):
                    anim = QtCore.QPropertyAnimation(eff, b"blurRadius", w)
                    anim.setStartValue(size)
                    anim.setEndValue(size * 2)
                    anim.setDuration(int(1000 / max(speed, 0.1)))
                    anim.setLoopCount(-1)
                    anim.setEasingCurve(QtCore.QEasingCurve.InOutQuad)
                    anim.start(QtCore.QAbstractAnimation.DeleteWhenStopped)
                    w._neon_anim = anim
                else:
                    w._neon_anim = None
        else:
            style = (
                f"#Sidebar {{ background-color: {sidebar_color}; }}\n"
                "QToolButton { color: white; border: none; padding: 10px; border-radius: 8px; }\n"
                "QToolButton:hover { background-color: rgba(255,255,255,0.08); }\n"
                "QLabel { color: #c7c7c7; }\n"
            )
            self.setStyleSheet(style)
            for w in [self.btn_toggle] + self.buttons:
                if hasattr(w, "_neon_anim") and w._neon_anim:
                    w._neon_anim.stop()
                    w._neon_anim = None
                w.setGraphicsEffect(None)

    def update_icons(self) -> None:
        """Update sidebar icons. Settings button was removed."""
        pass


class SettingsDialog(QtWidgets.QDialog):
    """Окно настроек приложения."""

    settings_changed = QtCore.Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки")
        self.resize(500, 400)
        main_lay = QtWidgets.QVBoxLayout(self)
        tabs = QtWidgets.QTabWidget(self)
        main_lay.addWidget(tabs)

        # --- Цвет и тема ---
        tab_color = QtWidgets.QWidget()
        form_color = QtWidgets.QFormLayout(tab_color)
        tabs.addTab(tab_color, "Цвет и тема")

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
        self.combo_accent.currentIndexChanged.connect(self._on_accent_changed)
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
        form_color.addRow("Цвет подсветки", self.combo_accent)

        self._workspace_color = QtGui.QColor(CONFIG.get("workspace_color", "#1e1e21"))
        self.btn_workspace = StyledPushButton(self)
        self._update_workspace_button()
        self.btn_workspace.clicked.connect(self.choose_workspace_color)
        form_color.addRow("Цвет рабочей области", self.btn_workspace)

        self._sidebar_color = QtGui.QColor(CONFIG.get("sidebar_color", "#1f1f23"))
        self.btn_sidebar = StyledPushButton(self)
        self._update_sidebar_button()
        self.btn_sidebar.clicked.connect(self.choose_sidebar_color)
        form_color.addRow("Цвет боковой панели", self.btn_sidebar)

        self.combo_theme = QtWidgets.QComboBox(self)
        self.combo_theme.addItems(["Темная", "Светлая"])
        self.combo_theme.setCurrentIndex(0 if CONFIG.get("theme","dark")=="dark" else 1)
        self.combo_theme.currentIndexChanged.connect(lambda _: self._save_config())
        form_color.addRow("Тема", self.combo_theme)

        # gradient controls
        grad = CONFIG.get("gradient_colors", ["#39ff14", "#2d7cdb"])
        self._grad_color1 = QtGui.QColor(grad[0])
        self._grad_color2 = QtGui.QColor(grad[1])
        self.btn_grad1 = StyledPushButton(self)
        self.btn_grad2 = StyledPushButton(self)
        self._update_grad_buttons()
        self.btn_grad1.clicked.connect(lambda: self.choose_grad_color(1))
        self.btn_grad2.clicked.connect(lambda: self.choose_grad_color(2))
        lay_grad = QtWidgets.QHBoxLayout(); lay_grad.addWidget(self.btn_grad1); lay_grad.addWidget(self.btn_grad2)
        form_color.addRow("Градиент", lay_grad)
        self.sld_grad_angle = QtWidgets.QSlider(QtCore.Qt.Horizontal, self)
        self.sld_grad_angle.setRange(0, 360)
        self.sld_grad_angle.setValue(int(CONFIG.get("gradient_angle", 0)))
        self.lbl_grad_angle = QtWidgets.QLabel(str(self.sld_grad_angle.value()))
        self.sld_grad_angle.valueChanged.connect(lambda v: (self.lbl_grad_angle.setText(str(v)), self._save_config()))
        lay_angle = QtWidgets.QHBoxLayout(); lay_angle.addWidget(self.sld_grad_angle,1); lay_angle.addWidget(self.lbl_grad_angle)
        form_color.addRow("Угол градиента", lay_angle)

        # monochrome
        self.chk_monochrome = QtWidgets.QCheckBox(self)
        self.chk_monochrome.setChecked(CONFIG.get("monochrome", False))
        self.chk_monochrome.toggled.connect(lambda _: self._save_config())
        form_color.addRow("Монохром", self.chk_monochrome)
        self.sld_mono_sat = QtWidgets.QSlider(QtCore.Qt.Horizontal, self)
        self.sld_mono_sat.setRange(0, 255)
        self.sld_mono_sat.setValue(int(CONFIG.get("mono_saturation", 100)))
        self.lbl_mono_sat = QtWidgets.QLabel(str(self.sld_mono_sat.value()))
        self.sld_mono_sat.valueChanged.connect(lambda v: (self.lbl_mono_sat.setText(str(v)), self._save_config()))
        lay_mono = QtWidgets.QHBoxLayout(); lay_mono.addWidget(self.sld_mono_sat,1); lay_mono.addWidget(self.lbl_mono_sat)
        form_color.addRow("Насыщенность", lay_mono)

        # glass effect
        self.combo_glass = QtWidgets.QComboBox(self)
        self.combo_glass.addItems(["Acrylic", "Mica", "Aero", "Нет"])
        cur_glass = CONFIG.get("glass_effect", "Acrylic")
        idx = self.combo_glass.findText(cur_glass) if cur_glass in ["Acrylic","Mica","Aero"] else 3
        self.combo_glass.setCurrentIndex(idx)
        self.combo_glass.currentIndexChanged.connect(lambda _: self._save_config())
        form_color.addRow("Стекло", self.combo_glass)
        self.sld_glass_opacity = QtWidgets.QSlider(QtCore.Qt.Horizontal, self)
        self.sld_glass_opacity.setRange(0,100)
        self.sld_glass_opacity.setValue(int(CONFIG.get("glass_opacity",0.5)*100))
        self.lbl_glass_opacity = QtWidgets.QLabel(str(self.sld_glass_opacity.value()/100))
        self.sld_glass_opacity.valueChanged.connect(lambda v: (self.lbl_glass_opacity.setText(str(v/100)), self._save_config()))
        lay_op = QtWidgets.QHBoxLayout(); lay_op.addWidget(self.sld_glass_opacity,1); lay_op.addWidget(self.lbl_glass_opacity)
        form_color.addRow("Прозрачность", lay_op)

        self.sld_glass_blur = QtWidgets.QSlider(QtCore.Qt.Horizontal, self)
        self.sld_glass_blur.setRange(0,50)
        self.sld_glass_blur.setValue(int(CONFIG.get("glass_blur",10)))
        self.lbl_glass_blur = QtWidgets.QLabel(str(self.sld_glass_blur.value()))
        self.sld_glass_blur.valueChanged.connect(lambda v: (self.lbl_glass_blur.setText(str(v)), self._save_config()))
        lay_blur = QtWidgets.QHBoxLayout(); lay_blur.addWidget(self.sld_glass_blur,1); lay_blur.addWidget(self.lbl_glass_blur)
        form_color.addRow("Размытие", lay_blur)

        # --- Анимация ---
        tab_anim = QtWidgets.QWidget()
        form_anim = QtWidgets.QFormLayout(tab_anim)
        tabs.addTab(tab_anim, "Анимация")

        self.chk_neon = QtWidgets.QCheckBox(self)
        self.chk_neon.setChecked(CONFIG.get("neon", False))
        self.chk_neon.toggled.connect(lambda _: self._save_config())
        form_anim.addRow("Неоновая подсветка", self.chk_neon)

        self.sld_neon_size = QtWidgets.QSlider(QtCore.Qt.Horizontal, self)
        self.sld_neon_size.setRange(1, 50)
        self.sld_neon_size.setValue(CONFIG.get("neon_size", 10))
        self.lbl_neon_size = QtWidgets.QLabel(str(self.sld_neon_size.value()), self)
        self.sld_neon_size.valueChanged.connect(lambda v: (self.lbl_neon_size.setText(str(v)), self._save_config()))
        lay_size = QtWidgets.QHBoxLayout(); lay_size.addWidget(self.sld_neon_size,1); lay_size.addWidget(self.lbl_neon_size)
        form_anim.addRow("Размер подсветки", lay_size)

        self.sld_neon_thickness = QtWidgets.QSlider(QtCore.Qt.Horizontal, self)
        self.sld_neon_thickness.setRange(1, 10)
        self.sld_neon_thickness.setValue(CONFIG.get("neon_thickness", 1))
        self.lbl_neon_thickness = QtWidgets.QLabel(str(self.sld_neon_thickness.value()), self)
        self.sld_neon_thickness.valueChanged.connect(lambda v: (self.lbl_neon_thickness.setText(str(v)), self._save_config()))
        lay_thick = QtWidgets.QHBoxLayout(); lay_thick.addWidget(self.sld_neon_thickness,1); lay_thick.addWidget(self.lbl_neon_thickness)
        form_anim.addRow("Толщина подсветки", lay_thick)

        self.sld_neon_intensity = QtWidgets.QSlider(QtCore.Qt.Horizontal, self)
        self.sld_neon_intensity.setRange(0,255)
        self.sld_neon_intensity.setValue(CONFIG.get("neon_intensity",255))
        self.lbl_neon_intensity = QtWidgets.QLabel(str(self.sld_neon_intensity.value()), self)
        self.sld_neon_intensity.valueChanged.connect(lambda v: (self.lbl_neon_intensity.setText(str(v)), self._save_config()))
        lay_int = QtWidgets.QHBoxLayout(); lay_int.addWidget(self.sld_neon_intensity,1); lay_int.addWidget(self.lbl_neon_intensity)
        form_anim.addRow("Интенсивность", lay_int)

        self.sld_anim_speed = QtWidgets.QSlider(QtCore.Qt.Horizontal, self)
        self.sld_anim_speed.setRange(50,200)
        self.sld_anim_speed.setValue(int(CONFIG.get("animation_speed",1.0)*100))
        self.lbl_anim_speed = QtWidgets.QLabel(str(self.sld_anim_speed.value()/100))
        self.sld_anim_speed.valueChanged.connect(lambda v: (self.lbl_anim_speed.setText(str(v/100)), self._save_config()))
        lay_spd = QtWidgets.QHBoxLayout(); lay_spd.addWidget(self.sld_anim_speed,1); lay_spd.addWidget(self.lbl_anim_speed)
        form_anim.addRow("Скорость", lay_spd)

        self.chk_neon_motion = QtWidgets.QCheckBox(self)
        self.chk_neon_motion.setChecked(CONFIG.get("neon_motion", False))
        self.chk_neon_motion.toggled.connect(lambda _: self._save_config())
        form_anim.addRow("Движение неона", self.chk_neon_motion)

        # --- Шрифты ---
        tab_fonts = QtWidgets.QWidget()
        form_fonts = QtWidgets.QFormLayout(tab_fonts)
        tabs.addTab(tab_fonts, "Шрифты")
        db = QtGui.QFontDatabase()
        self.font_header = QtWidgets.QComboBox(self)
        self.font_header.addItems(db.families())
        self.font_header.setCurrentText(CONFIG.get("header_font", "Exo 2"))
        self.font_header.currentTextChanged.connect(lambda _: self._save_config())
        form_fonts.addRow("Шрифт заголовков", self.font_header)

        self.font_text = QtWidgets.QComboBox(self)
        self.font_text.addItems(db.families())
        self.font_text.setCurrentText(CONFIG.get("text_font", "Inter"))
        self.font_text.currentTextChanged.connect(lambda _: self._save_config())
        form_fonts.addRow("Шрифт текста", self.font_text)

        # --- Иконки ---
        tab_icons = QtWidgets.QWidget()
        lay_icons = QtWidgets.QVBoxLayout(tab_icons)
        lay_icons.addStretch(1)
        lbl_icons = QtWidgets.QLabel("Нет настроек")
        lbl_icons.setAlignment(QtCore.Qt.AlignCenter)
        lay_icons.addWidget(lbl_icons)
        lay_icons.addStretch(1)
        tabs.addTab(tab_icons, "Иконки")

        # General options below tabs
        form_gen = QtWidgets.QFormLayout()
        main_lay.addLayout(form_gen)
        self.spin_day_rows = QtWidgets.QSpinBox(self)
        self.spin_day_rows.setRange(1, 20)
        self.spin_day_rows.setValue(CONFIG.get("day_rows", 6))
        self.spin_day_rows.valueChanged.connect(lambda _: self._save_config())
        form_gen.addRow("Строк на день", self.spin_day_rows)

        path_lay = QtWidgets.QHBoxLayout()
        self.edit_path = QtWidgets.QLineEdit(CONFIG.get("save_path", DATA_DIR), self)
        self.edit_path.editingFinished.connect(self._save_config)
        btn_browse = StyledPushButton("...", self)
        btn_browse.clicked.connect(self.browse_path)
        path_lay.addWidget(self.edit_path,1)
        path_lay.addWidget(btn_browse)
        form_gen.addRow("Путь сохранения", path_lay)

        box = QtWidgets.QDialogButtonBox(self)
        btn_save = StyledPushButton("Сохранить", self)
        btn_save.setIcon(icon("save"))
        btn_save.setIconSize(QtCore.QSize(20, 20))
        btn_cancel = StyledPushButton("Отмена", self)
        btn_cancel.setIcon(icon("x"))
        btn_cancel.setIconSize(QtCore.QSize(20, 20))
        box.addButton(btn_save, QtWidgets.QDialogButtonBox.AcceptRole)
        box.addButton(btn_cancel, QtWidgets.QDialogButtonBox.RejectRole)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        main_lay.addWidget(box)

    def browse_path(self):
        path = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Выбрать папку", self.edit_path.text()
        )
        if path:
            self.edit_path.setText(path)
            self._save_config()

    def _collect_config(self):
        return {
            "neon": self.chk_neon.isChecked(),
            "neon_size": self.sld_neon_size.value(),
            "neon_thickness": self.sld_neon_thickness.value(),
            "neon_intensity": self.sld_neon_intensity.value(),
            "neon_motion": self.chk_neon_motion.isChecked(),
            "animation_speed": self.sld_anim_speed.value() / 100,
            "accent_color": self._accent_color.name(),
            "gradient_colors": [self._grad_color1.name(), self._grad_color2.name()],
            "gradient_angle": self.sld_grad_angle.value(),
            "monochrome": self.chk_monochrome.isChecked(),
            "mono_saturation": self.sld_mono_sat.value(),
            "glass_effect": "" if self.combo_glass.currentText() == "Нет" else self.combo_glass.currentText(),
            "glass_opacity": self.sld_glass_opacity.value() / 100,
            "glass_blur": self.sld_glass_blur.value(),
            "workspace_color": self._workspace_color.name(),
            "sidebar_color": self._sidebar_color.name(),
            "theme": "dark" if self.combo_theme.currentIndex() == 0 else "light",
            "header_font": self.font_header.currentText(),
            "text_font": self.font_text.currentText(),
            "day_rows": self.spin_day_rows.value(),
            "save_path": self.edit_path.text().strip() or DATA_DIR,
        }

    def _save_config(self):
        config = self._collect_config()
        CONFIG.update(config)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(CONFIG, f, ensure_ascii=False, indent=2)
        self.settings_changed.emit()

    def accept(self):
        self._save_config()
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
        self.btn_workspace.setStyleSheet(
            f"background:{self._workspace_color.name()}; border:1px solid #555;"
        )

    def choose_workspace_color(self):
        color = QtWidgets.QColorDialog.getColor(
            self._workspace_color, self, "Цвет"
        )
        if color.isValid():
            self._workspace_color = color
            self._update_workspace_button()
            self._save_config()

    def _update_sidebar_button(self):
        self.btn_sidebar.setStyleSheet(
            f"background:{self._sidebar_color.name()}; border:1px solid #555;"
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
        self.btn_grad1.setStyleSheet(
            f"background:{self._grad_color1.name()}; border:1px solid #555;"
        )
        self.btn_grad2.setStyleSheet(
            f"background:{self._grad_color2.name()}; border:1px solid #555;"
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


class TopBar(QtWidgets.QWidget):
    prev_clicked = QtCore.Signal()
    next_clicked = QtCore.Signal()
    settings_clicked = QtCore.Signal()
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
        self.btn_prev.installEventFilter(NeonEventFilter(self.btn_prev))
        self.lbl_month = QtWidgets.QLabel("Месяц")
        self.lbl_month.setAlignment(QtCore.Qt.AlignCenter)
        self.lbl_month.setContentsMargins(8, 0, 8, 0)
        lay.addWidget(self.lbl_month)
        self.spin_year = QtWidgets.QSpinBox(self)
        self.spin_year.setRange(2000, 2100)
        self.spin_year.setValue(datetime.now().year)
        self.spin_year.setStyleSheet("padding:0 6px;")
        self.spin_year.valueChanged.connect(self.year_changed.emit)
        lay.addWidget(self.spin_year)
        self.spin_year.installEventFilter(NeonEventFilter(self.spin_year))
        lay.addStretch(1)
        self.btn_next = StyledToolButton(self)
        self.btn_next.setCursor(QtCore.Qt.PointingHandCursor)
        self.btn_next.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly)
        self.btn_next.clicked.connect(self.next_clicked)
        lay.addWidget(self.btn_next)
        self.btn_next.installEventFilter(NeonEventFilter(self.btn_next))
        self.btn_settings = StyledToolButton(self)
        self.btn_settings.setCursor(QtCore.Qt.PointingHandCursor)
        self.btn_settings.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly)
        self.btn_settings.clicked.connect(self.settings_clicked)
        lay.addWidget(self.btn_settings)
        self.btn_settings.installEventFilter(NeonEventFilter(self.btn_settings))
        self.update_icons()
        self.default_style = (
            "QLabel{color:#e5e5e5;} QToolButton{color:#e5e5e5; border:1px solid #555; border-radius:6px; padding:4px 10px;}"
        )
        self.setAutoFillBackground(True)
        color = QtGui.QColor(CONFIG.get("workspace_color", "#1e1e21"))
        if CONFIG.get("monochrome", False):
            color = theme_manager.apply_monochrome(color)
        self.apply_background(color)
        self.apply_style(CONFIG.get("neon", False))
        self.apply_fonts()

    def apply_fonts(self):
        font = QtGui.QFont(CONFIG.get("header_font", "Exo 2"))
        font.setBold(True)
        self.lbl_month.setFont(font)

    def update_icons(self) -> None:
        self.btn_prev.setIcon(icon("chevron-left"))
        self.btn_prev.setIconSize(QtCore.QSize(22, 22))
        self.btn_next.setIcon(icon("chevron-right"))
        self.btn_next.setIconSize(QtCore.QSize(22, 22))
        self.btn_settings.setIcon(icon("settings"))
        self.btn_settings.setIconSize(QtCore.QSize(22, 22))

    def apply_background(self, color: Union[str, QtGui.QColor]) -> None:
        qcolor = QtGui.QColor(color)
        pal = self.palette()
        pal.setColor(self.backgroundRole(), qcolor)
        self.setPalette(pal)

    def apply_style(self, neon):
        accent = QtGui.QColor(CONFIG.get("accent_color", "#39ff14"))
        if CONFIG.get("monochrome", False):
            h, s, v, _ = accent.getHsv()
            s = int(CONFIG.get("mono_saturation", 100))
            accent.setHsv(h, s, v)
        thickness = CONFIG.get("neon_thickness", 1)
        size = CONFIG.get("neon_size", 10)
        intensity = CONFIG.get("neon_intensity", 255)
        speed = CONFIG.get("animation_speed", 1.0)
        if neon:
            style = (
                f"QLabel{{color:{accent.name()};}} "
                f"QToolButton{{color:{accent.name()}; border:{thickness}px solid {accent.name()}; border-radius:6px; padding:4px 10px;}}"
            )
            self.setStyleSheet(style)
            for w in (self.lbl_month, self.btn_prev, self.btn_next, self.btn_settings):
                eff = QtWidgets.QGraphicsDropShadowEffect(self)
                eff.setOffset(0, 0)
                eff.setBlurRadius(size)
                c = QtGui.QColor(accent)
                c.setAlpha(intensity)
                eff.setColor(c)
                w.setGraphicsEffect(eff)
                if CONFIG.get("neon_motion", False):
                    anim = QtCore.QPropertyAnimation(eff, b"blurRadius", w)
                    anim.setStartValue(size)
                    anim.setEndValue(size * 2)
                    anim.setDuration(int(1000 / max(speed, 0.1)))
                    anim.setLoopCount(-1)
                    anim.setEasingCurve(QtCore.QEasingCurve.InOutQuad)
                    anim.start(QtCore.QAbstractAnimation.DeleteWhenStopped)
                    w._neon_anim = anim
                else:
                    w._neon_anim = None
        else:
            if CONFIG.get("theme", "dark") == "dark":
                style = self.default_style
            else:
                style = (
                    "QLabel{color:#000;} QToolButton{color:#000; border:1px solid #999; border-radius:6px; padding:4px 10px;}"
                )
            self.setStyleSheet(style)
            for w in (self.lbl_month, self.btn_prev, self.btn_next, self.btn_settings):
                if hasattr(w, "_neon_anim") and w._neon_anim:
                    w._neon_anim.stop()
                    w._neon_anim = None
                w.setGraphicsEffect(None)


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("План-график — Excel 1:1 + навигация по месяцам + левый бар")
        central = QtWidgets.QWidget(self)
        h = QtWidgets.QHBoxLayout(central); h.setContentsMargins(0,0,0,0); h.setSpacing(0)

        # left sidebar
        self.sidebar=CollapsibleSidebar(self); h.addWidget(self.sidebar)

        # right: vbox with topbar + table
        right = QtWidgets.QWidget(self); v = QtWidgets.QVBoxLayout(right); v.setContentsMargins(0,0,0,0); v.setSpacing(0)
        self.topbar = TopBar(self); v.addWidget(self.topbar)
        self.table = ExcelCalendarTable(self); v.addWidget(self.table, 1)
        h.addWidget(right, 1)

        self.setCentralWidget(central)
        self.showMaximized()

        # Connect topbar
        self.topbar.prev_clicked.connect(self.prev_month)
        self.topbar.next_clicked.connect(self.next_month)
        self.topbar.year_changed.connect(self.change_year)
        self.sidebar.btn_inputs.clicked.connect(self.open_input_dialog)
        self.sidebar.btn_release.clicked.connect(self.open_release_dialog)
        self.sidebar.btn_analytics.clicked.connect(self.open_analytics_dialog)
        self.sidebar.btn_tops.clicked.connect(self.open_top_dialog)
        self.topbar.settings_clicked.connect(self.open_settings_dialog)
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

    def open_analytics_dialog(self):
        dlg = AnalyticsDialog(self.table.year, self)
        dlg.exec()

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

    def open_top_dialog(self):
        dlg = TopDialog(self.table.year, self)
        dlg.exec()

    def open_settings_dialog(self):
        dlg = SettingsDialog(self)
        dlg.settings_changed.connect(self._on_settings_changed)
        dlg.exec()

    def _on_settings_changed(self):
        global CONFIG, BASE_SAVE_PATH
        CONFIG = load_config()
        BASE_SAVE_PATH = os.path.abspath(CONFIG.get("save_path", DATA_DIR))
        self.apply_settings()

    def apply_settings(self):
        theme_manager.set_text_font(CONFIG.get("text_font", "Inter"))
        theme_manager.set_header_font(CONFIG.get("header_font", "Exo 2"))
        load_icons(CONFIG.get("theme", "dark"))
        app = QtWidgets.QApplication.instance()
        self.topbar.update_icons()
        self.sidebar.update_icons()
        self.topbar.apply_fonts()
        header_font = QtGui.QFont(CONFIG.get("header_font"))
        self.table.setFont(app.font())
        self.table.horizontalHeader().setFont(header_font)
        for tbl in self.table.cell_tables.values():
            tbl.setFont(app.font())
            tbl.horizontalHeader().setFont(header_font)
        for lbl in self.table.day_labels.values():
            lbl.setFont(header_font)
        for w in [
            self.sidebar.btn_inputs,
            self.sidebar.btn_release,
            self.sidebar.btn_analytics,
            self.sidebar.btn_tops,
        ]:
            w.setFont(header_font)
        self.table.update_day_rows()
        accent = QtGui.QColor(CONFIG.get("accent_color", "#39ff14"))
        if CONFIG.get("monochrome", False):
            h, s, v, _ = accent.getHsv()
            s = int(CONFIG.get("mono_saturation", 100))
            accent.setHsv(h, s, v)
        self.topbar.apply_style(CONFIG.get("neon", False))
        self.sidebar.apply_style(CONFIG.get("neon", False), accent)
        speed = CONFIG.get("animation_speed", 1.0)
        self.sidebar.anim.setDuration(int(160 / max(speed, 0.1)))
        for dlg in app.topLevelWidgets():
            if isinstance(dlg, QtWidgets.QDialog):
                dlg.setFont(app.font())
                for tbl in dlg.findChildren(QtWidgets.QTableWidget):
                    tbl.setFont(app.font())
                    tbl.horizontalHeader().setFont(header_font)
        self.apply_theme()
        theme_manager.apply_glass_effect(self, CONFIG)





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
        self.topbar.apply_background(workspace)
        theme = CONFIG.get("theme", "dark")
        self.setStyleSheet(
            "QPushButton,"
            "QToolButton,QSpinBox,QDoubleSpinBox,QTimeEdit,"
            "QComboBox,QLineEdit{" + base + "}"
            f"""
            QSpinBox::up-button,QSpinBox::down-button{{
                border:1px solid transparent;
                border-radius:8px;
                width:16px;height:16px;
            }}
            QSpinBox::up-arrow{{ image:url(assets/icons/{theme}/chevron-up.svg); }}
            QSpinBox::down-arrow{{ image:url(assets/icons/{theme}/chevron-down.svg); }}
            """
        )
        self.table.setStyleSheet(f"QTableWidget {{ background-color: {workspace}; }}")
        for tbl in self.table.cell_tables.values():
            tbl.setStyleSheet(f"QTableWidget {{ background-color: {workspace}; }}")
        for container in self.table.cell_containers.values():
            container.setStyleSheet(f"background-color: {workspace};")
        sidebar = CONFIG.get("sidebar_color", "#1f1f23")
        if CONFIG.get("monochrome", False):
            sidebar = theme_manager.apply_monochrome(QtGui.QColor(sidebar)).name()
            accent = theme_manager.apply_monochrome(accent)
        self.sidebar.apply_style(CONFIG.get("neon", False), accent, sidebar)
    def closeEvent(self, event):
        self.table.save_current_month()
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(CONFIG, f, ensure_ascii=False, indent=2)
        super().closeEvent(event)


def main():
    QtCore.QLocale.setDefault(QtCore.QLocale("ru_RU"))
    app = QtWidgets.QApplication(sys.argv)
    register_fonts()
    load_icons(CONFIG.get("theme", "dark"))
    theme_manager.set_text_font(CONFIG.get("text_font", "Inter"))
    w = MainWindow()
    w.apply_settings()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
