# -*- coding: utf-8 -*-
import sys, os, math, json, calendar
from datetime import datetime, date
from typing import Dict, List

from PySide6 import QtWidgets, QtGui, QtCore
from dataclasses import dataclass, field, asdict

ASSETS = os.path.join(os.path.dirname(__file__), "..", "assets")
DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")
CONFIG_PATH = os.path.join(DATA_DIR, "config.json")


def load_config():
    default = {
        "neon": False,
        "header_font": "Arial",
        "text_font": "Arial",
        "save_path": DATA_DIR,
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

# роли для хранения структуры работ и отметки "18+"
WORK_ROLE = QtCore.Qt.UserRole
ADULT_ROLE = QtCore.Qt.UserRole + 1


@dataclass
class Work:
    plan: str
    done: bool = False
    is_adult: bool = False
    planned_chapters: int = 0
    done_chapters: int = 0


@dataclass
class MonthData:
    year: int
    month: int
    days: Dict[int, List[Work]] = field(default_factory=dict)

    @property
    def path(self) -> str:
        os.makedirs(DATA_DIR, exist_ok=True)
        return os.path.join(DATA_DIR, f"{self.year:04d}-{self.month:02d}.json")

    def save(self) -> None:
        data = {
            "year": self.year,
            "month": self.month,
            "days": {str(k): [asdict(w) for w in v] for k, v in self.days.items()},
        }
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls, year: int, month: int) -> "MonthData":
        path = os.path.join(DATA_DIR, f"{year:04d}-{month:02d}.json")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            days: Dict[int, List[Work]] = {}
            for k, v in data.get("days", {}).items():
                works_list = []
                for w in v:
                    works_list.append(
                        Work(
                            plan=w.get("plan", ""),
                            done=w.get("done", False),
                            is_adult=w.get("is_adult", False),
                            planned_chapters=w.get("planned_chapters", 0),
                            done_chapters=w.get("done_chapters", 0),
                        )
                    )
                days[int(k)] = works_list
            return cls(year=data.get("year", year), month=data.get("month", month), days=days)
        return cls(year=year, month=month)


class MonthDelegate(QtWidgets.QStyledItemDelegate):
    def paint(self, painter, option, index):
        super().paint(painter, option, index)

        if index.data(ADULT_ROLE):
            painter.save()
            painter.setPen(QtCore.Qt.NoPen)
            painter.setBrush(QtGui.QColor(200, 0, 0))
            rect = option.rect
            tri = QtGui.QPolygon(
                [rect.topRight(), rect.topRight() + QtCore.QPoint(-10, 0), rect.topRight() + QtCore.QPoint(0, 10)]
            )
            painter.drawPolygon(tri)
            painter.restore()

class WorkEditDialog(QtWidgets.QDialog):
    """Простое редактирование списка работ в ячейке."""

    def __init__(self, works, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Работы")
        self.resize(300, 200)
        self.works = works[:]

        self.lay = QtWidgets.QVBoxLayout(self)
        self.rows = []

        for w in self.works:
            self._add_row(w)

        btn_add = QtWidgets.QPushButton("Добавить", self)
        btn_add.clicked.connect(
            lambda: self._add_row(
                {
                    "plan": "",
                    "done": False,
                    "is_adult": False,
                    "planned_chapters": 0,
                    "done_chapters": 0,
                }
            )
        )
        self.lay.addWidget(btn_add)

        box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel, self)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        self.lay.addWidget(box)

    def _add_row(self, data):
        row_w = QtWidgets.QWidget(self)
        hl = QtWidgets.QHBoxLayout(row_w)
        hl.setContentsMargins(0, 0, 0, 0)
        cb_done = QtWidgets.QCheckBox("", row_w)
        cb_done.setChecked(data.get("done", False))
        hl.addWidget(cb_done)

        le = QtWidgets.QLineEdit(data.get("plan", ""), row_w)
        hl.addWidget(le, 1)

        sp_plan = QtWidgets.QSpinBox(row_w)
        sp_plan.setRange(0, 9999)
        sp_plan.setValue(data.get("planned_chapters", 0))
        sp_plan.setPrefix("П:")
        hl.addWidget(sp_plan)

        sp_done = QtWidgets.QSpinBox(row_w)
        sp_done.setRange(0, 9999)
        sp_done.setValue(data.get("done_chapters", 0))
        sp_done.setPrefix("Г:")
        hl.addWidget(sp_done)

        cb_adult = QtWidgets.QCheckBox("18+", row_w)
        cb_adult.setChecked(data.get("is_adult", False))
        hl.addWidget(cb_adult)

        self.lay.addWidget(row_w)
        self.rows.append((cb_done, le, sp_plan, sp_done, cb_adult))

    def get_works(self):
        works = []
        for cb_done, le, sp_plan, sp_done, cb_adult in self.rows:
            txt = le.text().strip()
            if not txt and sp_plan.value() == 0 and sp_done.value() == 0:
                continue
            works.append(
                {
                    "plan": txt,
                    "done": cb_done.isChecked(),
                    "is_adult": cb_adult.isChecked(),
                    "planned_chapters": sp_plan.value(),
                    "done_chapters": sp_done.value(),
                }
            )
        return works


class ReleaseDialog(QtWidgets.QDialog):
    """Диалог для управления выкладками в формате «День × Работы»."""

    def __init__(self, year, month, works, parent=None):
        super().__init__(parent)
        self.year = year
        self.month = month
        self.works = list(works)
        self.setWindowTitle("Выкладка")
        self.resize(600, 400)

        lay = QtWidgets.QVBoxLayout(self)
        self.days_in_month = calendar.monthrange(year, month)[1]
        self.table = QtWidgets.QTableWidget(self.days_in_month, len(self.works) + 1, self)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.AllEditTriggers)
        lay.addWidget(self.table)

        box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Close, self
        )
        box.accepted.connect(self.save)
        box.rejected.connect(self.reject)
        lay.addWidget(box)

        self._update_structure()
        self.load()

    def _update_structure(self):
        self.table.setColumnCount(len(self.works) + 1)
        self.table.setHorizontalHeaderLabels(["День"] + self.works)
        for d in range(1, self.days_in_month + 1):
            it = self.table.item(d - 1, 0)
            if not it:
                it = QtWidgets.QTableWidgetItem(str(d))
                it.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                self.table.setItem(d - 1, 0, it)
            else:
                it.setText(str(d))

    def file_path(self):
        return os.path.join(release_dir(self.year), f"{self.month:02d}.json")

    def load(self):
        path = self.file_path()
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            file_works = data.get("works", [])
            for w in file_works:
                if w not in self.works:
                    self.works.append(w)
            self._update_structure()
            for day_str, works_dict in data.get("days", {}).items():
                day = int(day_str)
                for work, text in works_dict.items():
                    if work in self.works:
                        col = self.works.index(work) + 1
                        self.table.setItem(day - 1, col, QtWidgets.QTableWidgetItem(str(text)))

    def save(self):
        data = {"works": self.works, "days": {}}
        for day in range(1, self.days_in_month + 1):
            row_data = {}
            for col, work in enumerate(self.works, 1):
                item = self.table.item(day - 1, col)
                if item:
                    txt = item.text().strip()
                    if txt:
                        row_data[work] = txt
            if row_data:
                data["days"][str(day)] = row_data
        os.makedirs(os.path.dirname(self.file_path()), exist_ok=True)
        with open(self.file_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.accept()


class InputDialog(QtWidgets.QDialog):
    """Форма ввода первоначальных данных за месяц."""

    def __init__(self, parent=None):
        super().__init__(parent)
        now = datetime.now()
        self.setWindowTitle("Вводные данные")
        self.resize(400, 500)

        form = QtWidgets.QFormLayout(self)

        self.spin_year = QtWidgets.QSpinBox(self)
        self.spin_year.setRange(2000, 2100)
        self.spin_year.setValue(now.year)
        form.addRow("Год", self.spin_year)

        self.combo_month = QtWidgets.QComboBox(self)
        for i, m in enumerate(RU_MONTHS, 1):
            self.combo_month.addItem(m, i)
        self.combo_month.setCurrentIndex(now.month - 1)
        form.addRow("Месяц", self.combo_month)

        self.edit_work = QtWidgets.QLineEdit(self)
        form.addRow("Работа", self.edit_work)

        self.edit_status = QtWidgets.QLineEdit(self)
        form.addRow("Статус", self.edit_status)

        self.chk_adult = QtWidgets.QCheckBox(self)
        form.addRow("18+", self.chk_adult)

        self.spin_total_chapters = QtWidgets.QSpinBox(self)
        self.spin_total_chapters.setRange(0, 100000)
        form.addRow("Всего глав", self.spin_total_chapters)

        self.spin_chars_per_chapter = QtWidgets.QSpinBox(self)
        self.spin_chars_per_chapter.setRange(0, 1000000)
        form.addRow("Знаков глава", self.spin_chars_per_chapter)

        self.spin_planned = QtWidgets.QSpinBox(self)
        self.spin_planned.setRange(0, 100000)
        form.addRow("Запланированно", self.spin_planned)

        self.spin_done = QtWidgets.QSpinBox(self)
        self.spin_done.setRange(0, 100000)
        form.addRow("Сделано глав", self.spin_done)

        self.spin_progress = QtWidgets.QDoubleSpinBox(self)
        self.spin_progress.setRange(0, 100)
        self.spin_progress.setSuffix("%")
        form.addRow("Прогресс перевода", self.spin_progress)

        self.edit_release = QtWidgets.QLineEdit(self)
        form.addRow("Выпуск", self.edit_release)

        self.dbl_profit = QtWidgets.QDoubleSpinBox(self)
        self.dbl_profit.setDecimals(2)
        self.dbl_profit.setRange(0, 1_000_000_000)
        form.addRow("Профит", self.dbl_profit)

        self.dbl_ads = QtWidgets.QDoubleSpinBox(self)
        self.dbl_ads.setDecimals(2)
        self.dbl_ads.setRange(0, 1_000_000_000)
        form.addRow("Затраты на рекламу", self.dbl_ads)

        self.spin_views = QtWidgets.QSpinBox(self)
        self.spin_views.setRange(0, 1_000_000_000)
        form.addRow("Просмотры", self.spin_views)

        self.spin_likes = QtWidgets.QSpinBox(self)
        self.spin_likes.setRange(0, 1_000_000_000)
        form.addRow("Лайки", self.spin_likes)

        self.spin_thanks = QtWidgets.QSpinBox(self)
        self.spin_thanks.setRange(0, 1_000_000_000)
        form.addRow("Спасибо", self.spin_thanks)

        box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel, self
        )
        box.accepted.connect(self.save)
        box.rejected.connect(self.reject)
        form.addRow(box)

    def save(self):
        year = self.spin_year.value()
        month = self.combo_month.currentData()
        record = {
            "work": self.edit_work.text().strip(),
            "status": self.edit_status.text().strip(),
            "adult": self.chk_adult.isChecked(),
            "total_chapters": self.spin_total_chapters.value(),
            "chars_per_chapter": self.spin_chars_per_chapter.value(),
            "planned": self.spin_planned.value(),
            "chapters": self.spin_done.value(),
            "progress": self.spin_progress.value(),
            "release": self.edit_release.text().strip(),
            "profit": self.dbl_profit.value(),
            "ads": self.dbl_ads.value(),
            "views": self.spin_views.value(),
            "likes": self.spin_likes.value(),
            "thanks": self.spin_thanks.value(),
        }
        record["chars"] = record["chapters"] * record["chars_per_chapter"]
        path = os.path.join(stats_dir(year), f"{year}.json")
        data = {}
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        month_list = data.setdefault(str(month), [])
        month_list.append(record)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.accept()


class AnalyticsDialog(QtWidgets.QDialog):
    """Годовая статистика: месяцы × показатели с колонкой "Итого за год"."""

    INDICATORS = [
        "Работ", "Завершенных", "Онгоингов", "Глав", "Знаков",
        "Просмотров", "Профит", "РК", "Лайков", "Спасибо",
        "Камса", "Потрачено на софт",
    ]

    def __init__(self, year, parent=None):
        super().__init__(parent)
        self.year = year
        self.setWindowTitle("Статистика")
        self.resize(900, 400)

        lay = QtWidgets.QVBoxLayout(self)
        top = QtWidgets.QHBoxLayout()
        top.addWidget(QtWidgets.QLabel("Год:"))
        self.spin_year = QtWidgets.QSpinBox(self)
        self.spin_year.setRange(2000, 2100)
        self.spin_year.setValue(year)
        self.spin_year.valueChanged.connect(self._year_changed)
        top.addWidget(self.spin_year)
        top.addStretch(1)
        self.btn_save = QtWidgets.QPushButton("Сохранить", self)
        self.btn_save.clicked.connect(self.save)
        top.addWidget(self.btn_save)
        lay.addLayout(top)

        cols = len(RU_MONTHS) + 1
        self.table = QtWidgets.QTableWidget(len(self.INDICATORS), cols, self)
        self.table.setHorizontalHeaderLabels(RU_MONTHS + ["Итого за год"])
        self.table.setVerticalHeaderLabels(self.INDICATORS)
        lay.addWidget(self.table, 1)

        # prepare items
        for r, name in enumerate(self.INDICATORS):
            for c in range(cols):
                it = QtWidgets.QTableWidgetItem("0")
                if name not in ("Камса", "Потрачено на софт") or c == cols - 1:
                    it.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                self.table.setItem(r, c, it)

        self.table.itemChanged.connect(self._item_changed)

        self._loading = False
        self._commissions = {str(m): 0.0 for m in range(1, 13)}
        self._software = {str(m): 0.0 for m in range(1, 13)}
        self.load(year)

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
        path = os.path.join(year_dir(year), f"{year}.json")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self._commissions.update({str(k): float(v) for k, v in data.get("commission", {}).items()})
            self._software.update({str(k): float(v) for k, v in data.get("software", {}).items()})

        # fill table with monthly values
        for m in range(1, 13):
            stats = self._calc_month_stats(year, m)
            for ind, val in stats.items():
                row = self.INDICATORS.index(ind)
                self.table.item(row, m - 1).setText(str(val))
            self.table.item(self.INDICATORS.index("Камса"), m - 1).setText(str(self._commissions[str(m)]))
            self.table.item(self.INDICATORS.index("Потрачено на софт"), m - 1).setText(str(self._software[str(m)]))

        self._recalculate()
        self._loading = False

    def save(self):
        path = os.path.join(year_dir(self.year), f"{self.year}.json")
        data = {
            "commission": self._commissions,
            "software": self._software,
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

        # totals
        for r in range(len(self.INDICATORS)):
            total = 0.0
            for c in range(cols):
                try:
                    total += float(self.table.item(r, c).text())
                except ValueError:
                    pass
            self.table.item(r, cols).setText(str(round(total, 2)))

class TopDialog(QtWidgets.QDialog):
    """Агрегирование и сохранение топов за период."""

    SORT_OPTIONS = [
        ("Профит", "profit"),
        ("Завершенность", "done"),
        ("Сделано глав", "chapters"),
        ("Знаков", "chars"),
        ("Просмотров", "views"),
        ("РК", "ads"),
        ("Лайков", "likes"),
        ("Спасибо", "thanks"),
    ]

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

        self.combo_mode = QtWidgets.QComboBox(self)
        self.combo_mode.addItem("Месяц", "month")
        self.combo_mode.addItem("Квартал", "quarter")
        self.combo_mode.addItem("Полугодие", "half")
        self.combo_mode.addItem("Год", "year")
        self.combo_mode.currentIndexChanged.connect(self._mode_changed)
        top.addWidget(self.combo_mode)

        self.combo_period = QtWidgets.QComboBox(self)
        top.addWidget(self.combo_period)
        self._mode_changed()  # fill periods

        top.addWidget(QtWidgets.QLabel("Сортировка:"))
        self.combo_sort = QtWidgets.QComboBox(self)
        for label, key in self.SORT_OPTIONS:
            self.combo_sort.addItem(label, key)
        top.addWidget(self.combo_sort)

        self.btn_calc = QtWidgets.QPushButton("Сформировать", self)
        self.btn_calc.clicked.connect(self.calculate)
        top.addWidget(self.btn_calc)
        lay.addLayout(top)

        headers = [
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
        self.table = QtWidgets.QTableWidget(0, len(headers), self)
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setStretchLastSection(True)
        lay.addWidget(self.table, 1)

        box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Close, self
        )
        box.accepted.connect(self.save)
        box.rejected.connect(self.reject)
        lay.addWidget(box)

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
        sort_key = self.combo_sort.currentData()
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

        results = sorted(totals.items(), key=lambda kv: kv[1].get(sort_key, 0), reverse=True)
        self.results = results
        self.table.setRowCount(0)
        for work, vals in results:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(work))
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(vals.get("status", "")))
            self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(vals["total_chapters"])))
            self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(str(vals["planned"])))
            self.table.setItem(row, 4, QtWidgets.QTableWidgetItem(str(vals["chapters"])))
            self.table.setItem(
                row,
                5,
                QtWidgets.QTableWidgetItem(f"{round(vals['progress'], 2)}%"),
            )
            self.table.setItem(row, 6, QtWidgets.QTableWidgetItem(vals.get("release", "")))
            self.table.setItem(row, 7, QtWidgets.QTableWidgetItem(str(vals["chars"])))
            self.table.setItem(row, 8, QtWidgets.QTableWidgetItem(str(vals["views"])))
            self.table.setItem(row, 9, QtWidgets.QTableWidgetItem(str(round(vals["profit"], 2))))
            self.table.setItem(row, 10, QtWidgets.QTableWidgetItem(str(round(vals["ads"], 2))))
            self.table.setItem(row, 11, QtWidgets.QTableWidgetItem(str(vals["likes"])))
            self.table.setItem(row, 12, QtWidgets.QTableWidgetItem(str(vals["thanks"])))

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
        data[key] = {
            "sort": self.combo_sort.currentData(),
            "results": results,
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.accept()

class ExcelCalendarTable(QtWidgets.QTableWidget):
    """Таблица календаря месяца, построенная на данных MonthData."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.verticalHeader().setVisible(False)
        self.setShowGrid(True)
        self.setWordWrap(True)
        self.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        wk_names = ["ПН", "ВТ", "СР", "ЧТ", "ПТ", "СБ", "ВС"]
        self.setColumnCount(7)
        self.setHorizontalHeaderLabels(wk_names)
        self.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.setItemDelegate(MonthDelegate(self))

        now = datetime.now()
        self.year = now.year
        self.month = now.month
        self.date_map: Dict[tuple[int, int], date] = {}

        self.load_month_data(self.year, self.month)

        self.itemDoubleClicked.connect(self.edit_cell)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_row_heights()

    def _update_row_heights(self):
        rows = self.rowCount()
        if rows:
            height = self.viewport().height() // rows
            for r in range(rows):
                self.setRowHeight(r, height)

    def _format_cell_text(self, day_num: int, works):
        lines = [str(day_num)]
        for w in works:
            prefix = "[x]" if w.get("done") else "[ ]"
            txt = f"{prefix} {w.get('plan', '')}"
            planned = w.get("planned_chapters", 0)
            done = w.get("done_chapters", 0)
            if planned or done:
                txt += f" ({done}/{planned})"
            if w.get("is_adult"):
                txt += " (18+)"
            lines.append(txt)
        return "\n".join(lines)

    def edit_cell(self, item):
        row, col = item.row(), item.column()
        day = self.date_map.get((row, col))
        if not day or day.month != self.month:
            return
        works = item.data(WORK_ROLE) or []
        dlg = WorkEditDialog(works, self)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            works = dlg.get_works()
            item.setData(WORK_ROLE, works)
            item.setData(ADULT_ROLE, any(w.get("is_adult") for w in works))
            item.setText(self._format_cell_text(day.day, works))

    # ---------- Persistence ----------
    def save_current_month(self):
        md = MonthData(year=self.year, month=self.month)
        for (r, c), day in self.date_map.items():
            if day.month != self.month:
                continue
            it = self.item(r, c)
            works = it.data(WORK_ROLE) or []
            if works:
                md.days[day.day] = [
                    Work(
                        plan=w.get("plan", ""),
                        done=w.get("done", False),
                        is_adult=w.get("is_adult", False),
                        planned_chapters=w.get("planned_chapters", 0),
                        done_chapters=w.get("done_chapters", 0),
                    )
                    for w in works
                ]
        md.save()

    def load_month_data(self, year: int, month: int):
        self.year = year
        self.month = month
        md = MonthData.load(year, month)
        cal = calendar.Calendar()
        weeks = cal.monthdatescalendar(year, month)
        self.date_map.clear()
        self.setRowCount(len(weeks))
        for r, week in enumerate(weeks):
            for c, day in enumerate(week):
                self.date_map[(r, c)] = day
                it = QtWidgets.QTableWidgetItem()
                if day.month != month:
                    it.setFlags(QtCore.Qt.NoItemFlags)
                    it.setText(str(day.day))
                    it.setForeground(QtGui.QBrush(QtGui.QColor(150, 150, 150)))
                else:
                    works = []
                    for w in md.days.get(day.day, []):
                        if isinstance(w, Work):
                            works.append(asdict(w))
                        else:
                            works.append(
                                {
                                    "plan": w.get("plan", ""),
                                    "done": w.get("done", False),
                                    "is_adult": w.get("is_adult", False),
                                    "planned_chapters": w.get("planned_chapters", 0),
                                    "done_chapters": w.get("done_chapters", 0),
                                }
                            )
                    it.setData(WORK_ROLE, works)
                    it.setData(ADULT_ROLE, any(w.get("is_adult") for w in works))
                    it.setText(self._format_cell_text(day.day, works))
                self.setItem(r, c, it)
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
        self.btn_toggle = QtWidgets.QToolButton(self)
        self.btn_toggle.setIcon(QtGui.QIcon(ICON_TOGGLE))
        self.btn_toggle.setIconSize(QtCore.QSize(28,28))
        self.btn_toggle.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly)
        self.btn_toggle.setCursor(QtCore.Qt.PointingHandCursor)
        self.btn_toggle.clicked.connect(self.toggle)
        lay.addWidget(self.btn_toggle)

        line = QtWidgets.QFrame(self); line.setFrameShape(QtWidgets.QFrame.HLine); line.setStyleSheet("color:#333;")
        lay.addWidget(line)

        items = [
            ("Вводные", ICON_TM),
            ("Выкладка", ICON_VYK),
            ("Аналитика", ICON_TG),
            ("Топы", ICON_TP),
        ]
        self.buttons=[]
        self.btn_inputs=None
        self.btn_release=None
        self.btn_analytics=None
        self.btn_tops=None
        for label, icon in items:
            b=QtWidgets.QToolButton(self)
            b.setIcon(QtGui.QIcon(icon)); b.setIconSize(QtCore.QSize(22,22))
            b.setText(" "+label)
            b.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
            b.setCursor(QtCore.Qt.PointingHandCursor)
            b.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
            lay.addWidget(b); self.buttons.append(b)
            if label == "Вводные":
                self.btn_inputs = b
            elif label == "Выкладка":
                self.btn_release = b
            elif label == "Аналитика":
                self.btn_analytics = b
            elif label == "Топы":
                self.btn_tops = b
        lay.addStretch(1)
        self._collapsed = False
        self.anim = QtCore.QPropertyAnimation(self, b"maximumWidth", self)
        self.anim.setDuration(160)
        self.setMinimumWidth(self.collapsed_width)
        self.setMaximumWidth(self.expanded_width)

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


class SettingsDialog(QtWidgets.QDialog):
    """Окно настроек приложения."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки")
        self.resize(400, 200)
        form = QtWidgets.QFormLayout(self)

        self.chk_neon = QtWidgets.QCheckBox(self)
        self.chk_neon.setChecked(CONFIG.get("neon", False))
        form.addRow("Неоновая подсветка", self.chk_neon)

        self.font_header = QtWidgets.QFontComboBox(self)
        self.font_header.setCurrentFont(QtGui.QFont(CONFIG.get("header_font", "Arial")))
        form.addRow("Шрифт заголовков", self.font_header)

        self.font_text = QtWidgets.QFontComboBox(self)
        self.font_text.setCurrentFont(QtGui.QFont(CONFIG.get("text_font", "Arial")))
        form.addRow("Шрифт текста", self.font_text)

        path_lay = QtWidgets.QHBoxLayout()
        self.edit_path = QtWidgets.QLineEdit(CONFIG.get("save_path", DATA_DIR), self)
        btn_browse = QtWidgets.QPushButton("...", self)
        btn_browse.clicked.connect(self.browse_path)
        path_lay.addWidget(self.edit_path, 1)
        path_lay.addWidget(btn_browse)
        form.addRow("Путь сохранения", path_lay)

        box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel, self
        )
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        form.addRow(box)

    def browse_path(self):
        path = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Выбрать папку", self.edit_path.text()
        )
        if path:
            self.edit_path.setText(path)

    def accept(self):
        config = {
            "neon": self.chk_neon.isChecked(),
            "header_font": self.font_header.currentFont().family(),
            "text_font": self.font_text.currentFont().family(),
            "save_path": self.edit_path.text().strip() or DATA_DIR,
        }
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        super().accept()


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
        self.btn_prev = QtWidgets.QToolButton(self)
        self.btn_prev.setText(" < ")
        self.btn_prev.clicked.connect(self.prev_clicked)
        lay.addWidget(self.btn_prev)
        self.lbl_month = QtWidgets.QLabel("Месяц")
        self.lbl_month.setAlignment(QtCore.Qt.AlignCenter)
        lay.addWidget(self.lbl_month)
        self.spin_year = QtWidgets.QSpinBox(self)
        self.spin_year.setRange(2000, 2100)
        self.spin_year.setValue(datetime.now().year)
        self.spin_year.valueChanged.connect(self.year_changed.emit)
        lay.addWidget(self.spin_year)
        lay.addStretch(1)
        self.btn_next = QtWidgets.QToolButton(self)
        self.btn_next.setText(" > ")
        self.btn_next.clicked.connect(self.next_clicked)
        lay.addWidget(self.btn_next)
        self.btn_settings = QtWidgets.QToolButton(self)
        self.btn_settings.setText("⚙")
        self.btn_settings.clicked.connect(self.settings_clicked)
        lay.addWidget(self.btn_settings)
        self.default_style = (
            "QLabel{color:#e5e5e5;} QToolButton{color:#e5e5e5; border:1px solid #555; border-radius:6px; padding:4px 10px;}"
        )
        self.neon_style = (
            "QLabel{color:#39ff14;} QToolButton{color:#39ff14; border:1px solid #39ff14; border-radius:6px; padding:4px 10px;}"
        )
        pal = self.palette()
        pal.setColor(self.backgroundRole(), QtGui.QColor(30, 30, 33))
        self.setAutoFillBackground(True)
        self.setPalette(pal)
        self.apply_style(CONFIG.get("neon", False))
        self.apply_fonts()

    def apply_fonts(self):
        font = QtGui.QFont(CONFIG.get("header_font", "Arial"))
        font.setBold(True)
        self.lbl_month.setFont(font)

    def apply_style(self, neon):
        self.setStyleSheet(self.neon_style if neon else self.default_style)


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
        dlg = InputDialog(self)
        dlg.spin_year.setValue(self.table.year)
        dlg.combo_month.setCurrentIndex(self.table.month - 1)
        dlg.exec()

    def open_analytics_dialog(self):
        dlg = AnalyticsDialog(self.table.year, self)
        dlg.exec()

    def _collect_work_names(self) -> List[str]:
        names = set()
        for (r, c), day in self.table.date_map.items():
            if day.month != self.table.month:
                continue
            it = self.table.item(r, c)
            if not it:
                continue
            for w in it.data(WORK_ROLE) or []:
                name = w.get("plan", "").strip()
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
        if dlg.exec():
            global CONFIG, BASE_SAVE_PATH
            CONFIG = load_config()
            BASE_SAVE_PATH = os.path.abspath(CONFIG.get("save_path", DATA_DIR))
            app = QtWidgets.QApplication.instance()
            app.setFont(QtGui.QFont(CONFIG.get("text_font", "Arial")))
            self.apply_settings()

    def apply_settings(self):
        self.topbar.apply_fonts()
        self.topbar.apply_style(CONFIG.get("neon", False))

    def closeEvent(self, event):
        self.table.save_current_month()
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(CONFIG, f, ensure_ascii=False, indent=2)
        super().closeEvent(event)


def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setFont(QtGui.QFont(CONFIG.get("text_font", "Arial")))
    w = MainWindow()
    w.apply_settings()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
