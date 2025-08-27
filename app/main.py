# -*- coding: utf-8 -*-
import sys, os, math, json, calendar
from datetime import datetime, date
from typing import Dict, List

from PySide6 import QtWidgets, QtGui, QtCore
from pydantic import BaseModel, Field

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
VERSION_FILE = os.path.join(os.path.dirname(__file__), "..", "VERSION")

RU_MONTHS = ["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"]

# роли для хранения структуры работ и отметки "18+"
WORK_ROLE  = QtCore.Qt.UserRole
ADULT_ROLE = QtCore.Qt.UserRole + 1


class Work(BaseModel):
    plan: str
    done: bool = False
    is_adult: bool = False


class MonthData(BaseModel):
    year: int
    month: int
    days: Dict[int, List[Work]] = Field(default_factory=dict)

    @property
    def path(self) -> str:
        os.makedirs(DATA_DIR, exist_ok=True)
        return os.path.join(DATA_DIR, f"{self.year:04d}-{self.month:02d}.json")

    def save(self) -> None:
        with open(self.path, "w", encoding="utf-8") as f:
            f.write(self.model_dump_json(indent=2, ensure_ascii=False))

    @classmethod
    def load(cls, year: int, month: int) -> "MonthData":
        path = os.path.join(DATA_DIR, f"{year:04d}-{month:02d}.json")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return cls.model_validate_json(f.read())
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
        btn_add.clicked.connect(lambda: self._add_row({"plan": "", "done": False, "is_adult": False}))
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
        cb_adult = QtWidgets.QCheckBox("18+", row_w)
        cb_adult.setChecked(data.get("is_adult", False))
        hl.addWidget(cb_adult)
        self.lay.addWidget(row_w)
        self.rows.append((cb_done, le, cb_adult))

    def get_works(self):
        works = []
        for cb_done, le, cb_adult in self.rows:
            txt = le.text().strip()
            if not txt:
                continue
            works.append({"plan": txt, "done": cb_done.isChecked(), "is_adult": cb_adult.isChecked()})
        return works


class ReleaseDialog(QtWidgets.QDialog):
    """Диалог для управления выкладками."""

    def __init__(self, year, month, parent=None):
        super().__init__(parent)
        self.year = year
        self.month = month
        self.setWindowTitle("Выкладка")
        self.resize(500, 300)

        lay = QtWidgets.QVBoxLayout(self)
        self.table = QtWidgets.QTableWidget(0, 4, self)
        self.table.setHorizontalHeaderLabels(["Дата", "Работа", "Глав", "Время"])
        self.table.horizontalHeader().setStretchLastSection(True)
        lay.addWidget(self.table)

        btn_row = QtWidgets.QHBoxLayout()
        btn_add = QtWidgets.QPushButton("Добавить", self)
        btn_add.clicked.connect(self.add_row)
        btn_del = QtWidgets.QPushButton("Удалить", self)
        btn_del.clicked.connect(self.remove_row)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_del)
        btn_row.addStretch(1)
        lay.addLayout(btn_row)

        box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Close, self)
        box.accepted.connect(self.save)
        box.rejected.connect(self.reject)
        lay.addWidget(box)

        self.load()

    def file_path(self):
        return os.path.join(release_dir(self.year), f"{self.month:02d}.json")

    def add_row(self, data=None):
        if data is None:
            data = {"date": "", "work": "", "chapters": "", "time": ""}
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(str(data.get("date", ""))))
        self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(data.get("work", "")))
        self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(data.get("chapters", ""))))
        self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(data.get("time", "")))

    def remove_row(self):
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

    def load(self):
        path = self.file_path()
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            for rec in data:
                self.add_row(rec)

    def save(self):
        data = []
        for r in range(self.table.rowCount()):
            date = self.table.item(r, 0)
            work = self.table.item(r, 1)
            chapters = self.table.item(r, 2)
            time = self.table.item(r, 3)
            if any(it and it.text().strip() for it in [date, work, chapters, time]):
                data.append({
                    "date": date.text().strip() if date else "",
                    "work": work.text().strip() if work else "",
                    "chapters": chapters.text().strip() if chapters else "",
                    "time": time.text().strip() if time else "",
                })
        with open(self.file_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.accept()

class YearStatsDialog(QtWidgets.QDialog):
    """Годовая статистика: месяцы × показатели с колонкой "Итого за год"."""

    INDICATORS = [
        "Работ", "Завершённых", "Онгоингов", "Глав", "Знаков",
        "Просмотров", "Профит", "Реклама (РК)", "Лайков", "Спасибо",
        "Комиссия", "Затраты на софт", "Чистыми",
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
                if name not in ("Комиссия", "Затраты на софт") or c == cols - 1:
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
            self.table.item(self.INDICATORS.index("Комиссия"), m - 1).setText(str(self._commissions[str(m)]))
            self.table.item(self.INDICATORS.index("Затраты на софт"), m - 1).setText(str(self._software[str(m)]))

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
        res = {k: 0 for k in self.INDICATORS if k not in ("Комиссия", "Затраты на софт", "Чистыми")}
        path = os.path.join(stats_dir(year), f"{year}.json")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            month_data = data.get(str(month), [])
            res["Работ"] = len(month_data)
            for rec in month_data:
                status = (rec.get("status", "") or "").lower()
                if "заверш" in status:
                    res["Завершённых"] += 1
                elif "онго" in status:
                    res["Онгоингов"] += 1
                res["Глав"] += int(rec.get("chapters", 0) or 0)
                res["Знаков"] += int(rec.get("chars", 0) or 0)
                res["Просмотров"] += int(rec.get("views", 0) or 0)
                res["Профит"] += float(rec.get("profit", 0) or 0)
                res["Реклама (РК)"] += float(rec.get("ads", 0) or 0)
                res["Лайков"] += int(rec.get("likes", 0) or 0)
                res["Спасибо"] += int(rec.get("thanks", 0) or 0)
        return res

    def _item_changed(self, item):
        if self._loading:
            return
        row = item.row()
        col = item.column()
        ind = self.INDICATORS[row]
        if ind == "Комиссия":
            try:
                self._commissions[str(col + 1)] = float(item.text())
            except ValueError:
                self._commissions[str(col + 1)] = 0.0
        elif ind == "Затраты на софт":
            try:
                self._software[str(col + 1)] = float(item.text())
            except ValueError:
                self._software[str(col + 1)] = 0.0
        self._recalculate()

    def _recalculate(self):
        row_profit = self.INDICATORS.index("Профит")
        row_comm = self.INDICATORS.index("Комиссия")
        row_soft = self.INDICATORS.index("Затраты на софт")
        row_net = self.INDICATORS.index("Чистыми")
        cols = len(RU_MONTHS)

        for c in range(cols):
            try:
                profit = float(self.table.item(row_profit, c).text())
            except ValueError:
                profit = 0.0
            comm = self._commissions[str(c + 1)]
            soft = self._software[str(c + 1)]
            net = profit - comm - soft
            it = self.table.item(row_net, c)
            it.setText(str(round(net, 2)))

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
        ("Завершённость", "done"),
        ("Глав", "chapters"),
        ("Знаков", "chars"),
        ("Просмотров", "views"),
        ("Реклама (РК)", "ads"),
        ("Лайков", "likes"),
        ("Спасибо", "thanks"),
    ]

    def __init__(self, year, mode="month", period=1, parent=None):
        super().__init__(parent)
        self.year = year
        self.mode = mode
        self.period = period
        self.results = []
        self.setWindowTitle("Топ")
        self.resize(700, 400)

        lay = QtWidgets.QVBoxLayout(self)
        top = QtWidgets.QHBoxLayout()
        top.addWidget(QtWidgets.QLabel("Год:"))
        self.spin_year = QtWidgets.QSpinBox(self)
        self.spin_year.setRange(2000, 2100)
        self.spin_year.setValue(year)
        top.addWidget(self.spin_year)

        self.combo_period = None
        if mode != "year":
            self.combo_period = QtWidgets.QComboBox(self)
            if mode == "month":
                top.addWidget(QtWidgets.QLabel("Месяц:"))
                for i, m in enumerate(RU_MONTHS, 1):
                    self.combo_period.addItem(m, i)
                self.combo_period.setCurrentIndex(period - 1)
            elif mode == "quarter":
                top.addWidget(QtWidgets.QLabel("Квартал:"))
                for i in range(1, 5):
                    self.combo_period.addItem(str(i), i)
                self.combo_period.setCurrentIndex(period - 1)
            elif mode == "half":
                top.addWidget(QtWidgets.QLabel("Полугодие:"))
                for i in range(1, 3):
                    self.combo_period.addItem(str(i), i)
                self.combo_period.setCurrentIndex(period - 1)
            top.addWidget(self.combo_period)

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
            "Глав",
            "Знаков",
            "Просмотров",
            "Профит",
            "Реклама",
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

    # --- helpers -------------------------------------------------------
    def _months_for_period(self):
        if self.mode == "month" and self.combo_period:
            return [self.combo_period.currentData()]
        if self.mode == "quarter" and self.combo_period:
            q = self.combo_period.currentData()
            start = (q - 1) * 3 + 1
            return list(range(start, start + 3))
        if self.mode == "half" and self.combo_period:
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
                            "chapters": 0,
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
                    t["chapters"] += int(rec.get("chapters", 0) or 0)
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
            self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(vals["chapters"])))
            self.table.setItem(row, 3, QtWidgets.QTableWidgetItem(str(vals["chars"])))
            self.table.setItem(row, 4, QtWidgets.QTableWidgetItem(str(vals["views"])))
            self.table.setItem(row, 5, QtWidgets.QTableWidgetItem(str(round(vals["profit"], 2))))
            self.table.setItem(row, 6, QtWidgets.QTableWidgetItem(str(round(vals["ads"], 2))))
            self.table.setItem(row, 7, QtWidgets.QTableWidgetItem(str(vals["likes"])))
            self.table.setItem(row, 8, QtWidgets.QTableWidgetItem(str(vals["thanks"])))

    def _period_key(self):
        if self.mode == "month" and self.combo_period:
            return f"M{self.combo_period.currentData():02d}"
        if self.mode == "quarter" and self.combo_period:
            return f"Q{self.combo_period.currentData()}"
        if self.mode == "half" and self.combo_period:
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
        data[key] = {
            "sort": self.combo_sort.currentData(),
            "results": [{"work": w, **vals} for w, vals in self.results],
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
        self.setItemDelegate(MonthDelegate(self))

        now = datetime.now()
        self.year = now.year
        self.month = now.month
        self.date_map: Dict[tuple[int, int], date] = {}

        self.load_month_data(self.year, self.month)

        self.itemDoubleClicked.connect(self.edit_cell)

    def _format_cell_text(self, day_num: int, works):
        lines = [str(day_num)]
        for w in works:
            prefix = "[x]" if w.get("done") else "[ ]"
            txt = f"{prefix} {w.get('plan', '')}"
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
                md.days[day.day] = [Work(**w) for w in works]
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
                    works = [w.model_dump() if isinstance(w, Work) else w for w in md.days.get(day.day, [])]
                    it.setData(WORK_ROLE, works)
                    it.setData(ADULT_ROLE, any(w.get("is_adult") for w in works))
                    it.setText(self._format_cell_text(day.day, works))
                self.setItem(r, c, it)
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
            ("Статистика", ICON_TG),
            ("Топ месяца", ICON_TM),
            ("Топ квартала", ICON_TQ),
            ("Топ полугода", ICON_TP),
            ("Топ года", ICON_TG),
        ]
        self.buttons=[]
        self.btn_stats=None
        self.btn_top_month=None
        self.btn_top_quarter=None
        self.btn_top_half=None
        self.btn_top_year=None
        for label, icon in items:
            b=QtWidgets.QToolButton(self)
            b.setIcon(QtGui.QIcon(icon)); b.setIconSize(QtCore.QSize(22,22))
            b.setText(" "+label)
            b.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
            b.setCursor(QtCore.Qt.PointingHandCursor)
            b.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
            lay.addWidget(b); self.buttons.append(b)
            if label == "Статистика":
                self.btn_stats = b
            elif label == "Топ месяца":
                self.btn_top_month = b
            elif label == "Топ квартала":
                self.btn_top_quarter = b
            elif label == "Топ полугода":
                self.btn_top_half = b
            elif label == "Топ года":
                self.btn_top_year = b

        # --- stats table ---
        self.current_year = datetime.now().year
        self.current_month = datetime.now().month
        self.sorted_col = 0
        self.sorted_order = QtCore.Qt.AscendingOrder

        self.stats_table = QtWidgets.QTableWidget(0, 4, self)
        self.stats_table.setHorizontalHeaderLabels(["Год", "Работа", "Статус", "18+"])
        self.stats_table.horizontalHeader().setStretchLastSection(True)
        self.stats_table.setSortingEnabled(True)
        self.stats_table.horizontalHeader().sortIndicatorChanged.connect(self._update_sort)
        lay.addWidget(self.stats_table)

        # input fields below table
        form = QtWidgets.QWidget(self)
        form_lay = QtWidgets.QFormLayout(form)
        self.edit_work = QtWidgets.QLineEdit(form)
        self.edit_status = QtWidgets.QLineEdit(form)
        self.edit_adult = QtWidgets.QCheckBox("", form)
        form_lay.addRow("Работа", self.edit_work)
        form_lay.addRow("Статус", self.edit_status)
        form_lay.addRow("18+", self.edit_adult)
        self.btn_save_stats = QtWidgets.QPushButton("Сохранить", form)
        self.btn_save_stats.clicked.connect(self.save_stats)
        form_lay.addRow(self.btn_save_stats)
        lay.addWidget(form)
        self.form_widget = form

        lay.addStretch(1)
        self._collapsed=False
        self.anim = QtCore.QPropertyAnimation(self, b"maximumWidth", self); self.anim.setDuration(160)
        self.setMaximumWidth(self.expanded_width)

    def _update_sort(self, col, order):
        self.sorted_col = col
        self.sorted_order = order

    def _add_row(self, work, status, adult):
        row = self.stats_table.rowCount()
        self.stats_table.insertRow(row)
        self.stats_table.setItem(row, 0, QtWidgets.QTableWidgetItem(str(self.current_year)))
        self.stats_table.setItem(row, 1, QtWidgets.QTableWidgetItem(work))
        self.stats_table.setItem(row, 2, QtWidgets.QTableWidgetItem(status))
        adult_item = QtWidgets.QTableWidgetItem("Да" if adult else "Нет")
        adult_item.setData(QtCore.Qt.UserRole, adult)
        self.stats_table.setItem(row, 3, adult_item)

    def _load_stats(self):
        path = os.path.join(stats_dir(self.current_year), f"{self.current_year}.json")
        self.stats_table.setRowCount(0)
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            month_data = data.get(str(self.current_month), [])
            for rec in month_data:
                self._add_row(rec.get("work", ""), rec.get("status", ""), rec.get("adult", False))
        if self.stats_table.rowCount() and self.sorted_col is not None:
            self.stats_table.sortItems(self.sorted_col, self.sorted_order)

    def _write_stats_file(self):
        path = os.path.join(stats_dir(self.current_year), f"{self.current_year}.json")
        data = {}
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        month_key = str(self.current_month)
        data[month_key] = []
        for row in range(self.stats_table.rowCount()):
            data[month_key].append({
                "work": self.stats_table.item(row,1).text(),
                "status": self.stats_table.item(row,2).text(),
                "adult": bool(self.stats_table.item(row,3).data(QtCore.Qt.UserRole))
            })
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def save_stats(self):
        work = self.edit_work.text().strip()
        status = self.edit_status.text().strip()
        adult = self.edit_adult.isChecked()
        if not work:
            return
        self._add_row(work, status, adult)
        self.stats_table.sortItems(self.sorted_col, self.sorted_order)
        self.edit_work.clear(); self.edit_status.clear(); self.edit_adult.setChecked(False)
        self._write_stats_file()

    def save_all(self):
        self._write_stats_file()

    def set_period(self, year, month):
        self.current_year = year
        self.current_month = month
        self._load_stats()

    def set_collapsed(self, collapsed: bool):
        self._collapsed=collapsed
        self.anim.stop(); self.anim.setStartValue(self.maximumWidth()); self.anim.setEndValue(self.collapsed_width if collapsed else self.expanded_width); self.anim.start()
        for b in self.buttons:
            b.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly if collapsed else QtCore.Qt.ToolButtonTextBesideIcon)
        self.stats_table.setVisible(not collapsed)
        self.form_widget.setVisible(not collapsed)
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
        self.sidebar.btn_stats.clicked.connect(self.open_year_stats_dialog)
        self.sidebar.btn_top_month.clicked.connect(lambda: self.open_top_dialog("month"))
        self.sidebar.btn_top_quarter.clicked.connect(lambda: self.open_top_dialog("quarter"))
        self.sidebar.btn_top_half.clicked.connect(lambda: self.open_top_dialog("half"))
        self.sidebar.btn_top_year.clicked.connect(lambda: self.open_top_dialog("year"))
        self.topbar.settings_clicked.connect(self.open_settings_dialog)
        self._update_month_label()
        self.sidebar.set_period(self.table.year, self.table.month)

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
        self.sidebar.set_period(self.table.year, self.table.month)

    def next_month(self):
        self.table.go_next_month(); self._update_month_label()
        self.sidebar.set_period(self.table.year, self.table.month)

    def change_year(self, year):
        self.table.save_current_month()
        self.table.year = year
        self.table.load_month_data(year, self.table.month)
        self._update_month_label()
        self.sidebar.set_period(self.table.year, self.table.month)

    def open_year_stats_dialog(self):
        dlg = YearStatsDialog(self.table.year, self)
        dlg.exec()

    def open_top_dialog(self, mode):
        if mode == "month":
            period = self.table.month
        elif mode == "quarter":
            period = (self.table.month - 1) // 3 + 1
        elif mode == "half":
            period = (self.table.month - 1) // 6 + 1
        else:
            period = 1
        dlg = TopDialog(self.table.year, mode, period, self)
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
        self.sidebar.save_all()
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
