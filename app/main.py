# -*- coding: utf-8 -*-
import sys, os, math, json, calendar
from datetime import datetime
from PySide6 import QtWidgets, QtGui, QtCore
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

ASSETS = os.path.join(os.path.dirname(__file__), "..", "assets")
DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")
EXCEL_PATH = os.path.join(ASSETS, "пример блоков.xlsx")
SHEET_NAME = "ЦЕНТРАЛЬНАЯ РАБОЧАЯ ОБЛАСТЬ"
ICON_TOGGLE = os.path.join(ASSETS, "gpt_icon.png")
ICON_VYKL = os.path.join(ASSETS, "ic_vykl.png")
ICON_TM   = os.path.join(ASSETS, "ic_tm.png")
ICON_TQ   = os.path.join(ASSETS, "ic_tq.png")
ICON_TP   = os.path.join(ASSETS, "ic_tp.png")
ICON_TG   = os.path.join(ASSETS, "ic_tg.png")

RU_MONTHS = ["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"]

def excel_col_width_to_pixels(width):
    if width is None: return 64
    return max(20, int((width * 7) + 5))

def excel_row_height_to_pixels(height_points):
    if height_points is None: return 20
    return max(16, int(round(height_points * (96.0/72.0))))

def qt_align(h, v):
    ah = {"left": QtCore.Qt.AlignLeft, "center": QtCore.Qt.AlignHCenter, "right": QtCore.Qt.AlignRight, "justify": QtCore.Qt.AlignJustify, None: QtCore.Qt.AlignLeft}.get(h, QtCore.Qt.AlignLeft)
    av = {"top": QtCore.Qt.AlignTop, "center": QtCore.Qt.AlignVCenter, "bottom": QtCore.Qt.AlignBottom, None: QtCore.Qt.AlignVCenter}.get(v, QtCore.Qt.AlignVCenter)
    return ah | av

class ExcelDelegate(QtWidgets.QStyledItemDelegate):
    def __init__(self, borders_map, rows, cols, parent=None):
        super().__init__(parent); self.borders_map=borders_map; self.rows=rows; self.cols=cols
    def paint(self, painter, option, index):
        super().paint(painter, option, index)
        r, c = index.row(), index.column(); sides = self.borders_map.get((r,c))
        if not sides: return
        pen = QtGui.QPen(QtGui.QColor(0,0,0)); pen.setWidth(1)
        painter.save(); painter.setPen(pen)
        rect = option.rect.adjusted(0,0,-1,-1)
        if sides.get("top"): painter.drawLine(rect.topLeft(), rect.topRight())
        if sides.get("left"): painter.drawLine(rect.topLeft(), rect.bottomLeft())
        if c==self.cols-1 and sides.get("right"): painter.drawLine(rect.topRight(), rect.bottomRight())
        if r==self.rows-1 and sides.get("bottom"): painter.drawLine(rect.bottomLeft(), rect.bottomRight())
        painter.restore()

class ExcelCalendarTable(QtWidgets.QTableWidget):
    """Рендер Excel + управление днями месяца и «хвостом» следующего/предыдущего месяца."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.setShowGrid(False); self.verticalHeader().setVisible(False); self.horizontalHeader().setVisible(False)
        self.setWordWrap(True); self.setEditTriggers(QtWidgets.QAbstractItemView.AllEditTriggers)

        now = datetime.now()
        self.year = now.year
        self.month = now.month

        self.base_col_widths=[]; self.base_row_heights=[]
        self.week_blocks=[]  # list of (numbers_row, weekdays_row, content_start_row) per неделю
        self.day_cols=(None,None)  # (start, end) индексы 7 колонок дней

        self.load_excel(EXCEL_PATH, SHEET_NAME)
        self.apply_month_numbers()

    # ---------- Persistence ----------
    def month_key(self, year=None, month=None):
        y = year or self.year; m = month or self.month
        return f"{y:04d}-{m:02d}"

    def save_current_month(self):
        path = os.path.join(DATA_DIR, self.month_key()+".json")
        data = {"rows": self.rowCount(), "cols": self.columnCount(), "grid": []}
        for r in range(self.rowCount()):
            row = []
            for c in range(self.columnCount()):
                it = self.item(r,c)
                row.append("" if it is None else it.text())
            data["grid"].append(row)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)

    def load_month_data(self, year, month):
        path = os.path.join(DATA_DIR, f"{year:04d}-{month:02d}.json")
        if not os.path.exists(path): return False
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        for r, row in enumerate(data.get("grid", [])):
            for c, txt in enumerate(row):
                it = self.item(r,c) or QtWidgets.QTableWidgetItem()
                it.setText(txt)
                self.setItem(r,c,it)
        return True

    # ---------- Excel load ----------
    def load_excel(self, path, sheet_name):
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter
        wb = load_workbook(path, data_only=True); ws = wb[sheet_name]
        min_row=min_col=max_row=max_col=None
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None or cell.coordinate in ws.merged_cells:
                    r,c=cell.row,cell.column
                    min_row = r if min_row is None else min(min_row,r)
                    min_col = c if min_col is None else min(min_col,c)
                    max_row = r if max_row is None else max(max_row,r)
                    max_col = c if max_col is None else max(max_col,c)
        if min_row is None: min_row=min_col=max_row=max_col=1
        rows=max_row-min_row+1; cols=max_col-min_col+1
        self.setRowCount(rows); self.setColumnCount(cols)

        borders={}
        for r in range(min_row, max_row+1):
            for c in range(min_col, max_col+1):
                b = ws.cell(r,c).border
                borders[(r-min_row,c-min_col)] = {"left":bool(b.left and b.left.style), "right":bool(b.right and b.right.style), "top":bool(b.top and b.top.style), "bottom":bool(b.bottom and b.bottom.style)}
        self.setItemDelegate(ExcelDelegate(borders, rows, cols, self))

        # sizes
        self.base_col_widths=[excel_col_width_to_pixels(ws.column_dimensions.get(get_column_letter(c)).width if ws.column_dimensions.get(get_column_letter(c)) else None) for c in range(min_col, max_col+1)]
        self.base_row_heights=[excel_row_height_to_pixels(ws.row_dimensions.get(r).height if ws.row_dimensions.get(r) else None) for r in range(min_row, max_row+1)]
        for i,w in enumerate(self.base_col_widths): self.setColumnWidth(i,w)
        for i,h in enumerate(self.base_row_heights): self.setRowHeight(i,h)

        # content + record weekday rows
        default_font_family="Calibri"; default_font_size=11
        wk_names = ["ПН","ВТ","СР","ЧТ","ПТ","СБ","ВС"]
        self.week_blocks.clear()
        self.day_cols=(None,None)

        def is_weekdays_row(texts):
            return texts == wk_names

        for r in range(min_row, max_row+1):
            # set items
            for c in range(min_col, max_col+1):
                cell=ws.cell(r,c)
                it = QtWidgets.QTableWidgetItem("" if cell.value is None else str(cell.value))
                f=cell.font; qf=QtGui.QFont(f.name or default_font_family, int(f.sz or default_font_size)); qf.setBold(bool(f.b)); qf.setItalic(bool(f.i)); it.setFont(qf)
                al=cell.alignment; it.setTextAlignment(qt_align(al.horizontal, al.vertical))
                it.setFlags(it.flags() | QtCore.Qt.ItemIsEditable)
                self.setItem(r-min_row, c-min_col, it)

            # detect weekdays row by scanning sequences of 7 cells
            row_texts = [self.item(r-min_row, c-min_col).text() for c in range(min_col, max_col+1)]
            for cstart in range(0, len(row_texts)-6):
                if row_texts[cstart:cstart+7] == wk_names:
                    numbers_row = (r-1)-min_row
                    weekdays_row = r-min_row
                    content_start_row = (r+1)-min_row
                    if self.day_cols==(None,None):
                        self.day_cols=(cstart, cstart+7)
                    self.week_blocks.append((numbers_row, weekdays_row, content_start_row))
                    break

        # merges
        for rng in ws.merged_cells.ranges:
            self.setSpan(rng.min_row-min_row, rng.min_col-min_col, rng.max_row-rng.min_row+1, rng.max_col-rng.min_col+1)

        self.update_layout()

    def update_layout(self):
        total_w=sum(self.base_col_widths) or 1; total_h=sum(self.base_row_heights) or 1
        vw=max(1, self.viewport().width()-2); vh=max(1, self.viewport().height()-2)
        sx=vw/total_w; sy=vh/total_h
        for i,w in enumerate(self.base_col_widths): self.setColumnWidth(i, max(20,int(round(w*sx))))
        for i,h in enumerate(self.base_row_heights): self.setRowHeight(i, max(16,int(round(h*sy))))
    def resizeEvent(self, e): super().resizeEvent(e); self.update_layout()

    # ---------- Month numbers logic ----------
    def apply_month_numbers(self):
        """Заполнить строки чисел дней для self.year/self.month. Хвост следующего/предыдущего месяца подсветить серым и сделать нередактируемым."""
        if not self.week_blocks or self.day_cols==(None,None): return
        y, m = self.year, self.month
        first_weekday, days_in_month = calendar.monthrange(y, m)  # Monday=0
        # Build 5 weeks x 7 days matrix of ints and a boolean is_current flag
        weeks = [[{"num":"", "active":False} for _ in range(7)] for _ in range(len(self.week_blocks))]

        # starting index in first week
        idx = first_weekday  # 0..6
        d = 1
        w = 0
        # fill current month
        while d <= days_in_month and w < len(weeks):
            while idx < 7 and d <= days_in_month:
                weeks[w][idx] = {"num":str(d), "active":True}
                d += 1; idx += 1
            w += 1; idx = 0

        # prev month tail for first week if needed
        if self.week_blocks:
            prev_y, prev_m = (y-1,12) if m==1 else (y,m-1)
            _, prev_days = calendar.monthrange(prev_y, prev_m)
            first_w = weeks[0]
            for i in range(7):
                if not first_w[i]["active"]:
                    # fill from right to left
                    count = sum(1 for j in range(7) if not first_w[j]["active"])
            # Fill leading blanks with tail of prev month
            lead = 0
            for i in range(7):
                if not first_w[i]["active"]:
                    lead += 1
            val = prev_days - lead + 1
            for i in range(lead):
                first_w[i] = {"num":str(val+i), "active":False}

        # next month tail in last week(s)
        next_y, next_m = (y+1,1) if m==12 else (y,m+1)
        nd = 1
        for wi in range(len(weeks)-1, -1, -1):
            row = weeks[wi]
            for i in range(7):
                if row[i]["num"]=="":
                    row[i] = {"num":str(nd), "active":False}; nd += 1

        # apply to table
        start_col, end_col = self.day_cols
        inactive_brush = QtGui.QBrush(QtGui.QColor(140,140,140))
        active_brush = QtGui.QBrush()
        for widx, (num_row, _, content_row) in enumerate(self.week_blocks):
            for i in range(7):
                col = start_col + i
                it = self.item(num_row, col) or QtWidgets.QTableWidgetItem()
                it.setText(weeks[widx][i]["num"])
                # grey inactive
                if weeks[widx][i]["active"]:
                    it.setForeground(active_brush); it.setFlags((it.flags() | QtCore.Qt.ItemIsEditable))
                else:
                    it.setForeground(inactive_brush); it.setFlags(it.flags() & ~QtCore.Qt.ItemIsEditable)
                self.setItem(num_row, col, it)

                # optionally grey out the entire content block below for inactive days
                # Здесь мы просто делаем первую строку контента нередактируемой/серой для неактивных дней
                content_it = self.item(content_row, col)
                if content_it:
                    if weeks[widx][i]["active"]:
                        content_it.setForeground(active_brush); content_it.setFlags(content_it.flags() | QtCore.Qt.ItemIsEditable)
                    else:
                        content_it.setForeground(inactive_brush); content_it.setFlags(content_it.flags() & ~QtCore.Qt.ItemIsEditable)

    # ---------- Navigation ----------
    def go_prev_month(self):
        self.save_current_month()
        if self.month==1: self.month=12; self.year-=1
        else: self.month-=1
        # load saved data if exists, else keep current content (Excel baseline)
        if not self.load_month_data(self.year, self.month):
            pass
        self.apply_month_numbers()

    def go_next_month(self):
        self.save_current_month()
        if self.month==12: self.month=1; self.year+=1
        else: self.month+=1
        if not self.load_month_data(self.year, self.month):
            pass
        self.apply_month_numbers()


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
            ("Выкладка", ICON_VYKL),
            ("Топ месяца", ICON_TM),
            ("Топ квартала", ICON_TQ),
            ("Топ полугода", ICON_TP),
            ("Топ года", ICON_TG),
        ]
        self.buttons=[]
        for label, icon in items:
            b=QtWidgets.QToolButton(self)
            b.setIcon(QtGui.QIcon(icon)); b.setIconSize(QtCore.QSize(22,22))
            b.setText(" "+label)
            b.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
            b.setCursor(QtCore.Qt.PointingHandCursor)
            b.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
            lay.addWidget(b); self.buttons.append(b)

        lay.addStretch(1)
        self._collapsed=False
        self.anim = QtCore.QPropertyAnimation(self, b"maximumWidth", self); self.anim.setDuration(160)
        self.setMaximumWidth(self.expanded_width)

    def set_collapsed(self, collapsed: bool):
        self._collapsed=collapsed
        self.anim.stop(); self.anim.setStartValue(self.maximumWidth()); self.anim.setEndValue(self.collapsed_width if collapsed else self.expanded_width); self.anim.start()
        for b in self.buttons:
            b.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly if collapsed else QtCore.Qt.ToolButtonTextBesideIcon)
        self.toggled.emit(not collapsed)

    def toggle(self): self.set_collapsed(not self._collapsed)


class TopBar(QtWidgets.QWidget):
    prev_clicked = QtCore.Signal()
    next_clicked = QtCore.Signal()
    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QtWidgets.QHBoxLayout(self); lay.setContentsMargins(8,8,8,8); lay.setSpacing(8)
        self.btn_prev = QtWidgets.QToolButton(self); self.btn_prev.setText(" < "); self.btn_prev.clicked.connect(self.prev_clicked); lay.addWidget(self.btn_prev)
        self.lbl_month = QtWidgets.QLabel("Месяц"); self.lbl_month.setAlignment(QtCore.Qt.AlignCenter); font=self.lbl_month.font(); font.setBold(True); self.lbl_month.setFont(font); lay.addWidget(self.lbl_month,1)
        self.btn_next = QtWidgets.QToolButton(self); self.btn_next.setText(" > "); self.btn_next.clicked.connect(self.next_clicked); lay.addWidget(self.btn_next)
        self.setStyleSheet("QLabel{color:#e5e5e5;} QToolButton{color:#e5e5e5; border:1px solid #555; border-radius:6px; padding:4px 10px;}")
        pal = self.palette(); pal.setColor(self.backgroundRole(), QtGui.QColor(30,30,33)); self.setAutoFillBackground(True); self.setPalette(pal)


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
        self._update_month_label()

    def _update_month_label(self):
        self.topbar.lbl_month.setText(f"{RU_MONTHS[self.table.month-1]} {self.table.year}")

    def prev_month(self):
        self.table.go_prev_month(); self._update_month_label()

    def next_month(self):
        self.table.go_next_month(); self._update_month_label()


def main():
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow(); w.show(); sys.exit(app.exec())

if __name__ == "__main__":
    main()
