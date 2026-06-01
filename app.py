#!/usr/bin/env python3
"""
Tutor Diff — Desktop GUI
PyQt6 wrapper around build_excel.py.
"""

import sys
import os
import tempfile
import subprocess
from datetime import datetime
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QMessageBox, QFrame,
    QSizePolicy, QProgressBar, QGraphicsDropShadowEffect,
    QAbstractButton, QCalendarWidget,
)
from PyQt6.QtCore import (
    Qt, QDate, QThread, pyqtSignal, QSettings, QSize,
    QVariantAnimation, QEasingCurve, QLocale, QTimer, QPoint,
    QEvent, QRect,
)
from PyQt6.QtGui import QFont, QDragEnterEvent, QDropEvent, QColor, QPainter

from openpyxl import load_workbook
from build_excel import run_comparison, build_report, fix_xlsx

# ── Palette ───────────────────────────────────────────────────────────────────
BG        = '#eef2f7'
CARD      = '#ffffff'
PRIMARY   = '#1e3a5f'
PRIMARY_H = '#2d5a8e'
ACCENT    = '#3b82f6'
SUCCESS   = '#059669'
SUCCESS_H = '#047857'
MUTED     = '#64748b'
BORDER    = '#dde3ea'
TEXT      = '#1e293b'


# ── Styled calendar popup ─────────────────────────────────────────────────────

class StyledCalendar(QCalendarWidget):
    """Beautiful calendar widget — used via QDateEdit.setCalendarWidget()."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setGridVisible(False)
        self.setVerticalHeaderFormat(QCalendarWidget.VerticalHeaderFormat.NoVerticalHeader)
        self.setHorizontalHeaderFormat(QCalendarWidget.HorizontalHeaderFormat.SingleLetterDayNames)
        self.setFirstDayOfWeek(Qt.DayOfWeek.Sunday)
        self.setLocale(QLocale(QLocale.Language.Hebrew, QLocale.Country.Israel))
        self.setMinimumWidth(290)
        self.setMinimumHeight(270)
        self._hovered:    QDate | None = None
        self._cell_rects: dict        = {}
        self._apply_style()
        QTimer.singleShot(0, self._style_nav_buttons)
        QTimer.singleShot(0, self._setup_hover)

    def _setup_hover(self):
        from PyQt6.QtWidgets import QAbstractItemView
        view = self.findChild(QAbstractItemView)
        if view:
            view.setMouseTracking(True)
            view.installEventFilter(self)

    def eventFilter(self, obj, event):
        from PyQt6.QtWidgets import QAbstractItemView
        if isinstance(obj, QAbstractItemView):
            t = event.type()
            if t == QEvent.Type.MouseMove:
                new_h = self._date_at(event.pos())
                if new_h != self._hovered:
                    self._hovered = new_h
                    self.updateCells()
            elif t == QEvent.Type.Leave:
                if self._hovered is not None:
                    self._hovered = None
                    self.updateCells()
        return super().eventFilter(obj, event)

    def _date_at(self, pos) -> 'QDate | None':
        for date, rect in self._cell_rects.items():
            if rect.contains(pos):
                return date
        return None

    def _apply_style(self):
        self.setStyleSheet(f"""
            QCalendarWidget QWidget#qt_calendar_navigationbar {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {PRIMARY}, stop:1 {PRIMARY_H});
                min-height: 50px;
                padding: 0 8px;
            }}
            QCalendarWidget QToolButton {{
                color: white; background: transparent;
                border: none; font-size: 14px; font-weight: bold;
                padding: 4px 10px; border-radius: 7px;
            }}
            QCalendarWidget QToolButton:hover {{
                background: rgba(255,255,255,0.18);
            }}
            QCalendarWidget QToolButton::menu-indicator {{
                image: none; width: 0;
            }}
            QCalendarWidget QSpinBox {{
                color: white;
                background: rgba(255,255,255,0.15);
                border: 1px solid rgba(255,255,255,0.25);
                border-radius: 5px; padding: 2px 6px;
                font-size: 13px; font-weight: bold;
            }}
            QCalendarWidget QSpinBox::up-button,
            QCalendarWidget QSpinBox::down-button {{ width: 0; height: 0; }}
            QCalendarWidget QAbstractItemView:enabled {{
                color: {TEXT}; background: white;
                selection-background-color: transparent;
                selection-color: white;
                outline: none; font-size: 13px;
            }}
            QCalendarWidget QAbstractItemView:disabled {{ color: #cbd5e1; }}
            QCalendarWidget QWidget {{ alternate-background-color: white; }}
        """)

    def _style_nav_buttons(self):
        for name, arrow in (('qt_calendar_prevmonth', '‹'),
                             ('qt_calendar_nextmonth', '›')):
            btn = self.findChild(QPushButton, name)
            if btn:
                btn.setText(arrow)
                btn.setStyleSheet("""
                    QPushButton {
                        color: white;
                        background: rgba(255,255,255,0.15);
                        border: none; border-radius: 7px;
                        font-size: 20px; font-weight: 300;
                        min-width: 32px; max-width: 32px;
                        min-height: 32px; max-height: 32px;
                    }
                    QPushButton:hover  { background: rgba(255,255,255,0.28); }
                    QPushButton:pressed { background: rgba(255,255,255,0.08); }
                """)

    def paintCell(self, painter, rect, date):
        if not date.isValid():
            return

        # Store rect for hover hit-testing
        self._cell_rects[date] = QRect(
            int(rect.x()), int(rect.y()),
            int(rect.width()), int(rect.height())
        )

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)

        today    = QDate.currentDate()
        selected = self.selectedDate()
        other_mo = date.month() != self.monthShown()
        hovered  = date == self._hovered and not other_mo

        d = min(rect.width(), rect.height()) - 8
        x = rect.center().x() - d / 2
        y = rect.center().y() - d / 2

        if date == selected:
            painter.setBrush(QColor(ACCENT))
            painter.drawEllipse(x, y, d, d)
            painter.setPen(QColor('#ffffff'))
            f = QFont(); f.setPointSize(12); f.setBold(True)
        elif date == today:
            painter.setBrush(QColor('#dbeafe'))
            painter.drawEllipse(x, y, d, d)
            painter.setPen(QColor(ACCENT))
            f = QFont(); f.setPointSize(12); f.setBold(True)
        elif hovered:
            painter.setBrush(QColor('#e8edf2'))
            painter.drawEllipse(x, y, d, d)
            painter.setPen(QColor(TEXT))
            f = QFont(); f.setPointSize(12)
        else:
            color = '#b8c8d8' if other_mo else ('#64748b' if date.dayOfWeek() == 7 else TEXT)
            painter.setPen(QColor(color))
            f = QFont(); f.setPointSize(12)

        painter.setFont(f)
        painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, str(date.day()))
        painter.restore()


# ── Date picker button ────────────────────────────────────────────────────────

class DatePickerButton(QPushButton):
    """Button that shows the selected date and opens StyledCalendar on click."""
    dateChanged = pyqtSignal(QDate)

    def __init__(self, initial: QDate, parent=None):
        super().__init__(parent)
        self._date = initial
        self._cal  = StyledCalendar()
        self._cal.setWindowFlags(
            Qt.WindowType.Popup | Qt.WindowType.FramelessWindowHint
        )
        self._cal.setSelectedDate(initial)
        self._cal.clicked.connect(self._pick)
        self.clicked.connect(self._open)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedHeight(36)
        self.setMinimumWidth(140)
        self.setStyleSheet(f"""
            QPushButton {{
                border: 1.5px solid {BORDER}; border-radius: 8px;
                padding: 5px 14px; background: #f8fafc;
                color: {TEXT}; font-size: 13px; text-align: center;
            }}
            QPushButton:hover {{
                background: white; border-color: {ACCENT};
            }}
        """)
        self._refresh()

    def _refresh(self):
        self.setText('📅  ' + self._date.toString('dd.MM.yyyy'))

    def _open(self):
        self._cal.setSelectedDate(self._date)
        # Position below the button
        pos = self.mapToGlobal(QPoint(0, self.height() + 4))
        self._cal.move(pos)
        self._cal.show()
        self._cal.raise_()

    def _pick(self, date: QDate):
        self._date = date
        self._refresh()
        self._cal.hide()
        self.dateChanged.emit(date)

    def date(self) -> QDate:
        return self._date


def _shadow(blur=20, y=5, alpha=30):
    e = QGraphicsDropShadowEffect()
    e.setBlurRadius(blur)
    e.setXOffset(0)
    e.setYOffset(y)
    e.setColor(QColor(0, 0, 0, alpha))
    return e


# ── Animated iOS-style toggle ─────────────────────────────────────────────────

class ToggleSwitch(QAbstractButton):
    """Smooth animated toggle — uses QAbstractButton for reliable click handling."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setFixedSize(46, 26)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self._pos = 3.0
        self._anim = QVariantAnimation(self)
        self._anim.setDuration(150)
        self._anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._anim.valueChanged.connect(self._update)
        self.toggled.connect(self._kick)

    def _update(self, v):
        self._pos = float(v)
        self.update()

    def _kick(self, checked: bool):
        self._anim.setStartValue(float(self._pos))
        self._anim.setEndValue(22.0 if checked else 3.0)
        self._anim.start()

    def paintEvent(self, _event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(ACCENT if self.isChecked() else '#c8d5e0'))
        p.drawRoundedRect(0, 0, 46, 26, 13, 13)
        p.setBrush(QColor('#ffffff'))
        p.drawEllipse(int(self._pos), 3, 20, 20)
        p.end()

    def sizeHint(self):
        return QSize(46, 26)


# ── Base card ─────────────────────────────────────────────────────────────────

class Card(QFrame):
    def __init__(self, parent=None, radius=14):
        super().__init__(parent)
        self._radius = radius
        self.setStyleSheet(f"""
            Card {{
                background: {CARD};
                border-radius: {radius}px;
                border: 1px solid {BORDER};
            }}
        """)
        self.setGraphicsEffect(_shadow())


# ── Drop zone ─────────────────────────────────────────────────────────────────

class DropZone(QFrame):
    file_changed = pyqtSignal(str)
    cleared      = pyqtSignal()

    def __init__(self, title: str, subtitle: str, parent=None):
        super().__init__(parent)
        self.file_path: str | None = None
        self.setAcceptDrops(True)
        self.setMinimumHeight(172)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.setGraphicsEffect(_shadow())
        self._build(title, subtitle)
        self._apply_style('empty')

    def _build(self, title, subtitle):
        lay = QVBoxLayout(self)
        lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.setSpacing(5)
        lay.setContentsMargins(16, 16, 16, 14)

        self._icon = QLabel('📂')
        self._icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._icon.setFont(QFont('', 32))

        self._title = QLabel(title)
        self._title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        f = QFont(); f.setBold(True); f.setPointSize(14)
        self._title.setFont(f)
        self._title.setStyleSheet(f'color: {TEXT}; background: transparent;')

        self._sub = QLabel(subtitle)
        self._sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._sub.setStyleSheet(f'color: {MUTED}; font-size: 11px; background: transparent;')

        self._file = QLabel('')
        self._file.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._file.setStyleSheet(f'color: {ACCENT}; font-size: 11px; font-weight: bold; background: transparent;')
        self._file.setWordWrap(True)

        self._meta = QLabel('')
        self._meta.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._meta.setStyleSheet('color: #9aa5b1; font-size: 10px; background: transparent;')

        btn_row = QHBoxLayout()
        btn_row.setSpacing(6)
        btn_row.addStretch()

        self._choose_btn = QPushButton('בחר קובץ')
        self._choose_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._choose_btn.setStyleSheet(f"""
            QPushButton {{
                background: #f1f5f9; border: 1px solid {BORDER};
                border-radius: 6px; padding: 5px 16px; font-size: 12px; color: {TEXT};
            }}
            QPushButton:hover {{ background: #e2e8f0; }}
        """)
        self._choose_btn.clicked.connect(self._choose)

        self._clear_btn = QPushButton('✕')
        self._clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._clear_btn.setFixedSize(28, 28)
        self._clear_btn.setToolTip('נקה קובץ')
        self._clear_btn.setStyleSheet("""
            QPushButton {
                background: #fee2e2; border: 1px solid #fca5a5;
                border-radius: 6px; font-size: 11px; color: #dc2626;
            }
            QPushButton:hover { background: #fca5a5; }
        """)
        self._clear_btn.clicked.connect(self._clear)
        self._clear_btn.setVisible(False)

        btn_row.addWidget(self._choose_btn)
        btn_row.addWidget(self._clear_btn)
        btn_row.addStretch()

        for w in (self._icon, self._title, self._sub, self._file, self._meta):
            lay.addWidget(w)
        lay.addLayout(btn_row)

    def _apply_style(self, state: str):
        styles = {
            'empty':  f'background:{CARD}; border:2px dashed #c0cdd9; border-radius:14px;',
            'filled': f'background:#f0f7ff; border:2px solid {ACCENT}; border-radius:14px;',
            'hover':  f'background:#dbeafe; border:2px dashed {ACCENT}; border-radius:14px;',
        }
        self.setStyleSheet(f'DropZone {{ {styles[state]} }}')
        self._icon.setText('✅' if state == 'filled' else '📂')

    def set_file(self, path: str):
        self.file_path = path
        name = Path(path).name
        self._file.setText(name if len(name) <= 36 else name[:33] + '…')
        try:
            kb   = os.path.getsize(path) // 1024
            date = datetime.fromtimestamp(os.path.getmtime(path)).strftime('%d.%m.%Y')
            self._meta.setText(f'{kb:,} KB  ·  {date}')
        except Exception:
            self._meta.setText('')
        self._clear_btn.setVisible(True)
        self._apply_style('filled')
        self.file_changed.emit(path)

    def _clear(self):
        self.file_path = None
        self._file.setText('')
        self._meta.setText('')
        self._clear_btn.setVisible(False)
        self._apply_style('empty')
        self.cleared.emit()

    def _choose(self):
        from PyQt6.QtWidgets import QFileDialog
        start = str(Path(self.file_path).parent) if self.file_path else os.path.expanduser('~')
        path, _ = QFileDialog.getOpenFileName(self, 'בחר קובץ Excel', start, 'Excel (*.xlsx *.xls)')
        if path:
            self.set_file(path)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0].toLocalFile().lower()
            if url.endswith(('.xlsx', '.xls')):
                event.acceptProposedAction()
                self._apply_style('hover')

    def dragLeaveEvent(self, event):
        self._apply_style('filled' if self.file_path else 'empty')

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            self.set_file(urls[0].toLocalFile())


# ── Date range card ───────────────────────────────────────────────────────────

class DateRangeCard(QFrame):
    """Clean card with toggle + inline date pickers. No QGroupBox."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet(f"""
            DateRangeCard {{
                background: {CARD};
                border: 1px solid {BORDER};
                border-radius: 14px;
            }}
        """)
        self.setGraphicsEffect(_shadow(14, 3, 20))
        self._build()

    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 14, 18, 14)
        root.setSpacing(12)

        # ── Header row ────────────────────────────────────────────────
        header = QHBoxLayout()
        header.setSpacing(10)

        icon_lbl = QLabel('📅')
        icon_lbl.setStyleSheet('font-size: 16px; background: transparent;')

        text_lbl = QLabel('סינון לפי טווח תאריכים')
        text_lbl.setStyleSheet(f'color: {TEXT}; font-size: 13px; font-weight: 600; background: transparent;')

        hint_lbl = QLabel('(אופציונלי)')
        hint_lbl.setStyleSheet(f'color: {MUTED}; font-size: 11px; background: transparent;')

        self.toggle = ToggleSwitch()
        self.toggle.toggled.connect(self._on_toggle)

        header.addWidget(icon_lbl)
        header.addWidget(text_lbl)
        header.addWidget(hint_lbl)
        header.addStretch()
        header.addWidget(self.toggle)
        root.addLayout(header)

        # ── Date pickers row (hidden by default) ───────────────────────
        self._pickers = QWidget()
        self._pickers.setStyleSheet('background: transparent;')
        prow = QHBoxLayout(self._pickers)
        prow.setContentsMargins(0, 0, 0, 0)
        prow.setSpacing(10)

        from_lbl = QLabel('מ:')
        from_lbl.setStyleSheet(f'color:{MUTED}; font-size:12px; background:transparent;')
        self.date_from = DatePickerButton(QDate.currentDate().addDays(-30))

        sep = QLabel('—')
        sep.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sep.setStyleSheet(f'color:{MUTED}; font-size:14px; background:transparent;')

        to_lbl = QLabel('עד:')
        to_lbl.setStyleSheet(f'color:{MUTED}; font-size:12px; background:transparent;')
        self.date_to = DatePickerButton(QDate.currentDate())

        prow.addWidget(from_lbl)
        prow.addWidget(self.date_from)
        prow.addWidget(sep)
        prow.addWidget(to_lbl)
        prow.addWidget(self.date_to)
        prow.addStretch()

        self._pickers.setVisible(False)
        root.addWidget(self._pickers)

    def _on_toggle(self, checked: bool):
        self._pickers.setVisible(checked)

    def is_active(self):
        return self.toggle.isChecked()

    def get_start(self):
        d = self.date_from.date()
        return datetime(d.year(), d.month(), d.day())

    def get_end(self):
        d = self.date_to.date()
        return datetime(d.year(), d.month(), d.day())


# ── Results card ──────────────────────────────────────────────────────────────

class ResultsCard(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setVisible(False)
        self._output_path: str | None = None
        self.setStyleSheet("""
            ResultsCard {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #ecfdf5, stop:1 #f0fdf4);
                border: 1.5px solid #6ee7b7;
                border-radius: 14px;
            }
        """)
        self.setGraphicsEffect(_shadow(16, 4, 25))

        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)

        self._stats = QLabel()
        self._stats.setAlignment(Qt.AlignmentFlag.AlignCenter)
        f = QFont(); f.setPointSize(13); f.setBold(True)
        self._stats.setFont(f)
        self._stats.setStyleSheet('color: #065f46; background: transparent;')

        self._open_btn = QPushButton('  פתח את הדוח  📂')
        self._open_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._open_btn.setMinimumHeight(40)
        self._open_btn.setStyleSheet(f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {SUCCESS}, stop:1 #10b981);
                color: white; border-radius: 9px;
                padding: 6px 28px; font-size: 14px; font-weight: bold;
            }}
            QPushButton:hover   {{ background: {SUCCESS_H}; }}
            QPushButton:pressed {{ background: #065f46; }}
        """)
        self._open_btn.clicked.connect(self._open)

        lay.addWidget(self._stats)
        lay.addWidget(self._open_btn, 0, Qt.AlignmentFlag.AlignCenter)

    def show_results(self, path: str, math_count: int, eng_count: int):
        self._output_path = path
        total = math_count + eng_count
        self._stats.setText(
            f'✓   נמצאו {total} פערים   ·   מתמטיקה: {math_count}   |   אנגלית: {eng_count}'
        )
        self.setVisible(True)

    def hide_results(self):
        self.setVisible(False)
        self._output_path = None

    def _open(self):
        if self._output_path:
            subprocess.run(['open', self._output_path], check=False)


# ── Worker ────────────────────────────────────────────────────────────────────

class Worker(QThread):
    finished = pyqtSignal(str, int, int, str)   # tmp_path, math, eng, drange
    error    = pyqtSignal(str)
    status   = pyqtSignal(str)
    progress = pyqtSignal(int)

    def __init__(self, darush, yitzua, use_filter, date_start, date_end, output):
        super().__init__()
        self.darush     = darush
        self.yitzua     = yitzua
        self.use_filter = use_filter
        self.date_start = date_start
        self.date_end   = date_end
        self.output     = output

    def run(self):
        try:
            tmp     = tempfile.mkdtemp()
            d_fixed = os.path.join(tmp, 'd.xlsx')
            y_fixed = os.path.join(tmp, 'y.xlsx')

            self.status.emit('מתקן קבצים…');   self.progress.emit(10)
            fix_xlsx(self.darush, d_fixed)
            fix_xlsx(self.yitzua, y_fixed)

            self.status.emit('טוען קבצים…');   self.progress.emit(25)
            wb1 = load_workbook(d_fixed)
            wb2 = load_workbook(y_fixed)

            self.status.emit('משווה…');         self.progress.emit(50)
            diffs, tfm, drange = run_comparison(wb1, wb2)

            if self.use_filter:
                diffs  = [d for d in diffs if self.date_start <= d['date'] <= self.date_end]
                drange = (f"{self.date_start.strftime('%d.%m.%Y')} — "
                          f"{self.date_end.strftime('%d.%m.%Y')}")

            if not diffs:
                self.error.emit('לא נמצאו פערים בטווח התאריכים שנבחר.')
                return

            self.status.emit('בונה דוח…');     self.progress.emit(80)
            build_report(diffs, tfm, drange, self.output)

            math_count = len([d for d in diffs if d['subject'] == 'מתמטיקה'])
            eng_count  = len([d for d in diffs if d['subject'] == 'אנגלית'])
            self.progress.emit(100)
            self.finished.emit(self.output, math_count, eng_count, drange)

        except Exception as exc:
            import traceback
            self.error.emit(f'שגיאה:\n{exc}\n\n{traceback.format_exc()}')


# ── Main window ───────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('מתכנן הפערים')
        self.setMinimumWidth(660)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self._settings = QSettings('TutorDiff', 'TutorDiff')
        self._build_ui()
        self._restore_last_files()

    def _build_ui(self):
        root_widget = QWidget()
        root_widget.setStyleSheet(f'background: {BG};')
        self.setCentralWidget(root_widget)
        root = QVBoxLayout(root_widget)
        root.setContentsMargins(28, 22, 28, 28)
        root.setSpacing(14)

        # ── Gradient header card ─────────────────────────────────────────
        header = QWidget()
        header.setStyleSheet(f"""
            QWidget {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 {PRIMARY}, stop:1 #2d5a8e);
                border-radius: 16px;
            }}
        """)
        header.setGraphicsEffect(_shadow(24, 7, 45))
        h_lay = QVBoxLayout(header)
        h_lay.setContentsMargins(24, 18, 24, 18)
        h_lay.setSpacing(4)

        title = QLabel('מתכנן הפערים')
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ft = QFont(); ft.setPointSize(22); ft.setBold(True)
        title.setFont(ft)
        title.setStyleSheet('color: #ffffff; background: transparent;')

        subtitle = QLabel('השווה קבצי Excel וצור דוח פערים מסודר')
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet('color: rgba(255,255,255,0.6); font-size: 12px; background: transparent;')

        h_lay.addWidget(title)
        h_lay.addWidget(subtitle)
        root.addWidget(header)

        # ── Drop zones ───────────────────────────────────────────────────
        zones_row = QHBoxLayout()
        zones_row.setSpacing(14)
        self.darush_zone = DropZone('דרוש תיקון', 'גרור לכאן את הקובץ לתיקון')
        self.yitzua_zone = DropZone('ייצוא שיטס',  'גרור לכאן את הקובץ התקין')
        self.darush_zone.file_changed.connect(lambda p: self._settings.setValue('darush_path', p))
        self.yitzua_zone.file_changed.connect(lambda p: self._settings.setValue('yitzua_path', p))
        self.darush_zone.cleared.connect(lambda: self._settings.setValue('darush_path', ''))
        self.yitzua_zone.cleared.connect(lambda: self._settings.setValue('yitzua_path', ''))
        zones_row.addWidget(self.darush_zone)
        zones_row.addWidget(self.yitzua_zone)
        root.addLayout(zones_row)

        # ── Date range card ──────────────────────────────────────────────
        self.date_card = DateRangeCard()
        root.addWidget(self.date_card)

        # ── Progress bar ─────────────────────────────────────────────────
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(5)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                background: #dde3ea; border-radius: 2px; border: none;
            }}
            QProgressBar::chunk {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {ACCENT}, stop:1 #93c5fd);
                border-radius: 2px;
            }}
        """)
        self.progress_bar.setVisible(False)
        root.addWidget(self.progress_bar)

        self.status_lbl = QLabel('')
        self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_lbl.setStyleSheet(f'color: {MUTED}; font-size: 12px; min-height: 16px;')
        root.addWidget(self.status_lbl)

        # ── Build button ─────────────────────────────────────────────────
        self.build_btn = QPushButton('✦   צור דוח')
        self.build_btn.setMinimumHeight(52)
        fb = QFont(); fb.setPointSize(16); fb.setBold(True)
        self.build_btn.setFont(fb)
        self.build_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.build_btn.setStyleSheet(f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {PRIMARY}, stop:1 {PRIMARY_H});
                color: white; border-radius: 13px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {PRIMARY_H}, stop:1 #3a6fa8);
            }}
            QPushButton:pressed {{ background: #0f2540; }}
            QPushButton:disabled {{ background: #94a3b8; color: #e2e8f0; }}
        """)
        self.build_btn.setGraphicsEffect(_shadow(12, 4, 35))
        self.build_btn.clicked.connect(self._run)
        root.addWidget(self.build_btn)

        # ── Results card ─────────────────────────────────────────────────
        self.results_card = ResultsCard()
        root.addWidget(self.results_card)

    # ── helpers ───────────────────────────────────────────────────────────────

    def _restore_last_files(self):
        for key, zone in [('darush_path', self.darush_zone),
                           ('yitzua_path', self.yitzua_zone)]:
            path = self._settings.value(key, '')
            if path and os.path.exists(path):
                zone.set_file(path)

    # ── run ───────────────────────────────────────────────────────────────────

    def _run(self):
        if not self.darush_zone.file_path:
            QMessageBox.warning(self, 'חסר קובץ', 'יש לבחור קובץ "דרוש תיקון".')
            return
        if not self.yitzua_zone.file_path:
            QMessageBox.warning(self, 'חסר קובץ', 'יש לבחור קובץ "ייצוא שיטס".')
            return

        # Build to a temp file first — we learn the date range only after processing
        tmp_output = os.path.join(tempfile.mkdtemp(), 'report.xlsx')

        self.results_card.hide_results()
        self.build_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.status_lbl.setText('מעבד…')

        self._worker = Worker(
            self.darush_zone.file_path,
            self.yitzua_zone.file_path,
            self.date_card.is_active(),
            date_start=self.date_card.get_start(),
            date_end=self.date_card.get_end(),
            output=tmp_output,
        )
        self._worker.finished.connect(self._on_done)
        self._worker.error.connect(self._on_error)
        self._worker.status.connect(self.status_lbl.setText)
        self._worker.progress.connect(self.progress_bar.setValue)
        self._worker.start()

    @staticmethod
    def _smart_filename(drange: str) -> str:
        """'01.05.2026 — 31.05.2026'  →  'דוח_פערים_01.05-31.05.xlsx'"""
        parts = [p.strip() for p in drange.split('—')]
        if len(parts) == 2:
            d1 = '.'.join(parts[0].split('.')[:2])   # "01.05"
            d2 = '.'.join(parts[1].split('.')[:2])   # "31.05"
            return f'דוח_פערים_{d1}-{d2}.xlsx'
        return 'דוח_פערים.xlsx'

    def _on_done(self, tmp_path: str, math_count: int, eng_count: int, drange: str):
        self.build_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_lbl.setText('')

        # Ask where to save, with smart filename derived from the actual date range
        from PyQt6.QtWidgets import QFileDialog
        import shutil
        filename = self._smart_filename(drange)
        default  = os.path.join(os.path.expanduser('~'), 'Downloads', filename)
        output, _ = QFileDialog.getSaveFileName(self, 'שמור דוח', default, 'Excel (*.xlsx)')
        if not output:
            try: os.remove(tmp_path)
            except OSError: pass
            return
        if not output.endswith('.xlsx'):
            output += '.xlsx'

        shutil.move(tmp_path, output)
        self.results_card.show_results(output, math_count, eng_count)
        self._mac_notify(math_count + eng_count)

    def _on_error(self, msg: str):
        self.build_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_lbl.setText('')
        QMessageBox.critical(self, 'שגיאה', msg)

    def _mac_notify(self, total: int):
        try:
            subprocess.run([
                'osascript', '-e',
                f'display notification "נמצאו {total} פערים — הדוח מוכן" '
                f'with title "מתכנן הפערים" sound name "Glass"'
            ], check=False)
        except Exception:
            pass


# ─────────────────────────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
    app.setStyle('Fusion')
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
