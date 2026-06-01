#!/usr/bin/env python3
"""
Tutor Diff — Desktop GUI
PyQt6 wrapper around build_excel.py.

The GUI only calls run_comparison() and build_report() from build_excel.py —
all skill logic lives there and can be updated independently.
"""

import sys
import os
import tempfile
from datetime import datetime
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QDateEdit, QMessageBox, QFrame,
    QSizePolicy, QCheckBox, QGroupBox,
)
from PyQt6.QtCore import Qt, QDate, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QDragEnterEvent, QDropEvent

from openpyxl import load_workbook
from build_excel import run_comparison, build_report, fix_xlsx


# ─────────────────────────────────────────────────────────────────────────────
# Drop zone widget
# ─────────────────────────────────────────────────────────────────────────────

class DropZone(QFrame):
    file_changed = pyqtSignal(str)

    def __init__(self, title: str, subtitle: str, parent=None):
        super().__init__(parent)
        self.file_path: str | None = None
        self.setAcceptDrops(True)
        self.setMinimumHeight(150)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self._build(title, subtitle)
        self._style(False)

    def _build(self, title, subtitle):
        lay = QVBoxLayout(self)
        lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.setSpacing(8)
        lay.setContentsMargins(16, 16, 16, 16)

        self._icon = QLabel('📂')
        self._icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._icon.setFont(QFont('', 28))

        self._title = QLabel(title)
        self._title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        f = QFont()
        f.setBold(True)
        f.setPointSize(13)
        self._title.setFont(f)

        self._sub = QLabel(subtitle)
        self._sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._sub.setStyleSheet('color: #888; font-size: 11px;')

        self._file = QLabel('')
        self._file.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._file.setStyleSheet('color: #2a7ae2; font-size: 11px;')
        self._file.setWordWrap(True)

        btn = QPushButton('בחר קובץ')
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setStyleSheet("""
            QPushButton {
                background:#f0f4f8; border:1px solid #ccd;
                border-radius:4px; padding:4px 14px; font-size:12px;
            }
            QPushButton:hover { background:#e0e8f0; }
        """)
        btn.clicked.connect(self._choose)

        for w in (self._icon, self._title, self._sub, self._file):
            lay.addWidget(w)
        lay.addWidget(btn, 0, Qt.AlignmentFlag.AlignCenter)

    def _style(self, has_file: bool):
        if has_file:
            self.setStyleSheet("""
                DropZone {
                    border:2px solid #2a7ae2;
                    border-radius:10px;
                    background:#f0f7ff;
                }
            """)
            self._icon.setText('✅')
        else:
            self.setStyleSheet("""
                DropZone {
                    border:2px dashed #bbb;
                    border-radius:10px;
                    background:#fafafa;
                }
            """)
            self._icon.setText('📂')

    def set_file(self, path: str):
        self.file_path = path
        name = Path(path).name
        self._file.setText(name if len(name) <= 38 else name[:35] + '…')
        self._style(True)
        self.file_changed.emit(path)

    def _choose(self):
        from PyQt6.QtWidgets import QFileDialog
        path, _ = QFileDialog.getOpenFileName(
            self, 'בחר קובץ Excel', '', 'Excel (*.xlsx *.xls)'
        )
        if path:
            self.set_file(path)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0].toLocalFile().lower()
            if url.endswith(('.xlsx', '.xls')):
                event.acceptProposedAction()
                self.setStyleSheet("""
                    DropZone {
                        border:2px dashed #2a7ae2;
                        border-radius:10px;
                        background:#e8f0fe;
                    }
                """)

    def dragLeaveEvent(self, event):
        self._style(self.file_path is not None)

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            self.set_file(urls[0].toLocalFile())


# ─────────────────────────────────────────────────────────────────────────────
# Background worker
# ─────────────────────────────────────────────────────────────────────────────

class Worker(QThread):
    finished = pyqtSignal(str)
    error    = pyqtSignal(str)
    status   = pyqtSignal(str)

    def __init__(self, darush: str, yitzua: str,
                 use_filter: bool, date_start: datetime, date_end: datetime,
                 output: str):
        super().__init__()
        self.darush     = darush
        self.yitzua     = yitzua
        self.use_filter = use_filter
        self.date_start = date_start
        self.date_end   = date_end
        self.output     = output

    def run(self):
        try:
            tmp = tempfile.mkdtemp()
            d_fixed = os.path.join(tmp, 'd.xlsx')
            y_fixed = os.path.join(tmp, 'y.xlsx')

            self.status.emit('מתקן קבצים…')
            fix_xlsx(self.darush, d_fixed)
            fix_xlsx(self.yitzua, y_fixed)

            self.status.emit('טוען קבצים…')
            wb1 = load_workbook(d_fixed)
            wb2 = load_workbook(y_fixed)

            self.status.emit('משווה…')
            diffs, tfm, drange = run_comparison(wb1, wb2)

            if self.use_filter:
                diffs  = [d for d in diffs if self.date_start <= d['date'] <= self.date_end]
                drange = (f"{self.date_start.strftime('%d.%m.%Y')} — "
                          f"{self.date_end.strftime('%d.%m.%Y')}")

            if not diffs:
                self.error.emit('לא נמצאו פערים בטווח התאריכים שנבחר.')
                return

            self.status.emit('בונה דוח…')
            build_report(diffs, tfm, drange, self.output)
            self.finished.emit(self.output)

        except Exception as exc:
            import traceback
            self.error.emit(f'שגיאה:\n{exc}\n\n{traceback.format_exc()}')


# ─────────────────────────────────────────────────────────────────────────────
# Main window
# ─────────────────────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('דוח פערים מתרגלים')
        self.setMinimumWidth(620)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self._build_ui()

    def _build_ui(self):
        root_widget = QWidget()
        self.setCentralWidget(root_widget)
        root = QVBoxLayout(root_widget)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(18)

        # Title
        title = QLabel('דוח פערים מתרגלים')
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        f = QFont()
        f.setPointSize(22)
        f.setBold(True)
        title.setFont(f)
        title.setStyleSheet('color: #1a3c5e; margin-bottom: 4px;')
        root.addWidget(title)

        # Drop zones (RTL: darush = right, yitzua = left)
        zones_row = QHBoxLayout()
        zones_row.setSpacing(16)
        self.darush_zone = DropZone('דרוש תיקון', 'גרור לכאן את הקובץ לתיקון')
        self.yitzua_zone = DropZone('ייצוא שיטס', 'גרור לכאן את הקובץ התקין')
        zones_row.addWidget(self.darush_zone)
        zones_row.addWidget(self.yitzua_zone)
        root.addLayout(zones_row)

        # Date range group
        date_group = QGroupBox('סינון לפי טווח תאריכים (אופציונלי)')
        date_group.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        date_row = QHBoxLayout(date_group)
        date_row.setSpacing(10)

        self.use_dates = QCheckBox('פעיל')
        self.use_dates.stateChanged.connect(self._toggle_dates)

        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate.currentDate().addDays(-30))
        self.date_from.setEnabled(False)
        self.date_from.setDisplayFormat('dd.MM.yyyy')

        dash = QLabel('—')
        dash.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate())
        self.date_to.setEnabled(False)
        self.date_to.setDisplayFormat('dd.MM.yyyy')

        date_row.addWidget(self.use_dates)
        date_row.addWidget(self.date_from)
        date_row.addWidget(dash)
        date_row.addWidget(self.date_to)
        date_row.addStretch()
        root.addWidget(date_group)

        # Status
        self.status_lbl = QLabel('')
        self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_lbl.setStyleSheet('color: #555; font-size: 12px; min-height: 18px;')
        root.addWidget(self.status_lbl)

        # Build button
        self.build_btn = QPushButton('צור דוח')
        self.build_btn.setMinimumHeight(46)
        f2 = QFont()
        f2.setPointSize(15)
        f2.setBold(True)
        self.build_btn.setFont(f2)
        self.build_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.build_btn.setStyleSheet("""
            QPushButton {
                background:#1a3c5e; color:white;
                border-radius:8px;
            }
            QPushButton:hover  { background:#2a5a8e; }
            QPushButton:pressed { background:#0f2540; }
            QPushButton:disabled { background:#aaa; color:#eee; }
        """)
        self.build_btn.clicked.connect(self._run)
        root.addWidget(self.build_btn)

    def _toggle_dates(self, state):
        on = state == Qt.CheckState.Checked.value
        self.date_from.setEnabled(on)
        self.date_to.setEnabled(on)

    def _run(self):
        darush = self.darush_zone.file_path
        yitzua = self.yitzua_zone.file_path
        if not darush:
            QMessageBox.warning(self, 'חסר קובץ', 'יש לבחור קובץ "דרוש תיקון".')
            return
        if not yitzua:
            QMessageBox.warning(self, 'חסר קובץ', 'יש לבחור קובץ "ייצוא שיטס".')
            return

        from PyQt6.QtWidgets import QFileDialog
        default_path = os.path.join(os.path.expanduser('~'), 'Desktop', 'דוח_פערים.xlsx')
        output, _ = QFileDialog.getSaveFileName(
            self, 'שמור דוח', default_path, 'Excel (*.xlsx)'
        )
        if not output:
            return
        if not output.endswith('.xlsx'):
            output += '.xlsx'

        q_from = self.date_from.date()
        q_to   = self.date_to.date()
        start  = datetime(q_from.year(), q_from.month(), q_from.day())
        end    = datetime(q_to.year(),   q_to.month(),   q_to.day())

        self.build_btn.setEnabled(False)
        self.status_lbl.setText('מעבד…')

        self._worker = Worker(
            darush, yitzua,
            self.use_dates.isChecked(),
            date_start=start, date_end=end,
            output=output,
        )
        self._worker.finished.connect(self._on_done)
        self._worker.error.connect(self._on_error)
        self._worker.status.connect(self.status_lbl.setText)
        self._worker.start()

    def _on_done(self, path: str):
        self.build_btn.setEnabled(True)
        self.status_lbl.setText('הדוח נשמר בהצלחה ✓')
        QMessageBox.information(self, 'הצלחה', f'הדוח נשמר:\n{path}')

    def _on_error(self, msg: str):
        self.build_btn.setEnabled(True)
        self.status_lbl.setText('שגיאה')
        QMessageBox.critical(self, 'שגיאה', msg)


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
