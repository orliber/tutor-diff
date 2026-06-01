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
import subprocess
from datetime import datetime
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QDateEdit, QMessageBox, QFrame,
    QSizePolicy, QCheckBox, QGroupBox, QProgressBar,
)
from PyQt6.QtCore import Qt, QDate, QThread, pyqtSignal, QSettings
from PyQt6.QtGui import QFont, QDragEnterEvent, QDropEvent

from openpyxl import load_workbook
from build_excel import run_comparison, build_report, fix_xlsx


# ─────────────────────────────────────────────────────────────────────────────
# Drop zone widget
# ─────────────────────────────────────────────────────────────────────────────

class DropZone(QFrame):
    file_changed = pyqtSignal(str)
    cleared = pyqtSignal()

    def __init__(self, title: str, subtitle: str, parent=None):
        super().__init__(parent)
        self.file_path: str | None = None
        self.setAcceptDrops(True)
        self.setMinimumHeight(165)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self._build(title, subtitle)
        self._style(False)

    def _build(self, title, subtitle):
        lay = QVBoxLayout(self)
        lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.setSpacing(5)
        lay.setContentsMargins(14, 14, 14, 14)

        self._icon = QLabel('📂')
        self._icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._icon.setFont(QFont('', 30))

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
        self._file.setStyleSheet('color: #2a7ae2; font-size: 11px; font-weight: bold;')
        self._file.setWordWrap(True)

        self._meta = QLabel('')
        self._meta.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._meta.setStyleSheet('color: #aaa; font-size: 10px;')

        btn_row = QHBoxLayout()
        btn_row.setSpacing(6)
        btn_row.addStretch()

        self._choose_btn = QPushButton('בחר קובץ')
        self._choose_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._choose_btn.setStyleSheet("""
            QPushButton {
                background: #f0f4f8; border: 1px solid #ccd;
                border-radius: 5px; padding: 4px 14px; font-size: 12px;
            }
            QPushButton:hover { background: #e0e8f0; }
        """)
        self._choose_btn.clicked.connect(self._choose)

        self._clear_btn = QPushButton('✕')
        self._clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._clear_btn.setFixedSize(26, 26)
        self._clear_btn.setToolTip('נקה קובץ')
        self._clear_btn.setStyleSheet("""
            QPushButton {
                background: #fee2e2; border: 1px solid #fca5a5;
                border-radius: 5px; font-size: 11px; color: #dc2626;
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

    def _style(self, has_file: bool):
        if has_file:
            self.setStyleSheet("""
                DropZone {
                    border: 2px solid #2a7ae2;
                    border-radius: 12px;
                    background: #f0f7ff;
                }
            """)
            self._icon.setText('✅')
        else:
            self.setStyleSheet("""
                DropZone {
                    border: 2px dashed #c0c8d4;
                    border-radius: 12px;
                    background: #f8fafc;
                }
            """)
            self._icon.setText('📂')

    def set_file(self, path: str):
        self.file_path = path
        name = Path(path).name
        self._file.setText(name if len(name) <= 36 else name[:33] + '…')
        try:
            kb   = os.path.getsize(path) // 1024
            date = datetime.fromtimestamp(os.path.getmtime(path)).strftime('%d.%m.%Y')
            self._meta.setText(f'{kb:,} KB  •  {date}')
        except Exception:
            self._meta.setText('')
        self._clear_btn.setVisible(True)
        self._style(True)
        self.file_changed.emit(path)

    def _clear(self):
        self.file_path = None
        self._file.setText('')
        self._meta.setText('')
        self._clear_btn.setVisible(False)
        self._style(False)
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
                self.setStyleSheet("""
                    DropZone {
                        border: 2px dashed #2a7ae2;
                        border-radius: 12px;
                        background: #dbeafe;
                    }
                """)

    def dragLeaveEvent(self, event):
        self._style(self.file_path is not None)

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            self.set_file(urls[0].toLocalFile())


# ─────────────────────────────────────────────────────────────────────────────
# Results card
# ─────────────────────────────────────────────────────────────────────────────

class ResultsCard(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setVisible(False)
        self._output_path: str | None = None
        self.setStyleSheet("""
            ResultsCard {
                background: #f0fdf4;
                border: 1.5px solid #86efac;
                border-radius: 10px;
            }
        """)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 12, 16, 12)
        lay.setSpacing(8)

        self._stats = QLabel()
        self._stats.setAlignment(Qt.AlignmentFlag.AlignCenter)
        f = QFont()
        f.setPointSize(13)
        f.setBold(True)
        self._stats.setFont(f)
        self._stats.setStyleSheet('color: #15803d;')

        self._open_btn = QPushButton('📂   פתח את הדוח')
        self._open_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._open_btn.setMinimumHeight(36)
        self._open_btn.setStyleSheet("""
            QPushButton {
                background: #16a34a; color: white;
                border-radius: 7px; padding: 6px 24px;
                font-size: 13px; font-weight: bold;
            }
            QPushButton:hover  { background: #15803d; }
            QPushButton:pressed { background: #166534; }
        """)
        self._open_btn.clicked.connect(self._open)

        lay.addWidget(self._stats)
        lay.addWidget(self._open_btn, 0, Qt.AlignmentFlag.AlignCenter)

    def show_results(self, path: str, math_count: int, eng_count: int):
        self._output_path = path
        total = math_count + eng_count
        self._stats.setText(
            f'✓  נמצאו {total} פערים  ·  מתמטיקה: {math_count}  |  אנגלית: {eng_count}'
        )
        self.setVisible(True)

    def hide_results(self):
        self.setVisible(False)
        self._output_path = None

    def _open(self):
        if self._output_path:
            subprocess.run(['open', self._output_path], check=False)


# ─────────────────────────────────────────────────────────────────────────────
# Background worker
# ─────────────────────────────────────────────────────────────────────────────

class Worker(QThread):
    finished = pyqtSignal(str, int, int)   # path, math_count, eng_count
    error    = pyqtSignal(str)
    status   = pyqtSignal(str)
    progress = pyqtSignal(int)             # 0–100

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
            self.finished.emit(self.output, math_count, eng_count)

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
        self.setMinimumWidth(640)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self._settings = QSettings('TutorDiff', 'TutorDiff')
        self._build_ui()
        self._restore_last_files()

    def _build_ui(self):
        root_widget = QWidget()
        root_widget.setStyleSheet('background: #f1f5f9;')
        self.setCentralWidget(root_widget)
        root = QVBoxLayout(root_widget)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(14)

        # Title
        title = QLabel('דוח פערים מתרגלים')
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        f = QFont()
        f.setPointSize(22)
        f.setBold(True)
        title.setFont(f)
        title.setStyleSheet('color: #1a3c5e; margin-bottom: 2px;')
        root.addWidget(title)

        # Drop zones (RTL: darush = right, yitzua = left)
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

        # Date range
        date_group = QGroupBox('סינון לפי טווח תאריכים (אופציונלי)')
        date_group.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        date_group.setStyleSheet("""
            QGroupBox {
                background: white;
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                margin-top: 8px;
                padding: 8px 4px 4px 4px;
                font-size: 12px;
                color: #64748b;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                right: 12px; top: 0px;
            }
        """)
        date_row = QHBoxLayout(date_group)
        date_row.setSpacing(10)

        self.use_dates = QCheckBox('פעיל')
        self.use_dates.stateChanged.connect(self._toggle_dates)

        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate.currentDate().addDays(-30))
        self.date_from.setEnabled(False)
        self.date_from.setDisplayFormat('dd.MM.yyyy')

        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate())
        self.date_to.setEnabled(False)
        self.date_to.setDisplayFormat('dd.MM.yyyy')

        date_row.addWidget(self.use_dates)
        date_row.addWidget(self.date_from)
        date_row.addWidget(QLabel('—'))
        date_row.addWidget(self.date_to)
        date_row.addStretch()
        root.addWidget(date_group)

        # Progress bar (hidden until processing)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(6)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                background: #e2e8f0; border-radius: 3px; border: none;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #2a7ae2, stop:1 #60a5fa);
                border-radius: 3px;
            }
        """)
        self.progress_bar.setVisible(False)
        root.addWidget(self.progress_bar)

        # Status label
        self.status_lbl = QLabel('')
        self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_lbl.setStyleSheet('color: #94a3b8; font-size: 12px; min-height: 16px;')
        root.addWidget(self.status_lbl)

        # Build button
        self.build_btn = QPushButton('צור דוח')
        self.build_btn.setMinimumHeight(50)
        f2 = QFont()
        f2.setPointSize(16)
        f2.setBold(True)
        self.build_btn.setFont(f2)
        self.build_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.build_btn.setStyleSheet("""
            QPushButton {
                background: #1a3c5e; color: white;
                border-radius: 10px;
            }
            QPushButton:hover   { background: #2a5a8e; }
            QPushButton:pressed { background: #0f2540; }
            QPushButton:disabled { background: #94a3b8; color: #e2e8f0; }
        """)
        self.build_btn.clicked.connect(self._run)
        root.addWidget(self.build_btn)

        # Results card (hidden until done)
        self.results_card = ResultsCard()
        root.addWidget(self.results_card)

    # ── helpers ───────────────────────────────────────────────────────────────

    def _toggle_dates(self, state):
        on = state == Qt.CheckState.Checked.value
        self.date_from.setEnabled(on)
        self.date_to.setEnabled(on)

    def _restore_last_files(self):
        for key, zone in [('darush_path', self.darush_zone),
                           ('yitzua_path', self.yitzua_zone)]:
            path = self._settings.value(key, '')
            if path and os.path.exists(path):
                zone.set_file(path)

    # ── run ───────────────────────────────────────────────────────────────────

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
        default = os.path.join(os.path.expanduser('~'), 'Desktop', 'דוח_פערים.xlsx')
        output, _ = QFileDialog.getSaveFileName(self, 'שמור דוח', default, 'Excel (*.xlsx)')
        if not output:
            return
        if not output.endswith('.xlsx'):
            output += '.xlsx'

        q_from = self.date_from.date()
        q_to   = self.date_to.date()

        self.results_card.hide_results()
        self.build_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.status_lbl.setText('מעבד…')

        self._worker = Worker(
            darush, yitzua,
            self.use_dates.isChecked(),
            date_start=datetime(q_from.year(), q_from.month(), q_from.day()),
            date_end=datetime(q_to.year(),   q_to.month(),   q_to.day()),
            output=output,
        )
        self._worker.finished.connect(self._on_done)
        self._worker.error.connect(self._on_error)
        self._worker.status.connect(self.status_lbl.setText)
        self._worker.progress.connect(self.progress_bar.setValue)
        self._worker.start()

    def _on_done(self, path: str, math_count: int, eng_count: int):
        self.build_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_lbl.setText('')
        self.results_card.show_results(path, math_count, eng_count)
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
                f'with title "דוח פערים מתרגלים" sound name "Glass"'
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
