"""
main.py
-------
The whole app is one screen:

    1. Drag files in (or click "Choose files")   -- PDF / CSV / Excel
    2. (optional) tick "Also create a Tally file" and fill the two boxes
    3. Click "Convert"
    4. A friendly result appears with an "Open output folder" button.

Outputs are written to a "Bank Statements" folder on the Desktop so a
non-technical user never has to decide where things go.

Run for development:   python main.py
Build a Windows .exe:  see README.md
"""

from __future__ import annotations

import os
import sys
import logging

from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QListWidget, QFileDialog, QProgressBar, QCheckBox, QLineEdit, QFrame,
    QMessageBox, QGroupBox, QFormLayout,
)

from core import pipeline

logging.basicConfig(level=logging.INFO,
                    format="%(levelname)s %(name)s: %(message)s")

SUPPORTED = (".pdf", ".csv", ".tsv", ".xlsx", ".xlsm", ".xls", ".xltx")


def default_output_dir() -> str:
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    base = desktop if os.path.isdir(desktop) else os.path.expanduser("~")
    return os.path.join(base, "Bank Statements")


def app_base_dir() -> str:
    """Folder next to the exe (or this script) — where config/ledger_map.csv lives."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# --------------------------------------------------------------------------- #
# Background worker (keeps the UI responsive during conversion)
# --------------------------------------------------------------------------- #
class Worker(QObject):
    progress = Signal(int, str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, files, output_dir, make_tally, company, bank, ledger_map):
        super().__init__()
        self.files = files
        self.output_dir = output_dir
        self.make_tally = make_tally
        self.company = company
        self.bank = bank
        self.ledger_map = ledger_map

    def start(self):
        try:
            res = pipeline.run(
                files=self.files,
                output_dir=self.output_dir,
                make_tally=self.make_tally,
                company_name=self.company,
                bank_ledger=self.bank,
                ledger_map_path=self.ledger_map,
                progress=lambda p, m: self.progress.emit(p, m),
            )
            self.finished.emit(res)
        except Exception as exc:  # noqa: BLE001 - surface any failure kindly
            logging.exception("Conversion failed")
            self.failed.emit(str(exc))


# --------------------------------------------------------------------------- #
# Drag-and-drop file list
# --------------------------------------------------------------------------- #
class DropList(QListWidget):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setSelectionMode(QListWidget.ExtendedSelection)
        self.setMinimumHeight(150)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        for url in e.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(SUPPORTED):
                self.add_unique(path)

    def add_unique(self, path):
        existing = {self.item(i).data(Qt.UserRole) for i in range(self.count())}
        if path not in existing:
            from PySide6.QtWidgets import QListWidgetItem
            item = QListWidgetItem(os.path.basename(path))
            item.setData(Qt.UserRole, path)
            self.addItem(item)

    def all_paths(self):
        return [self.item(i).data(Qt.UserRole) for i in range(self.count())]


# --------------------------------------------------------------------------- #
# Main window
# --------------------------------------------------------------------------- #
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bank Statement to Excel")
        self.setMinimumWidth(560)
        self.output_dir = default_output_dir()
        self._thread = None
        self._worker = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(14)

        title = QLabel("Bank Statement to Excel")
        title.setFont(QFont("Segoe UI", 18, QFont.Bold))
        layout.addWidget(title)

        hint = QLabel("Drag your PDF, CSV or Excel statements below, "
                      "or click “Choose files”.")
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#555;")
        layout.addWidget(hint)

        self.filelist = DropList()
        layout.addWidget(self.filelist)

        btn_row = QHBoxLayout()
        choose = QPushButton("Choose files…")
        choose.clicked.connect(self.choose_files)
        remove = QPushButton("Remove selected")
        remove.clicked.connect(self.remove_selected)
        clear = QPushButton("Clear all")
        clear.clicked.connect(self.filelist.clear)
        btn_row.addWidget(choose)
        btn_row.addWidget(remove)
        btn_row.addWidget(clear)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        # ---- Tally options (collapsible-ish group) ---- #
        self.tally_chk = QCheckBox("Also create a Tally import file")
        self.tally_chk.stateChanged.connect(self._toggle_tally)
        layout.addWidget(self.tally_chk)

        self.tally_box = QGroupBox("Tally details")
        form = QFormLayout(self.tally_box)
        self.company_edit = QLineEdit()
        self.company_edit.setPlaceholderText("Company name exactly as in Tally")
        self.bank_edit = QLineEdit()
        self.bank_edit.setPlaceholderText("Bank ledger name exactly as in Tally")
        form.addRow("Company:", self.company_edit)
        form.addRow("Bank ledger:", self.bank_edit)
        note = QLabel("Every payee ledger must already exist in Tally. Edit "
                      "config/ledger_map.csv to map payee names to ledgers.")
        note.setWordWrap(True)
        note.setStyleSheet("color:#777; font-size:11px;")
        form.addRow(note)
        self.tally_box.setVisible(False)
        layout.addWidget(self.tally_box)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("color:#ddd;")
        layout.addWidget(line)

        # ---- convert + progress ---- #
        self.convert_btn = QPushButton("Convert")
        self.convert_btn.setMinimumHeight(44)
        self.convert_btn.setStyleSheet(
            "QPushButton{background:#2E5496;color:white;font-size:15px;"
            "font-weight:bold;border-radius:6px;}"
            "QPushButton:disabled{background:#9db3d6;}"
        )
        self.convert_btn.clicked.connect(self.convert)
        layout.addWidget(self.convert_btn)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        self.status = QLabel("")
        self.status.setWordWrap(True)
        layout.addWidget(self.status)

        self.open_btn = QPushButton("Open output folder")
        self.open_btn.setVisible(False)
        self.open_btn.clicked.connect(self.open_output)
        layout.addWidget(self.open_btn)

    # ---- actions ---- #
    def _toggle_tally(self):
        self.tally_box.setVisible(self.tally_chk.isChecked())

    def choose_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Choose statement files", "",
            "Statements (*.pdf *.csv *.tsv *.xlsx *.xlsm *.xls);;All files (*.*)",
        )
        for p in paths:
            self.filelist.add_unique(p)

    def remove_selected(self):
        for item in self.filelist.selectedItems():
            self.filelist.takeItem(self.filelist.row(item))

    def convert(self):
        files = self.filelist.all_paths()
        if not files:
            QMessageBox.information(self, "No files",
                                    "Please add at least one statement file.")
            return
        make_tally = self.tally_chk.isChecked()
        if make_tally and not self.bank_edit.text().strip():
            QMessageBox.information(
                self, "Bank ledger needed",
                "Enter the bank ledger name (exactly as it appears in Tally).")
            return

        self.convert_btn.setEnabled(False)
        self.open_btn.setVisible(False)
        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.status.setText("")

        ledger_map = os.path.join(app_base_dir(), "config", "ledger_map.csv")

        self._thread = QThread()
        self._worker = Worker(
            files=files,
            output_dir=self.output_dir,
            make_tally=make_tally,
            company=self.company_edit.text().strip(),
            bank=self.bank_edit.text().strip(),
            ledger_map=ledger_map,
        )
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.start)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self._thread.start()

    def _on_progress(self, pct, msg):
        self.progress.setValue(pct)
        self.status.setText(msg)

    def _on_finished(self, res):
        self._cleanup_thread()
        self.convert_btn.setEnabled(True)
        self.progress.setValue(100)
        if res.excel_path:
            self.status.setText(
                f"✅ Done — {res.message}\nSaved to: {self.output_dir}")
            self.open_btn.setVisible(True)
        else:
            self.status.setText("⚠️ " + (res.message or "Nothing was created."))

    def _on_failed(self, err):
        self._cleanup_thread()
        self.convert_btn.setEnabled(True)
        self.progress.setVisible(False)
        QMessageBox.critical(
            self, "Something went wrong",
            "Sorry, the conversion failed.\n\nDetails: " + err)

    def _cleanup_thread(self):
        if self._thread:
            self._thread.quit()
            self._thread.wait()
            self._thread = None
            self._worker = None

    def open_output(self):
        os.makedirs(self.output_dir, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(self.output_dir)  # noqa: SIM115 (Windows-only)
        elif sys.platform == "darwin":
            os.system(f'open "{self.output_dir}"')
        else:
            os.system(f'xdg-open "{self.output_dir}"')


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
