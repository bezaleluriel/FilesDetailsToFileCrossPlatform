import sys
from pathlib import Path

from PySide6.QtCore import QObject, QThread, Signal, Slot, QMimeData
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core import HEAD_COLS, export_docx, rows_to_html_document, rows_to_html_table, rows_to_tsv, scan_folder


class ScanWorker(QObject):
    finished = Signal(list)
    error = Signal(str)

    def __init__(self, folder: str) -> None:
        super().__init__()
        self.folder = folder

    @Slot()
    def run(self) -> None:
        try:
            rows = scan_folder(self.folder)
        except Exception as e:  # noqa: BLE001
            self.error.emit(str(e))
            return
        self.finished.emit(rows)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Folder File Table → Word (Windows + macOS)")
        self.resize(1120, 720)

        self.rows = []

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        top = QHBoxLayout()
        layout.addLayout(top)

        self.folder_edit = QLineEdit()
        self.folder_edit.setPlaceholderText("Choose a folder…")
        top.addWidget(self.folder_edit, 1)

        browse_btn = QPushButton("Browse…")
        browse_btn.clicked.connect(self.browse)
        top.addWidget(browse_btn)

        self.scan_btn = QPushButton("Scan")
        self.scan_btn.clicked.connect(self.scan)
        top.addWidget(self.scan_btn)

        actions = QHBoxLayout()
        layout.addLayout(actions)

        self.copy_btn = QPushButton("Copy (Word table)")
        self.copy_btn.clicked.connect(self.copy_word)
        actions.addWidget(self.copy_btn)

        export_html_btn = QPushButton("Export HTML…")
        export_html_btn.clicked.connect(self.export_html)
        actions.addWidget(export_html_btn)

        export_docx_btn = QPushButton("Export DOCX…")
        export_docx_btn.clicked.connect(self.export_docx)
        actions.addWidget(export_docx_btn)

        actions.addStretch(1)
        self.status_lbl = QLabel("Choose a folder to scan.")
        actions.addWidget(self.status_lbl)

        self.table = QTableWidget(0, len(HEAD_COLS))
        self.table.setHorizontalHeaderLabels(HEAD_COLS)
        self.table.setSortingEnabled(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        layout.addWidget(self.table, 1)

        self._thread: QThread | None = None
        self._worker: ScanWorker | None = None

        self._set_busy(False)

    def _set_busy(self, busy: bool) -> None:
        self.scan_btn.setDisabled(busy)
        self.copy_btn.setDisabled(busy or not self.rows)

    @Slot()
    def browse(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Choose folder to scan")
        if folder:
            self.folder_edit.setText(folder)
            self.status_lbl.setText("Ready to scan.")

    @Slot()
    def scan(self) -> None:
        folder = self.folder_edit.text().strip()
        if not folder:
            QMessageBox.warning(self, "No folder", "Please choose a folder first.")
            return
        if not Path(folder).is_dir():
            QMessageBox.critical(self, "Invalid folder", "That folder does not exist.")
            return

        self.status_lbl.setText("Scanning…")
        self.rows = []
        self._set_busy(True)
        self.table.setRowCount(0)

        self._thread = QThread(self)
        self._worker = ScanWorker(folder)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._scan_finished)
        self._worker.error.connect(self._scan_error)
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.start()

    @Slot(list)
    def _scan_finished(self, rows: list) -> None:
        self.rows = rows
        self._populate_table()
        self.status_lbl.setText(f"Found {len(self.rows)} files.")
        self._set_busy(False)

    @Slot(str)
    def _scan_error(self, msg: str) -> None:
        QMessageBox.critical(self, "Scan failed", msg)
        self.status_lbl.setText("Scan failed.")
        self._set_busy(False)

    def _populate_table(self) -> None:
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(self.rows))
        for i, r in enumerate(self.rows):
            vals = [
                r.file_name,
                r.last_2_paths,
                r.inner_path,
                r.run_v,
                r.data_v,
                r.created,
                str(r.size_bytes),
            ]
            for j, v in enumerate(vals):
                item = QTableWidgetItem(v)
                self.table.setItem(i, j, item)
        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(True)
        self._set_busy(False)

    @Slot()
    def copy_word(self) -> None:
        if not self.rows:
            QMessageBox.information(self, "Nothing to copy", "Scan a folder first.")
            return

        tsv = rows_to_tsv(self.rows)
        # Word pastes a real table when HTML is present.
        html_doc = (
            "<html><head><meta charset='utf-8'></head><body>"
            + rows_to_html_table(self.rows)
            + "</body></html>"
        )

        md = QMimeData()
        md.setHtml(html_doc)
        md.setText(tsv)
        QGuiApplication.clipboard().setMimeData(md)
        self.status_lbl.setText("Copied table to clipboard (HTML + TSV).")

    @Slot()
    def export_html(self) -> None:
        if not self.rows:
            QMessageBox.information(self, "Nothing to export", "Scan a folder first.")
            return
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Export HTML table", "", "HTML file (*.html)"
        )
        if not out_path:
            return
        title = f"File table - {self.folder_edit.text().strip()}"
        try:
            Path(out_path).write_text(rows_to_html_document(self.rows, title), encoding="utf-8")
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Export failed", str(e))
            return
        self.status_lbl.setText(f"Exported HTML: {out_path}")

    @Slot()
    def export_docx(self) -> None:
        if not self.rows:
            QMessageBox.information(self, "Nothing to export", "Scan a folder first.")
            return
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Export Word document", "", "Word document (*.docx)"
        )
        if not out_path:
            return
        title = f"File table - {self.folder_edit.text().strip()}"
        try:
            export_docx(self.rows, out_path=out_path, title=title)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Export failed", str(e))
            return
        self.status_lbl.setText(f"Exported DOCX: {out_path}")


def main() -> int:
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

