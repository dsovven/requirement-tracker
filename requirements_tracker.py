"""
Requirements Tracker - PDF Requirement Capture Tool

Opens PDF files, allows rectangular screenshot capture of requirements,
stamps requirement numbers on the PDF, and generates a tracking document.
"""

import sys
import os
from datetime import datetime
from io import BytesIO
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

import fitz  # PyMuPDF

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGraphicsView, QGraphicsScene, QToolBar, QAction, QFileDialog,
    QLabel, QLineEdit, QCheckBox, QListWidget, QListWidgetItem,
    QSplitter, QMessageBox, QPushButton, QScrollArea, QSizePolicy,
    QStyle, QGroupBox, QFormLayout
)
from PyQt5.QtCore import (
    Qt, QRectF, QRect, QPoint, QSize, pyqtSignal, QBuffer, QByteArray
)
from PyQt5.QtGui import (
    QPixmap, QImage, QPainter, QPen, QColor, QBrush, QFont,
    QKeySequence, QIcon
)

try:
    from docx import Document as DocxDocument
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------

@dataclass
class Requirement:
    number: str          # "1", "2", "7.1", etc.
    screenshot: QPixmap  # cropped image of the requirement area
    page: int            # 0-based page index
    pdf_rect: tuple      # (x0, y0, x1, y1) in PDF points


def pixmap_to_bytes(pixmap: QPixmap) -> BytesIO:
    """Convert a QPixmap to a BytesIO PNG stream."""
    ba = QByteArray()
    buf = QBuffer(ba)
    buf.open(QBuffer.WriteOnly)
    pixmap.save(buf, "PNG")
    buf.close()
    bio = BytesIO(ba.data())
    bio.seek(0)
    return bio


# ---------------------------------------------------------------------------
# PDF Page Widget  (handles rendering + rectangle drawing)
# ---------------------------------------------------------------------------

class PDFPageWidget(QWidget):
    """Displays a single PDF page pixmap and lets the user draw rectangles."""

    selection_made = pyqtSignal(QRectF)   # rectangle in *pixmap* coordinates
    zoom_requested = pyqtSignal(int)      # wheel delta for zoom

    def __init__(self, parent=None):
        super().__init__(parent)
        self._pixmap: Optional[QPixmap] = None
        self._drawing = False
        self._start: Optional[QPoint] = None
        self._current_rect: Optional[QRect] = None
        self.setCursor(Qt.CrossCursor)

    def set_pixmap(self, pixmap: QPixmap):
        self._pixmap = pixmap
        self.setFixedSize(pixmap.size())
        self.update()

    # -- painting ----------------------------------------------------------

    def paintEvent(self, event):
        if not self._pixmap:
            return
        painter = QPainter(self)
        painter.drawPixmap(0, 0, self._pixmap)
        if self._current_rect:
            pen = QPen(QColor(220, 40, 40), 2, Qt.DashLine)
            painter.setPen(pen)
            painter.setBrush(QBrush(QColor(220, 40, 40, 35)))
            painter.drawRect(self._current_rect)

    # -- mouse events ------------------------------------------------------

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drawing = True
            self._start = event.pos()
            self._current_rect = QRect(self._start, self._start)

    def mouseMoveEvent(self, event):
        if self._drawing:
            self._current_rect = QRect(self._start, event.pos()).normalized()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self._drawing:
            self._drawing = False
            rect = QRect(self._start, event.pos()).normalized()
            self._current_rect = None
            self.update()
            if rect.width() > 10 and rect.height() > 10:
                self.selection_made.emit(QRectF(rect))

    def wheelEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            self.zoom_requested.emit(event.angleDelta().y())
            event.accept()
        else:
            event.ignore()  # propagate to scroll area


# ---------------------------------------------------------------------------
# Zoomable Scroll Area
# ---------------------------------------------------------------------------

class ZoomScrollArea(QScrollArea):
    """QScrollArea that forwards Ctrl+Wheel to a zoom signal."""
    zoom_requested = pyqtSignal(int)

    def wheelEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            self.zoom_requested.emit(event.angleDelta().y())
            event.accept()
        else:
            super().wheelEvent(event)


# ---------------------------------------------------------------------------
# PDF Viewer  (scroll area + page widget + navigation helpers)
# ---------------------------------------------------------------------------

class PDFViewer(QWidget):
    """Composite widget: scroll area with a PDFPageWidget inside."""

    selection_made = pyqtSignal(int, QRectF)  # (page, rect in pixmap coords)
    page_changed = pyqtSignal(int, int)       # (current 0-based, total)

    RENDER_ZOOM = 2.0   # default render resolution multiplier

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self._scroll = ZoomScrollArea()
        self._scroll.setWidgetResizable(False)
        self._scroll.setAlignment(Qt.AlignCenter)
        self._page_widget = PDFPageWidget()
        self._scroll.setWidget(self._page_widget)
        layout.addWidget(self._scroll)

        # internal state
        self._doc: Optional[fitz.Document] = None
        self._current_page = 0
        self._zoom = self.RENDER_ZOOM

        # signals
        self._page_widget.selection_made.connect(self._on_selection)
        self._page_widget.zoom_requested.connect(self._on_zoom_wheel)
        self._scroll.zoom_requested.connect(self._on_zoom_wheel)

    # -- public API --------------------------------------------------------

    @property
    def render_zoom(self):
        return self._zoom

    @property
    def current_page(self):
        return self._current_page

    @property
    def total_pages(self):
        return len(self._doc) if self._doc else 0

    def set_document(self, doc: fitz.Document, page: int = None):
        self._doc = doc
        if page is not None:
            self._current_page = page
        self._current_page = max(0, min(self._current_page, self.total_pages - 1))
        self._render()

    def next_page(self):
        if self._current_page < self.total_pages - 1:
            self._current_page += 1
            self._render()

    def prev_page(self):
        if self._current_page > 0:
            self._current_page -= 1
            self._render()

    def go_to_page(self, page: int):
        page = max(0, min(page, self.total_pages - 1))
        if page != self._current_page:
            self._current_page = page
            self._render()

    def zoom_in(self):
        self._zoom = min(self._zoom * 1.25, 8.0)
        self._render()

    def zoom_out(self):
        self._zoom = max(self._zoom / 1.25, 0.5)
        self._render()

    def fit_width(self):
        if not self._doc:
            return
        page = self._doc[self._current_page]
        vp_w = self._scroll.viewport().width() - 20  # small margin
        self._zoom = max(vp_w / page.rect.width, 0.5)
        self._render()

    def scroll_to_pdf_point(self, pdf_x, pdf_y):
        px = int(pdf_x * self._zoom)
        py = int(pdf_y * self._zoom)
        self._scroll.ensureVisible(px, py, 100, 100)

    # -- internal ----------------------------------------------------------

    def _render(self):
        if not self._doc:
            return
        page = self._doc[self._current_page]
        mat = fitz.Matrix(self._zoom, self._zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = QImage(pix.samples, pix.width, pix.height,
                     pix.stride, QImage.Format_RGB888).copy()
        self._page_widget.set_pixmap(QPixmap.fromImage(img))
        self.page_changed.emit(self._current_page, self.total_pages)

    def _on_selection(self, rect: QRectF):
        self.selection_made.emit(self._current_page, rect)

    def _on_zoom_wheel(self, delta):
        if delta > 0:
            self.zoom_in()
        else:
            self.zoom_out()


# ---------------------------------------------------------------------------
# Requirement List Item Widget
# ---------------------------------------------------------------------------

class ReqItemWidget(QWidget):
    """Custom widget shown for each requirement in the list."""

    def __init__(self, req: Requirement, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)

        num = QLabel(req.number)
        num.setFont(QFont("Segoe UI", 11, QFont.Bold))
        num.setStyleSheet("color: #c0392b;")
        num.setFixedWidth(55)
        num.setAlignment(Qt.AlignCenter)

        thumb = QLabel()
        scaled = req.screenshot.scaled(
            120, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation
        )
        thumb.setPixmap(scaled)
        thumb.setFixedSize(120, 80)
        thumb.setAlignment(Qt.AlignCenter)
        thumb.setStyleSheet("border: 1px solid #bbb;")

        info = QLabel(f"Page {req.page + 1}")
        info.setFont(QFont("Segoe UI", 9))
        info.setAlignment(Qt.AlignCenter)
        info.setFixedWidth(55)

        layout.addWidget(num)
        layout.addWidget(thumb)
        layout.addWidget(info)


# ---------------------------------------------------------------------------
# Requirements Panel  (right-hand sidebar)
# ---------------------------------------------------------------------------

class RequirementsPanel(QWidget):
    """Sidebar listing captured requirements + numbering controls."""

    delete_requested = pyqtSignal(int)  # list row index

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)

        # -- numbering controls ---
        ctrl_group = QGroupBox("Capture Controls")
        ctrl_layout = QFormLayout(ctrl_group)

        self.next_num_edit = QLineEdit()
        self.next_num_edit.setReadOnly(True)
        self.next_num_edit.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.next_num_edit.setAlignment(Qt.AlignCenter)
        self.next_num_edit.setStyleSheet(
            "background: #fff; color: #c0392b; border: 2px solid #c0392b; "
            "border-radius: 4px; padding: 4px;"
        )
        ctrl_layout.addRow("Next Req #:", self.next_num_edit)

        self.sub_check = QCheckBox("Sub-requirement mode")
        ctrl_layout.addRow(self.sub_check)

        self.sub_parent_label = QLabel("")
        self.sub_parent_label.setFont(QFont("Segoe UI", 9))
        ctrl_layout.addRow(self.sub_parent_label)

        layout.addWidget(ctrl_group)

        # -- requirements list ---
        list_label = QLabel("Captured Requirements")
        list_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        layout.addWidget(list_label)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.SingleSelection)
        self.list_widget.setSpacing(2)
        layout.addWidget(self.list_widget, 1)

        self.delete_btn = QPushButton("Delete Selected")
        self.delete_btn.clicked.connect(self._on_delete)
        layout.addWidget(self.delete_btn)

    def refresh(self, requirements: List[Requirement]):
        self.list_widget.clear()
        for req in requirements:
            item_widget = ReqItemWidget(req)
            item = QListWidgetItem()
            item.setSizeHint(item_widget.sizeHint())
            self.list_widget.addItem(item)
            self.list_widget.setItemWidget(item, item_widget)

    def _on_delete(self):
        row = self.list_widget.currentRow()
        if row >= 0:
            self.delete_requested.emit(row)


# ---------------------------------------------------------------------------
# Main Window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    SCREENSHOT_ZOOM = 3.0  # render zoom for high-res screenshots

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Requirements Tracker")
        self.resize(1400, 900)

        # -- state --
        self._pdf_path: Optional[str] = None
        self._original_bytes: Optional[bytes] = None
        self._doc: Optional[fitz.Document] = None
        self._markup_path: Optional[str] = None
        self._requirements: List[Requirement] = []
        self._next_main = 1
        self._next_sub = 1
        self._last_main = 0
        self._unsaved_changes = False

        # -- widgets --
        self._build_toolbar()
        self._build_central()
        self._build_statusbar()
        self._connect_signals()
        self._update_number_display()

    # ===================== UI construction =================================

    def _build_toolbar(self):
        tb = self.addToolBar("Main")
        tb.setMovable(False)
        tb.setIconSize(QSize(20, 20))
        style = self.style()

        # file actions
        self._act_open = tb.addAction(
            style.standardIcon(QStyle.SP_DialogOpenButton), "Open PDF"
        )
        self._act_save = tb.addAction(
            style.standardIcon(QStyle.SP_DialogSaveButton), "Save Markup"
        )
        self._act_export = tb.addAction(
            style.standardIcon(QStyle.SP_FileDialogDetailedView),
            "Export Requirements Doc"
        )
        tb.addSeparator()

        # navigation
        self._act_prev = tb.addAction(
            style.standardIcon(QStyle.SP_ArrowLeft), "Prev Page"
        )
        self._page_label = QLabel(" Page - / - ")
        self._page_label.setFont(QFont("Segoe UI", 10))
        tb.addWidget(self._page_label)
        self._act_next = tb.addAction(
            style.standardIcon(QStyle.SP_ArrowRight), "Next Page"
        )
        tb.addSeparator()

        # zoom
        self._act_zoom_out = tb.addAction(
            style.standardIcon(QStyle.SP_ArrowDown), "Zoom Out"
        )
        self._zoom_label = QLabel(" 100% ")
        self._zoom_label.setFont(QFont("Segoe UI", 10))
        tb.addWidget(self._zoom_label)
        self._act_zoom_in = tb.addAction(
            style.standardIcon(QStyle.SP_ArrowUp), "Zoom In"
        )
        self._act_fit = tb.addAction("Fit Width")

    def _build_central(self):
        splitter = QSplitter(Qt.Horizontal)

        self._viewer = PDFViewer()
        splitter.addWidget(self._viewer)

        self._panel = RequirementsPanel()
        self._panel.setMinimumWidth(280)
        self._panel.setMaximumWidth(400)
        splitter.addWidget(self._panel)

        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)
        self.setCentralWidget(splitter)

    def _build_statusbar(self):
        self._status = self.statusBar()
        self._status.showMessage("Open a PDF to begin.")

    def _connect_signals(self):
        self._act_open.triggered.connect(self._open_pdf)
        self._act_save.triggered.connect(self._save_markup)
        self._act_export.triggered.connect(self._manual_export)
        self._act_prev.triggered.connect(self._viewer.prev_page)
        self._act_next.triggered.connect(self._viewer.next_page)
        self._act_zoom_in.triggered.connect(self._viewer.zoom_in)
        self._act_zoom_out.triggered.connect(self._viewer.zoom_out)
        self._act_fit.triggered.connect(self._viewer.fit_width)

        self._viewer.selection_made.connect(self._handle_selection)
        self._viewer.page_changed.connect(self._on_page_changed)
        self._panel.delete_requested.connect(self._delete_requirement)
        self._panel.sub_check.stateChanged.connect(
            lambda _: self._update_number_display()
        )
        self._panel.list_widget.currentRowChanged.connect(
            self._on_list_selection_changed
        )

    # ===================== File operations =================================

    def _open_pdf(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open PDF / Drawing / SOW", "",
            "PDF Files (*.pdf);;All Files (*)"
        )
        if not path:
            return

        try:
            with open(path, "rb") as f:
                self._original_bytes = f.read()
            doc = fitz.open(stream=self._original_bytes, filetype="pdf")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open PDF:\n{e}")
            return

        # reset state
        self._pdf_path = path
        self._markup_path = None
        self._doc = doc
        self._unsaved_changes = False
        self._requirements.clear()
        self._next_main = 1
        self._next_sub = 1
        self._last_main = 0
        self._panel.sub_check.setChecked(False)

        self._viewer.set_document(self._doc, page=0)
        self._viewer.fit_width()
        self._panel.refresh(self._requirements)
        self._update_number_display()
        self._status.showMessage(f"Opened: {os.path.basename(path)}")

    # ===================== Selection / capture =============================

    def _handle_selection(self, page_num: int, pixmap_rect: QRectF):
        """Called when the user finishes drawing a rectangle."""
        if not self._doc:
            QMessageBox.information(
                self, "No PDF", "Open a PDF file first."
            )
            return

        zoom = self._viewer.render_zoom

        # convert pixmap coords → PDF points
        pdf_x0 = pixmap_rect.x() / zoom
        pdf_y0 = pixmap_rect.y() / zoom
        pdf_x1 = pixmap_rect.right() / zoom
        pdf_y1 = pixmap_rect.bottom() / zoom
        pdf_rect = (pdf_x0, pdf_y0, pdf_x1, pdf_y1)

        # capture a clean high-res screenshot from the ORIGINAL pdf
        screenshot = self._capture_clean(page_num, pdf_rect)
        if screenshot is None or screenshot.isNull():
            return

        # determine requirement number
        num_str = self._allocate_number()

        req = Requirement(
            number=num_str,
            screenshot=screenshot,
            page=page_num,
            pdf_rect=pdf_rect,
        )
        self._requirements.append(req)

        # rebuild in-memory stamped view (no disk save)
        self._rebuild_view()
        self._panel.refresh(self._requirements)
        self._panel.list_widget.setCurrentRow(len(self._requirements) - 1)
        self._update_number_display()
        self._status.showMessage(
            f"Requirement {num_str} captured  (unsaved)"
        )

    def _capture_clean(self, page_num: int, pdf_rect: tuple) -> Optional[QPixmap]:
        """Render the given rectangle from the *original* PDF at high res."""
        try:
            doc = fitz.open(stream=self._original_bytes, filetype="pdf")
            page = doc[page_num]
            clip = fitz.Rect(pdf_rect)
            mat = fitz.Matrix(self.SCREENSHOT_ZOOM, self.SCREENSHOT_ZOOM)
            pix = page.get_pixmap(matrix=mat, clip=clip, alpha=False)
            img = QImage(
                pix.samples, pix.width, pix.height,
                pix.stride, QImage.Format_RGB888
            ).copy()
            doc.close()
            return QPixmap.fromImage(img)
        except Exception as e:
            QMessageBox.warning(self, "Capture Error", str(e))
            return None

    # ===================== Numbering =======================================

    def _allocate_number(self) -> str:
        if self._panel.sub_check.isChecked() and self._last_main > 0:
            num_str = f"{self._last_main}.{self._next_sub}"
            self._next_sub += 1
        else:
            num_str = str(self._next_main)
            self._last_main = self._next_main
            self._next_main += 1
            self._next_sub = 1
        return num_str

    def _update_number_display(self):
        if self._panel.sub_check.isChecked() and self._last_main > 0:
            nxt = f"{self._last_main}.{self._next_sub}"
            self._panel.sub_parent_label.setText(
                f"Sub-requirements under {self._last_main}"
            )
        else:
            nxt = str(self._next_main)
            self._panel.sub_parent_label.setText("")
        self._panel.next_num_edit.setText(nxt)

    # ===================== PDF stamping / rebuild ==========================

    def _rebuild_view(self):
        """Recreate all stamps on a fresh copy and display (no disk save)."""
        try:
            doc = fitz.open(stream=self._original_bytes, filetype="pdf")
            for req in self._requirements:
                page = doc[req.page]
                r = fitz.Rect(req.pdf_rect)
                self._stamp_page(page, r, req.number)

            cur = self._viewer.current_page
            if self._doc:
                self._doc.close()
            self._doc = doc
            self._viewer.set_document(self._doc, page=cur)
            self._unsaved_changes = True
        except Exception as e:
            QMessageBox.warning(self, "Rebuild Error", str(e))

    def _save_markup(self):
        """Save the marked-up PDF to disk (user-triggered)."""
        if not self._requirements:
            QMessageBox.information(
                self, "Nothing to save",
                "Capture some requirements first."
            )
            return

        # prompt for path if not yet chosen
        if not self._markup_path:
            base = os.path.splitext(self._pdf_path)[0]
            default = base + "_markup.pdf"
            path, _ = QFileDialog.getSaveFileName(
                self, "Save Marked-Up PDF As", default,
                "PDF Files (*.pdf)"
            )
            if not path:
                return
            self._markup_path = path

        try:
            doc = fitz.open(stream=self._original_bytes, filetype="pdf")
            for req in self._requirements:
                page = doc[req.page]
                r = fitz.Rect(req.pdf_rect)
                self._stamp_page(page, r, req.number)
            doc.save(self._markup_path)
            doc.close()
            self._unsaved_changes = False
            self._status.showMessage(
                f"Saved: {os.path.basename(self._markup_path)}"
            )
        except Exception as e:
            QMessageBox.warning(self, "Save Error", str(e))
            return

        # also export requirements doc alongside
        self._auto_export_docx()

    @staticmethod
    def _stamp_page(page, sel_rect: fitz.Rect, number: str):
        """Draw a dashed outline and a numbered stamp on a PDF page."""
        red = (0.85, 0.15, 0.15)
        white = (1, 1, 1)

        # dashed outline around captured area
        page.draw_rect(sel_rect, color=red, width=0.75, dashes="[3 3]")

        # stamp label
        fontsize = 10
        fontname = "helv"
        text_w = fitz.get_text_length(number, fontname=fontname, fontsize=fontsize)
        pad = 3

        # position stamp at upper-left corner of selection
        sx = sel_rect.x0
        sy = sel_rect.y0
        stamp = fitz.Rect(
            sx - 1,
            sy - fontsize - 2 * pad,
            sx + text_w + 2 * pad + 1,
            sy,
        )

        # clamp stamp so it doesn't go above page
        if stamp.y0 < 0:
            shift = -stamp.y0
            stamp.y0 += shift
            stamp.y1 += shift

        page.draw_rect(stamp, color=red, fill=white, width=1.5)
        # text baseline sits near the bottom of the stamp box
        page.insert_text(
            fitz.Point(stamp.x0 + pad, stamp.y1 - pad),
            number,
            fontsize=fontsize,
            fontname=fontname,
            color=red,
        )

    # ===================== Requirements document export ====================

    def _auto_export_docx(self):
        if not HAS_DOCX or not self._markup_path:
            return
        path = os.path.splitext(self._markup_path)[0] + "_requirements.docx"
        self._export_docx(path)

    def _manual_export(self):
        if not self._requirements:
            QMessageBox.information(
                self, "Nothing to export",
                "Capture some requirements first."
            )
            return
        if not HAS_DOCX:
            QMessageBox.warning(
                self, "Missing Dependency",
                "Install python-docx to export:\n  pip install python-docx"
            )
            return
        default = (
            os.path.splitext(self._markup_path)[0] + "_requirements.docx"
            if self._markup_path else "requirements.docx"
        )
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Requirements Document", default,
            "Word Documents (*.docx)"
        )
        if path:
            self._export_docx(path)
            self._status.showMessage(f"Exported: {os.path.basename(path)}")

    def _export_docx(self, path: str):
        try:
            doc = DocxDocument()

            # title
            doc.add_heading("Requirements Tracker", level=0)
            if self._pdf_path:
                doc.add_paragraph(
                    f"Source: {os.path.basename(self._pdf_path)}"
                )
            doc.add_paragraph(
                f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            )
            doc.add_paragraph(
                f"Total requirements: {len(self._requirements)}"
            )
            doc.add_paragraph("")

            # table
            table = doc.add_table(rows=1, cols=4)
            table.style = "Table Grid"
            headers = ["Req #", "Screenshot", "Page", "Notes"]
            for i, h in enumerate(headers):
                cell = table.rows[0].cells[i]
                cell.text = h
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True

            for req in self._requirements:
                row = table.add_row()
                row.cells[0].text = req.number

                # embed screenshot
                img_io = pixmap_to_bytes(req.screenshot)
                paragraph = row.cells[1].paragraphs[0]
                run = paragraph.add_run()
                run.add_picture(img_io, width=Inches(3.5))

                row.cells[2].text = str(req.page + 1)
                row.cells[3].text = ""

            doc.save(path)
        except Exception as e:
            QMessageBox.warning(
                self, "Export Error", f"Failed to export:\n{e}"
            )

    # ===================== Delete ==========================================

    def _delete_requirement(self, row: int):
        if 0 <= row < len(self._requirements):
            removed = self._requirements.pop(row)
            self._rebuild_view()
            self._panel.refresh(self._requirements)
            self._status.showMessage(
                f"Deleted requirement {removed.number}  (unsaved)"
            )

    # ===================== Navigation / UI updates =========================

    def _on_page_changed(self, current: int, total: int):
        self._page_label.setText(f" Page {current + 1} / {total} ")
        zoom_pct = int(self._viewer.render_zoom / PDFViewer.RENDER_ZOOM * 100)
        self._zoom_label.setText(f" {zoom_pct}% ")

    def _on_list_selection_changed(self, row: int):
        if 0 <= row < len(self._requirements):
            req = self._requirements[row]
            self._viewer.go_to_page(req.page)
            self._viewer.scroll_to_pdf_point(req.pdf_rect[0], req.pdf_rect[1])

    # ===================== Keyboard shortcuts ==============================

    def keyPressEvent(self, event):
        key = event.key()
        mod = event.modifiers()

        if mod == Qt.ControlModifier:
            if key == Qt.Key_O:
                self._open_pdf()
                return
            if key == Qt.Key_S:
                self._save_markup()
                return
            if key == Qt.Key_E:
                self._manual_export()
                return
            if key == Qt.Key_Equal or key == Qt.Key_Plus:
                self._viewer.zoom_in()
                return
            if key == Qt.Key_Minus:
                self._viewer.zoom_out()
                return

        if key == Qt.Key_PageDown or key == Qt.Key_Right:
            self._viewer.next_page()
            return
        if key == Qt.Key_PageUp or key == Qt.Key_Left:
            self._viewer.prev_page()
            return
        if key == Qt.Key_Delete:
            row = self._panel.list_widget.currentRow()
            if row >= 0:
                self._delete_requirement(row)
            return
        if key == Qt.Key_F:
            self._viewer.fit_width()
            return

        super().keyPressEvent(event)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # light stylesheet
    app.setStyleSheet("""
        QMainWindow { background: #f5f5f5; }
        QToolBar { background: #e8e8e8; spacing: 6px; padding: 4px; }
        QGroupBox { font-weight: bold; margin-top: 8px; }
        QGroupBox::title { subcontrol-origin: margin; left: 8px; }
        QListWidget { background: #fff; border: 1px solid #ccc; }
        QListWidget::item:selected { background: #dbeafe; }
        QPushButton { padding: 6px 14px; }
        QStatusBar { background: #e8e8e8; }
    """)

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
