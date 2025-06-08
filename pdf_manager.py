print("Script top-level: Starting imports...") # DEBUG
import io
import json
import logging
import os
import subprocess # For opening files with external apps
import sys
import tempfile
import time # For preview regeneration delay
from functools import lru_cache, partial

import fitz # PyMuPDF
from PIL import Image, ImageQt, ExifTags, ImageOps, UnidentifiedImageError
from PySide6.QtCore import (
    QBuffer, QByteArray, QDir, QEvent, QIODevice,
    QItemSelectionModel, QMargins, QMetaObject, QPoint,
    QRect, QSettings, QSize, Qt, QThread, QTimer, Signal,
    Slot, QStandardPaths
)
from PySide6.QtGui import (
    QAction, QKeySequence, QPixmap, QImage, QPainter, QIcon, QColor, QPalette, QTransform
)
from PySide6.QtCore import Qt, QSize, QTimer, QEvent, QPoint, QMimeData, QRect
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QListWidget, QListWidgetItem, QAbstractItemView, QSplitter,
    QCheckBox, QFileDialog, QDialog, QLineEdit, QSpinBox, QDialogButtonBox,
    QMessageBox, QProgressDialog, QSizePolicy, QStyle, QStyledItemDelegate,
    QSpacerItem, QGroupBox, QInputDialog
)
from pypdf import PdfReader, PdfWriter, Transformation
from pypdf import errors as pypdf_errors # Import errors submodule
# This assumes image_blender_gui.py is in the same directory or Python path
try:
    from image_blender_gui import ImageBlenderWindow
except ImportError:
    ImageBlenderWindow = None
    print("Warning: image_blender_gui.py not found. Image Blender functionality will be disabled.")


print("Script top-level: Imports finished.") # DEBUG

# Configuration
CONFIG_FILE = 'config.json'
DEFAULT_ORDERING_KEYWORDS = {
    "drawing": 10,
    "plan": 20,
    "detail": 30,
    "schedule": 40,
    "specification": 50,
    "report": 60,
    "photo": 70,
    "other": 100
}

# Define A4 dimensions in points (1 point = 1/72 inch)
A4_WIDTH_PT = 595.0 # Use float for precision
A4_HEIGHT_PT = 842.0 # Use float for precision

# Setup basic logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def load_config():
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
    except FileNotFoundError:
        logger.info(f"{CONFIG_FILE} not found. Using default configuration.")
        return {}
    except json.JSONDecodeError:
        logger.error(f"Error decoding {CONFIG_FILE}. Returning empty config.")
        return {}
    return config

def save_config(config):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4)
    except IOError as e:
        logger.error(f"Failed to save configuration to {CONFIG_FILE}: {e}")

class PreviewWorker(QThread):
    """Worker thread for generating file previews."""
    previewReady = Signal(QPixmap, str)
    errorOccurred = Signal(str)

    def __init__(self, file_path, preview_size, dpi, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.preview_size = preview_size
        self.dpi = dpi
        self._is_running = True

    def run(self):
        if not self._is_running:
            return

        try:
            pixmap = None
            if self.file_path.lower().endswith('.pdf'):
                doc = fitz.open(self.file_path)
                if doc.page_count > 0:
                    page = doc.load_page(0)
                    pix = page.get_pixmap(dpi=self.dpi)
                    image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
                    pixmap = QPixmap.fromImage(image)
                doc.close()
            elif self.file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
                pixmap = QPixmap(self.file_path)

            if pixmap and not pixmap.isNull():
                scaled_pixmap = pixmap.scaled(self.preview_size, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                if self._is_running:
                    self.previewReady.emit(scaled_pixmap, self.file_path)
            elif self._is_running:
                self.errorOccurred.emit(f"Unsupported or invalid file: {os.path.basename(self.file_path)}")
        except Exception as e:
            logger.error(f"Error generating preview for {self.file_path}: {e}")
            if self._is_running:
                self.errorOccurred.emit(f"Error: {e}")

    def stop(self):
        self._is_running = False


class OrderableListItemWidget(QWidget):
    orderAttempted = Signal()

    def __init__(self, filename, full_path, initial_order, parent=None):
        super().__init__(parent)
        self.full_path = full_path
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 2, 5, 2)
        
        self.order_spinbox = QSpinBox()
        self.order_spinbox.setMinimum(1)
        self.order_spinbox.setMaximum(9999) 
        self.order_spinbox.setValue(initial_order)
        self.order_spinbox.setFixedWidth(60)
        self.order_spinbox.valueChanged.connect(self.orderAttempted.emit)
        layout.addWidget(self.order_spinbox)
        
        self.filename_label = QLabel(filename)
        self.filename_label.setToolTip(full_path)
        layout.addWidget(self.filename_label, 1)
        
        self.setLayout(layout)

    def get_order(self):
        return self.order_spinbox.value()
        
    def set_order(self, order, silent=True):
        if silent:
            self.order_spinbox.blockSignals(True)
        self.order_spinbox.setValue(order)
        if silent:
            self.order_spinbox.blockSignals(False)

class SettingsDialog(QDialog):
    def __init__(self, current_inkscape_path, current_gimp_path,
                 current_ordering_keywords,
                 current_libreoffice_draw_path,
                 current_pdf_to_png_dpi,
                 parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setMinimumWidth(600)

        self.new_inkscape_path = current_inkscape_path
        self.new_gimp_path = current_gimp_path
        self.new_ordering_keywords = current_ordering_keywords.copy()
        self.new_libreoffice_draw_path = current_libreoffice_draw_path
        self.new_pdf_to_png_dpi = current_pdf_to_png_dpi

        layout = QVBoxLayout(self)
        
        self.inkscape_path_edit = QLineEdit(self.new_inkscape_path)
        self.gimp_path_edit = QLineEdit(self.new_gimp_path)
        self.libreoffice_draw_path_edit = QLineEdit(self.new_libreoffice_draw_path)

        layout.addLayout(self._create_path_editor("Inkscape Path:", self.inkscape_path_edit, self._browse_inkscape_path))
        layout.addLayout(self._create_path_editor("GIMP Path:", self.gimp_path_edit, self._browse_gimp_path))
        layout.addLayout(self._create_path_editor("LibreOffice Draw Path:", self.libreoffice_draw_path_edit, self._browse_libreoffice_draw_path))

        dpi_layout = QHBoxLayout()
        dpi_layout.addWidget(QLabel("PDF to PNG Conversion DPI:"))
        self.dpi_spinbox = QSpinBox()
        self.dpi_spinbox.setMinimum(72)
        self.dpi_spinbox.setMaximum(600)
        self.dpi_spinbox.setValue(self.new_pdf_to_png_dpi)
        dpi_layout.addWidget(self.dpi_spinbox)
        dpi_layout.addStretch()
        layout.addLayout(dpi_layout)

        keywords_layout = QHBoxLayout()
        keywords_layout.addWidget(QLabel("Filename Ordering Keywords (JSON):"))
        self.ordering_keywords_edit = QLineEdit(json.dumps(self.new_ordering_keywords))
        self.ordering_keywords_edit.setReadOnly(True) 
        keywords_layout.addWidget(self.ordering_keywords_edit)
        btn_edit_keywords = QPushButton("Edit...")
        btn_edit_keywords.clicked.connect(self._edit_keywords)
        keywords_layout.addWidget(btn_edit_keywords)
        layout.addLayout(keywords_layout)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def _create_path_editor(self, label, line_edit, browse_slot):
        layout = QHBoxLayout()
        layout.addWidget(QLabel(label))
        layout.addWidget(line_edit)
        btn_browse = QPushButton("Browse...")
        btn_browse.clicked.connect(browse_slot)
        layout.addWidget(btn_browse)
        return layout
        
    def _browse_inkscape_path(self):
        self._browse_path(self.inkscape_path_edit, "Select Inkscape Executable")

    def _browse_gimp_path(self):
        self._browse_path(self.gimp_path_edit, "Select GIMP Executable")

    def _browse_libreoffice_draw_path(self):
        self._browse_path(self.libreoffice_draw_path_edit, "Select LibreOffice Executable")
        
    def _browse_path(self, line_edit, title):
        path, _ = QFileDialog.getOpenFileName(self, title, os.path.dirname(line_edit.text()) or os.path.expanduser("~"))
        if path:
            line_edit.setText(path)

    def _edit_keywords(self):
        text, ok = QInputDialog.getMultiLineText(self, "Edit Ordering Keywords", "Enter keywords as JSON:", json.dumps(self.new_ordering_keywords, indent=4))
        if ok and text:
            try:
                parsed = json.loads(text)
                if isinstance(parsed, dict):
                    self.new_ordering_keywords = parsed
                    self.ordering_keywords_edit.setText(json.dumps(self.new_ordering_keywords))
                else:
                    QMessageBox.warning(self, "Invalid Format", "Keywords must be a valid JSON object.")
            except json.JSONDecodeError as e:
                QMessageBox.warning(self, "JSON Error", f"Could not parse JSON: {e}")

    def accept(self):
        self.new_inkscape_path = self.inkscape_path_edit.text()
        self.new_gimp_path = self.gimp_path_edit.text()
        self.new_libreoffice_draw_path = self.libreoffice_draw_path_edit.text()
        self.new_pdf_to_png_dpi = self.dpi_spinbox.value()
        super().accept()

class PDFToolApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = load_config()
        self.config.setdefault('factory_paperwork_dir', None)
        self.config.setdefault('delete_after_conversion', False)
        self.config.setdefault('delete_after_combination', False)
        self.config.setdefault('open_in_libreoffice', False)
        self.config.setdefault('pdf_to_png_dpi', 150)
        self.config.setdefault('show_preview', True)
        self.config.setdefault('ordering_keywords', DEFAULT_ORDERING_KEYWORDS)
        
        self.preview_worker = None
        self.image_blender_window = None
        self._init_ui()
        self._load_all_lists()
        self.setWindowTitle("PDF Management Tool")
        self.setGeometry(100, 100, 1400, 900)
        self.resizeEvent = self._on_resize_event

    def _init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        # Main vertical splitter
        main_splitter = QSplitter(Qt.Orientation.Vertical)

        # Top panel with source/combine lists and preview
        top_panel = QWidget()
        top_layout = QHBoxLayout(top_panel)
        top_splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left Panel (Source and Combine lists)
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        
        dir_layout = QHBoxLayout()
        self.factory_paperwork_dir_label = QLabel(f"Source: {self.config['factory_paperwork_dir'] or 'Not set'}")
        dir_layout.addWidget(self.factory_paperwork_dir_label, 1)
        self.btn_set_factory_paperwork_dir = QPushButton("Set Source Directory")
        self.btn_set_factory_paperwork_dir.clicked.connect(self._set_factory_paperwork_dir)
        dir_layout.addWidget(self.btn_set_factory_paperwork_dir)
        left_layout.addLayout(dir_layout)

        # Source List Group
        source_group = QGroupBox("Available Files")
        source_group_layout = QVBoxLayout(source_group)
        self.source_files_list = QListWidget()
        self.source_files_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.source_files_list.itemSelectionChanged.connect(self._update_button_states)
        self.source_files_list.itemDoubleClicked.connect(self._add_to_selected_list_handler)
        self.source_files_list.itemClicked.connect(self._on_source_item_clicked)
        source_group_layout.addWidget(self.source_files_list)
        left_layout.addWidget(source_group)

        # Add to Selection Button
        self.btn_add_to_selection = QPushButton("Add to Combine List ->")
        self.btn_add_to_selection.clicked.connect(self._add_to_selected_list_handler)
        left_layout.addWidget(self.btn_add_to_selection)

        # Combine List Group
        combine_group = QGroupBox("Files to Combine (Drag to Reorder)")
        combine_group_layout = QVBoxLayout(combine_group)
        self.selected_files_list = QListWidget()
        self.selected_files_list.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.selected_files_list.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.selected_files_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.selected_files_list.itemSelectionChanged.connect(self._update_button_states)
        self.selected_files_list.itemClicked.connect(self._on_selected_item_clicked)
        self.selected_files_list.model().rowsMoved.connect(self._handle_rows_moved)
        combine_group_layout.addWidget(self.selected_files_list)
        left_layout.addWidget(combine_group)

        # Remove from Selection Button
        self.btn_remove_from_selection = QPushButton("<- Remove from Combine List")
        self.btn_remove_from_selection.clicked.connect(self._remove_from_selected_list_handler)
        left_layout.addWidget(self.btn_remove_from_selection)
        
        top_splitter.addWidget(left_panel)
        
        # Right Panel (Preview)
        self.preview_widget = QWidget()
        preview_layout = QVBoxLayout(self.preview_widget)
        self.preview_label = QLabel("Select a file to preview")
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setMinimumSize(200, 200)
        self.preview_label.setStyleSheet("border: 1px solid gray; background-color: #f0f0f0;")
        self.preview_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored)
        preview_layout.addWidget(self.preview_label, 1)
        top_splitter.addWidget(self.preview_widget)
        top_splitter.setSizes([400, 600]) # Initial ratio

        top_layout.addWidget(top_splitter)
        main_splitter.addWidget(top_panel)
        
        # Bottom Panel (Actions)
        bottom_panel = QWidget()
        bottom_layout = QHBoxLayout(bottom_panel)

        # Actions Group
        actions_group = QGroupBox("Actions")
        actions_layout = QHBoxLayout(actions_group)
        self.btn_open_inkscape = QPushButton("Open with Inkscape")
        self.btn_open_inkscape.clicked.connect(self._open_selected_pdf_with_inkscape)
        actions_layout.addWidget(self.btn_open_inkscape)
        self.btn_open_gimp = QPushButton("Open with GIMP")
        self.btn_open_gimp.clicked.connect(self._open_selected_png_with_gimp)
        actions_layout.addWidget(self.btn_open_gimp)
        self.btn_convert_pdf_to_png = QPushButton("PDF to PNG")
        self.btn_convert_pdf_to_png.clicked.connect(self._convert_selected_pdf_to_png)
        actions_layout.addWidget(self.btn_convert_pdf_to_png)
        self.btn_convert_png_to_pdf = QPushButton("PNG to PDF")
        self.btn_convert_png_to_pdf.clicked.connect(self._convert_selected_png_to_pdf)
        actions_layout.addWidget(self.btn_convert_png_to_pdf)
        self.delete_after_conversion_checkbox = QCheckBox("Delete after conversion")
        self.delete_after_conversion_checkbox.setChecked(self.config['delete_after_conversion'])
        self.delete_after_conversion_checkbox.stateChanged.connect(self._save_toggle_config)
        actions_layout.addWidget(self.delete_after_conversion_checkbox)
        bottom_layout.addWidget(actions_group)

        # Combination Group
        combination_group_box = QGroupBox("Combination")
        combination_layout = QHBoxLayout(combination_group_box)
        self.btn_combine_files = QPushButton("Combine Files into PDF")
        self.btn_combine_files.clicked.connect(self._perform_pdf_combination)
        combination_layout.addWidget(self.btn_combine_files)
        self.delete_after_combination_checkbox = QCheckBox("Delete after combination")
        self.delete_after_combination_checkbox.setChecked(self.config['delete_after_combination'])
        self.delete_after_combination_checkbox.stateChanged.connect(self._save_toggle_config)
        combination_layout.addWidget(self.delete_after_combination_checkbox)
        self.open_in_libreoffice_checkbox = QCheckBox("Open in LibreDraw")
        self.open_in_libreoffice_checkbox.setChecked(self.config.get('open_in_libreoffice', False))
        self.open_in_libreoffice_checkbox.stateChanged.connect(self._save_toggle_config)
        combination_layout.addWidget(self.open_in_libreoffice_checkbox)
        bottom_layout.addWidget(combination_group_box)
        
        main_splitter.addWidget(bottom_panel)
        main_splitter.setSizes([800, 200]) # Main layout ratio

        main_layout.addWidget(main_splitter)

        # Bottom buttons
        settings_layout = QHBoxLayout()
        settings_layout.addStretch()
        self.btn_settings = QPushButton("Settings")
        self.btn_settings.clicked.connect(self._show_settings_dialog)
        settings_layout.addWidget(self.btn_settings)
        if ImageBlenderWindow:
            self.btn_open_blender = QPushButton("Open Image Blender")
            self.btn_open_blender.clicked.connect(self._open_image_blender)
            settings_layout.addWidget(self.btn_open_blender)
        main_layout.addLayout(settings_layout)

        self.statusBar().showMessage("Ready")
        self.setAcceptDrops(True)
        self._update_button_states()

    def _on_source_item_clicked(self, item):
        self.selected_files_list.clearSelection()
        self._update_preview(item.data(Qt.ItemDataRole.UserRole))

    def _on_selected_item_clicked(self, item):
        self.source_files_list.clearSelection()
        widget = self.selected_files_list.itemWidget(item)
        if widget:
            self._update_preview(widget.full_path)

    def _show_settings_dialog(self):
        dialog = SettingsDialog(
            current_inkscape_path=self.config.get('inkscape_path', ''),
            current_gimp_path=self.config.get('gimp_path', ''),
            current_ordering_keywords=self.config.get('ordering_keywords', DEFAULT_ORDERING_KEYWORDS),
            current_libreoffice_draw_path=self.config.get('libreoffice_draw_path', ''),
            current_pdf_to_png_dpi=self.config.get('pdf_to_png_dpi', 150),
            parent=self
        )
        if dialog.exec():
            self.config['inkscape_path'] = dialog.new_inkscape_path
            self.config['gimp_path'] = dialog.new_gimp_path
            self.config['libreoffice_draw_path'] = dialog.new_libreoffice_draw_path
            self.config['ordering_keywords'] = dialog.new_ordering_keywords
            self.config['pdf_to_png_dpi'] = dialog.new_pdf_to_png_dpi
            save_config(self.config)
            self.statusBar().showMessage("Settings saved.", 3000)
            self._resort_combine_list()

    def _set_factory_paperwork_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Source Directory", self.config['factory_paperwork_dir'] or QDir.homePath())
        if directory:
            self.config['factory_paperwork_dir'] = directory
            save_config(self.config)
            self._load_all_lists()

    def _load_all_lists(self):
        self._load_source_list()
        self._reconcile_selected_files_list()
        self._update_button_states()

    def _load_source_list(self):
        self.source_files_list.clear()
        path = self.config.get('factory_paperwork_dir')
        self.factory_paperwork_dir_label.setText(f"Source: {path or 'Not set'}")
        if path and os.path.isdir(path):
            try:
                files = [f for f in os.listdir(path) if f.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg'))]
                files.sort()
                for filename in files:
                    full_path = os.path.join(path, filename)
                    item = QListWidgetItem(filename)
                    item.setData(Qt.ItemDataRole.UserRole, full_path)
                    self.source_files_list.addItem(item)
                self.factory_paperwork_dir_label.setText(f"Source ({len(files)} items): {path}")
            except Exception as e:
                logger.error(f"Failed to load source list: {e}")
                self.statusBar().showMessage(f"Error loading files: {e}", 5000)

    def _reconcile_selected_files_list(self):
        items_to_remove = []
        for i in range(self.selected_files_list.count()):
            widget = self.selected_files_list.itemWidget(self.selected_files_list.item(i))
            if not (widget and os.path.exists(widget.full_path)):
                items_to_remove.append(i)
        for i in reversed(items_to_remove):
            self.selected_files_list.takeItem(i)

    def _update_button_states(self):
        has_source_selection = len(self.source_files_list.selectedItems()) > 0
        has_selected_selection = len(self.selected_files_list.selectedItems()) > 0
        has_items_to_combine = self.selected_files_list.count() > 0

        self.btn_add_to_selection.setEnabled(has_source_selection)
        self.btn_remove_from_selection.setEnabled(has_selected_selection)
        self.btn_combine_files.setEnabled(has_items_to_combine)
        self.btn_convert_pdf_to_png.setEnabled(has_source_selection)
        self.btn_convert_png_to_pdf.setEnabled(has_source_selection)
        self.btn_open_inkscape.setEnabled(has_source_selection)
        self.btn_open_gimp.setEnabled(has_source_selection)

    def _add_to_selected_list_handler(self):
        selected_items = self.source_files_list.selectedItems()
        if not selected_items: return

        current_paths = {self.selected_files_list.itemWidget(self.selected_files_list.item(i)).full_path for i in range(self.selected_files_list.count())}
        max_order = 0
        for i in range(self.selected_files_list.count()):
            widget = self.selected_files_list.itemWidget(self.selected_files_list.item(i))
            if widget and widget.get_order() > max_order:
                max_order = widget.get_order()

        for item in selected_items:
            full_path = item.data(Qt.ItemDataRole.UserRole)
            if full_path in current_paths: continue
            
            filename = os.path.basename(full_path)
            order = self._get_order_for_filename(filename)
            if order == DEFAULT_ORDERING_KEYWORDS.get("other", 100):
                max_order += 10
                order = max_order

            new_item = QListWidgetItem()
            widget = OrderableListItemWidget(filename, full_path, order)
            widget.orderAttempted.connect(self._resort_combine_list)
            new_item.setSizeHint(widget.sizeHint())
            self.selected_files_list.addItem(new_item)
            self.selected_files_list.setItemWidget(new_item, widget)

        self._resort_combine_list()
        self._update_button_states()

    def _resort_combine_list(self):
        items_data = []
        for i in range(self.selected_files_list.count()):
            item = self.selected_files_list.item(i)
            widget = self.selected_files_list.itemWidget(item)
            items_data.append({
                'filename': widget.filename_label.text(),
                'full_path': widget.full_path,
                'order': widget.get_order()
            })
        
        items_data.sort(key=lambda x: x['order'])
        
        self.selected_files_list.clear()

        for data in items_data:
            new_item = QListWidgetItem()
            widget = OrderableListItemWidget(data['filename'], data['full_path'], data['order'])
            widget.orderAttempted.connect(self._resort_combine_list)
            new_item.setSizeHint(widget.sizeHint())
            self.selected_files_list.addItem(new_item)
            self.selected_files_list.setItemWidget(new_item, widget)


    def _get_order_for_filename(self, filename):
        filename_lower = filename.lower()
        keywords = self.config.get('ordering_keywords', DEFAULT_ORDERING_KEYWORDS)
        for keyword, order in keywords.items():
            if keyword in filename_lower:
                return order
        return keywords.get("other", 100)
    
    def _remove_from_selected_list_handler(self):
        for item in self.selected_files_list.selectedItems():
            row = self.selected_files_list.row(item)
            self.selected_files_list.takeItem(row)
        self._update_button_states()

    def _handle_rows_moved(self, parent, start, end, destination, row):
        # After a drag-drop, re-sequence the order numbers based on visual order
        for i in range(self.selected_files_list.count()):
            widget = self.selected_files_list.itemWidget(self.selected_files_list.item(i))
            if widget:
                widget.set_order((i + 1) * 10)
    
    def _update_preview(self, file_path):
        if not file_path:
            self.preview_label.setText("Select a file to preview")
            self.preview_label.setPixmap(QPixmap())
            return

        if self.preview_worker and self.preview_worker.isRunning():
            self.preview_worker.stop()
            self.preview_worker.wait()

        self.preview_label.setText("Generating preview...")
        self.preview_worker = PreviewWorker(file_path, self.preview_label.size(), self.config['pdf_to_png_dpi'])
        self.preview_worker.previewReady.connect(self._on_preview_generated)
        self.preview_worker.errorOccurred.connect(self.statusBar().showMessage)
        self.preview_worker.start()

    def _on_preview_generated(self, pixmap, file_path):
        self.preview_label.setPixmap(pixmap)

    def _save_toggle_config(self):
        self.config['delete_after_conversion'] = self.delete_after_conversion_checkbox.isChecked()
        self.config['delete_after_combination'] = self.delete_after_combination_checkbox.isChecked()
        self.config['open_in_libreoffice'] = self.open_in_libreoffice_checkbox.isChecked()
        save_config(self.config)

    def _on_resize_event(self, event):
        super().resizeEvent(event)

    def _perform_pdf_combination(self):
        if self.selected_files_list.count() == 0:
            return

        ordered_files = [self.selected_files_list.itemWidget(self.selected_files_list.item(i)).full_path for i in range(self.selected_files_list.count())]
        
        default_dir = os.path.dirname(ordered_files[0]) if ordered_files else self.config.get('factory_paperwork_dir')
        output_path, _ = QFileDialog.getSaveFileName(self, "Save Combined PDF", os.path.join(default_dir or '', "combined.pdf"), "PDF Files (*.pdf)")

        if not output_path: return
            
        writer = PdfWriter()
        progress = QProgressDialog("Combining files...", "Cancel", 0, len(ordered_files), self)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.show()

        try:
            for i, f_path in enumerate(ordered_files):
                progress.setValue(i)
                if progress.wasCanceled(): break
                
                reader = None
                pdf_stream = None
                try:
                    if f_path.lower().endswith(('.png', '.jpg', '.jpeg')):
                        with Image.open(f_path) as img:
                            if img.mode == 'RGBA':
                                bg = Image.new('RGB', img.size, (255, 255, 255))
                                bg.paste(img, mask=img.split()[3])
                                img = bg
                            
                            pdf_stream = io.BytesIO()
                            img.save(pdf_stream, "PDF", resolution=100.0)
                            pdf_stream.seek(0)
                            reader = PdfReader(pdf_stream)

                    elif f_path.lower().endswith('.pdf'):
                        reader = PdfReader(f_path)
                    
                    if reader:
                        for page in reader.pages:
                            if page.mediabox.width > page.mediabox.height:
                                page.rotate(90)

                            target_page = writer.add_blank_page(width=A4_WIDTH_PT, height=A4_HEIGHT_PT)
                            scale = min(A4_WIDTH_PT / page.mediabox.width, A4_HEIGHT_PT / page.mediabox.height)
                            tx = (A4_WIDTH_PT - page.mediabox.width * scale) / 2
                            ty = (A4_HEIGHT_PT - page.mediabox.height * scale) / 2
                            transform = Transformation().scale(scale).translate(tx, ty)
                            target_page.merge_transformed_page(page, transform)
                finally:
                    if reader and hasattr(reader, 'stream') and reader.stream:
                        reader.stream.close() 
                    if pdf_stream:
                        pdf_stream.close()
            
            if not progress.wasCanceled():
                with open(output_path, "wb") as fp:
                    writer.write(fp)
                QMessageBox.information(self, "Success", "Files combined successfully.")
                if self.config.get('delete_after_combination'):
                    for f in ordered_files: os.remove(f)
                    self._load_all_lists()
                
                if self.config.get('open_in_libreoffice', False):
                    libreoffice_path = self.config.get('libreoffice_draw_path')
                    if libreoffice_path and os.path.exists(libreoffice_path):
                        subprocess.Popen([libreoffice_path, output_path])
                    else:
                        QMessageBox.warning(self, "LibreOffice Not Found", "Path to LibreOffice Draw is not set or invalid in settings.")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred during PDF combination: {e}")
            logger.error(f"Error in PDF combination: {e}", exc_info=True)
        finally:
            writer.close()
            progress.close()

    def _open_selected_pdf_with_inkscape(self):
        selected_items = self.source_files_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Selection Error", "No PDF file selected.")
            return

        file_path = selected_items[0].data(Qt.ItemDataRole.UserRole)
        if not file_path.lower().endswith('.pdf'):
            QMessageBox.warning(self, "File Type Error", "Selected file is not a PDF.")
            return

        inkscape_path = self.config.get('inkscape_path')
        if not inkscape_path or not os.path.exists(inkscape_path):
            QMessageBox.critical(self, "Configuration Error", "Inkscape path is not configured or invalid. Please set it in Settings.")
            return
        
        try:
            subprocess.Popen([inkscape_path, file_path])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open with Inkscape: {e}")

    def _open_selected_png_with_gimp(self):
        selected_items = self.source_files_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Selection Error", "No image file selected.")
            return

        file_path = selected_items[0].data(Qt.ItemDataRole.UserRole)
        if not file_path.lower().endswith(('.png', '.jpg', '.jpeg')):
            QMessageBox.warning(self, "File Type Error", "Selected file is not a supported image for GIMP.")
            return

        gimp_path = self.config.get('gimp_path')
        if not gimp_path or not os.path.exists(gimp_path):
            QMessageBox.critical(self, "Configuration Error", "GIMP path is not configured or invalid. Please set it in Settings.")
            return
        
        try:
            subprocess.Popen([gimp_path, file_path])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open with GIMP: {e}")

    def _convert_selected_pdf_to_png(self):
        selected_items = self.source_files_list.selectedItems()
        if not selected_items: return
        
        pdf_path = selected_items[0].data(Qt.ItemDataRole.UserRole)
        if not pdf_path.lower().endswith('.pdf'):
            QMessageBox.warning(self, "File Type Error", "Please select a PDF file.")
            return

        output_path = os.path.splitext(pdf_path)[0] + ".png"
        if os.path.exists(output_path):
            if QMessageBox.question(self, "File Exists", "Output PNG already exists. Overwrite?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.No:
                return
        
        try:
            doc = fitz.open(pdf_path)
            page = doc.load_page(0)
            pix = page.get_pixmap(dpi=self.config.get('pdf_to_png_dpi', 150))
            pix.save(output_path)
            QMessageBox.information(self, "Success", "PDF converted to PNG.")
            if self.config.get('delete_after_conversion'):
                os.remove(pdf_path)
            self._load_source_list()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to convert PDF: {e}")

    def _convert_selected_png_to_pdf(self):
        selected_items = self.source_files_list.selectedItems()
        if not selected_items: return
        
        image_paths = [item.data(Qt.ItemDataRole.UserRole) for item in selected_items if item.data(Qt.ItemDataRole.UserRole).lower().endswith(('.png', '.jpg', '.jpeg'))]
        if not image_paths:
            QMessageBox.warning(self, "Selection Error", "No valid image files selected.")
            return
            
        output_path = os.path.splitext(image_paths[0])[0] + ".pdf"
        if os.path.exists(output_path):
            if QMessageBox.question(self, "File Exists", "Output PDF already exists. Overwrite?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.No:
                return

        writer = PdfWriter()
        try:
            for img_path in image_paths:
                with Image.open(img_path) as img:
                    if img.mode == 'RGBA': img = img.convert('RGB')
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        img.save(tmp.name, "PDF")
                        reader = PdfReader(tmp.name)
                        writer.add_page(reader.pages[0])
                        reader.stream.close()
                    os.remove(tmp.name)
            
            with open(output_path, "wb") as f_out:
                writer.write(f_out)
            QMessageBox.information(self, "Success", f"{len(image_paths)} image(s) converted to PDF.")
            if self.config.get('delete_after_conversion'):
                for path in image_paths: os.remove(path)
            self._load_source_list()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to convert images: {e}")
        finally:
            writer.close()

    def _open_image_blender(self):
        if ImageBlenderWindow:
            if not self.image_blender_window:
                self.image_blender_window = ImageBlenderWindow()
            self.image_blender_window.show()
            
    def closeEvent(self, event):
        if self.image_blender_window: self.image_blender_window.close()
        save_config(self.config)
        super().closeEvent(event)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dropEvent(self, event):
        urls = [url.toLocalFile() for url in event.mimeData().urls()]
        self._add_files_from_paths(urls)

# MAIN EXECUTION BLOCK
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFToolApp()
    window.show()
    sys.exit(app.exec())