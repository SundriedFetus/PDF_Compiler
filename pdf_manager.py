import os
import sys
import json
import logging
import subprocess
# from functools import partial # Removed as it was unused
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QLabel, QListWidget, QListWidgetItem, QFileDialog,
                               QMessageBox, QProgressDialog, QCheckBox, QSplitter, QSpinBox,
                               QDialog, QLineEdit, QDialogButtonBox, QAbstractItemView,
                               QSizePolicy)
from PySide6.QtCore import (Qt, QStandardPaths, Signal, QUrl, QMimeData, QSize, # Added QSize for OrderableListItemWidget
                              QEvent, QPoint, QRect) # Added QEvent, QPoint, QRect for potential future use or if implicitly needed by Qt
from PySide6.QtGui import (QDropEvent, QDragEnterEvent, QPixmap, QImage, QResizeEvent, QAction, # QAction is used for menu/toolbar, keep for now
                             QPainter, QPen, QColor, QFontMetrics, QMouseEvent, QGuiApplication) # Added for custom widgets and event handling

import fitz  # PyMuPDF
from pypdf import PdfWriter, PdfReader, errors as pypdf_errors

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

# Setup basic logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def load_config():
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
    except FileNotFoundError:
        logger.info(f"{CONFIG_FILE} not found. Using default configuration.")
        return {} # Return empty dict to signify using defaults or a fresh setup
    except json.JSONDecodeError:
        # QMessageBox.warning(None, "Config Error", f"Error decoding {CONFIG_FILE}. Using default configuration.")
        # Avoid UI elements in non-UI code; log instead or handle in UI part
        logger.error(f"Error decoding {CONFIG_FILE}. Returning empty config.")
        return {}
    return config

def save_config(config):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4)
    except IOError as e:
        # QMessageBox.critical(None, "Config Error", f"Failed to save configuration to {CONFIG_FILE}: {e}")
        logger.error(f"Failed to save configuration to {CONFIG_FILE}: {e}")


class OrderableListItemWidget(QWidget):
    orderChanged = Signal()

    def __init__(self, filename, full_path, order, parent=None):
        super().__init__(parent)
        self.filename = filename
        self.full_path = full_path
        self._order = order

        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 2, 5, 2) # Reduced margins

        self.order_spinbox = QSpinBox()
        self.order_spinbox.setMinimum(1)
        self.order_spinbox.setMaximum(999)
        self.order_spinbox.setValue(self._order)
        self.order_spinbox.setFixedWidth(50)
        self.order_spinbox.valueChanged.connect(self._emit_order_changed)
        layout.addWidget(self.order_spinbox)

        self.label = QLabel(self.filename)
        self.label.setToolTip(self.full_path)
        layout.addWidget(self.label, 1) # Give stretch to label

    def get_order(self):
        return self.order_spinbox.value()

    def set_order(self, order):
        self.order_spinbox.blockSignals(True)
        self.order_spinbox.setValue(order)
        self._order = order
        self.order_spinbox.blockSignals(False)

    def _emit_order_changed(self):
        self._order = self.order_spinbox.value()
        self.orderChanged.emit()

    def sizeHint(self): # Ensure proper sizing in the list widget
        return QSize(super().sizeHint().width(), self.order_spinbox.sizeHint().height() + 4) # Add a little padding

class SettingsDialog(QDialog):
    def __init__(self, current_inkscape_path, current_gimp_path, 
                 current_delete_originals, current_ordering_keywords,
                 current_libreoffice_draw_path, # New parameter
                 parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setMinimumWidth(600)

        self.new_inkscape_path = current_inkscape_path
        self.new_gimp_path = current_gimp_path
        self.new_delete_originals_on_action = current_delete_originals
        # Ensure new_ordering_keywords is a dictionary, fallback to default if not
        if not isinstance(current_ordering_keywords, dict):
            logger.warning("Received non-dict ordering_keywords for SettingsDialog, falling back to default.")
            self.new_ordering_keywords = DEFAULT_ORDERING_KEYWORDS.copy() # Use a copy
        else:
            self.new_ordering_keywords = current_ordering_keywords.copy() # Use a copy to avoid modifying original
        self.new_libreoffice_draw_path = current_libreoffice_draw_path

        layout = QVBoxLayout(self)

        # Inkscape Path
        inkscape_layout = QHBoxLayout()
        inkscape_layout.addWidget(QLabel("Inkscape Path:"))
        self.inkscape_path_edit = QLineEdit(self.new_inkscape_path)
        inkscape_layout.addWidget(self.inkscape_path_edit)
        btn_browse_inkscape = QPushButton("Browse...")
        btn_browse_inkscape.clicked.connect(self._browse_inkscape_path)
        inkscape_layout.addWidget(btn_browse_inkscape)
        layout.addLayout(inkscape_layout)

        # GIMP Path
        gimp_layout = QHBoxLayout()
        gimp_layout.addWidget(QLabel("GIMP Path:"))
        self.gimp_path_edit = QLineEdit(self.new_gimp_path)
        gimp_layout.addWidget(self.gimp_path_edit)
        btn_browse_gimp = QPushButton("Browse...")
        btn_browse_gimp.clicked.connect(self._browse_gimp_path)
        gimp_layout.addWidget(btn_browse_gimp)
        layout.addLayout(gimp_layout)

        # LibreOffice Draw Path (New Section)
        libreoffice_draw_layout = QHBoxLayout()
        libreoffice_draw_layout.addWidget(QLabel("LibreOffice Draw Path:"))
        self.libreoffice_draw_path_edit = QLineEdit(self.new_libreoffice_draw_path)
        libreoffice_draw_layout.addWidget(self.libreoffice_draw_path_edit)
        btn_browse_libreoffice_draw = QPushButton("Browse...")
        btn_browse_libreoffice_draw.clicked.connect(self._browse_libreoffice_draw_path)
        libreoffice_draw_layout.addWidget(btn_browse_libreoffice_draw)
        layout.addLayout(libreoffice_draw_layout)

        # Delete Originals Checkbox
        self.delete_originals_checkbox = QCheckBox("Delete original files after successful conversion/combination")
        self.delete_originals_checkbox.setChecked(self.new_delete_originals_on_action) # Corrected
        layout.addWidget(self.delete_originals_checkbox)

        # Ordering Keywords
        keywords_layout = QHBoxLayout()
        keywords_layout.addWidget(QLabel("Filename Ordering Keywords (JSON):"))
        self.ordering_keywords_edit = QLineEdit(json.dumps(self.new_ordering_keywords)) # Corrected
        self.ordering_keywords_edit.setReadOnly(True) 
        keywords_layout.addWidget(self.ordering_keywords_edit)
        btn_edit_keywords = QPushButton("Edit Keywords...")
        btn_edit_keywords.clicked.connect(self._edit_keywords)
        keywords_layout.addWidget(btn_edit_keywords)
        layout.addLayout(keywords_layout)

        # Dialog Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def _browse_inkscape_path(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Inkscape Executable", os.path.dirname(self.new_inkscape_path) or os.path.expanduser("~"))
        if path:
            self.inkscape_path_edit.setText(path)

    def _browse_gimp_path(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select GIMP Executable", os.path.dirname(self.gimp_path_edit.text()) or os.path.expanduser("~"))
        if path:
            self.gimp_path_edit.setText(path)

    def _browse_libreoffice_draw_path(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select LibreOffice Draw Executable", os.path.dirname(self.libreoffice_draw_path_edit.text()) or os.path.expanduser("~"))
        if path:
            self.libreoffice_draw_path_edit.setText(path)

    def _edit_keywords(self):
        # Launch a simple editor dialog for now
        editor_dialog = QDialog(self)
        editor_dialog.setWindowTitle("Edit Ordering Keywords")
        editor_dialog.setMinimumWidth(400)

        layout = QVBoxLayout(editor_dialog)

        text_edit = QLineEdit(json.dumps(self.new_ordering_keywords, indent=4))
        text_edit.setText(json.dumps(self.new_ordering_keywords, indent=4))
        layout.addWidget(text_edit)

        btn_ok = QPushButton("OK")
        btn_ok.clicked.connect(lambda: self._on_keywords_edit_accept(text_edit.text(), editor_dialog))
        layout.addWidget(btn_ok)

        editor_dialog.exec()

    def _on_keywords_edit_accept(self, text, dialog):
        try:
            # Attempt to load the JSON from the text edit
            parsed_keywords = json.loads(text)
            if isinstance(parsed_keywords, dict):
                self.new_ordering_keywords = parsed_keywords
                self.ordering_keywords_edit.setText(json.dumps(parsed_keywords))
                dialog.accept()
            else:
                QMessageBox.warning(self, "Invalid Format", "Keywords must be a valid JSON object.")
        except json.JSONDecodeError as e:
            QMessageBox.warning(self, "JSON Decode Error", f"Failed to decode JSON: {e}")

    def accept(self):
        self.new_inkscape_path = self.inkscape_path_edit.text()
        self.new_gimp_path = self.gimp_path_edit.text()
        self.new_libreoffice_draw_path = self.libreoffice_draw_path_edit.text() # Store new path
        self.new_delete_originals_on_action = self.delete_originals_checkbox.isChecked()
        # self.new_ordering_keywords is updated directly by KeywordsEditorDialog
        super().accept()

class PDFToolApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = load_config()
        self.current_factory_paperwork_dir = self.config.get('factory_paperwork_dir', None)
        self.preview_worker = None
        self.preview_cache = {}
        self.last_previewed_path = None
        self._init_ui()
        self._load_all_lists()
        self.setWindowTitle("PDF Management Tool")
        self.setGeometry(100, 100, 1200, 800)
        self.resizeEvent = self._on_resize_event

    def _init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)

        file_management_widget = QWidget()
        file_management_layout = QVBoxLayout(file_management_widget)

        dir_layout = QHBoxLayout()
        self.factory_paperwork_dir_label = QLabel(f"Factory Paperwork (0 items): {self.current_factory_paperwork_dir or 'Not set'}")
        dir_layout.addWidget(self.factory_paperwork_dir_label, 1)
        self.btn_set_factory_paperwork_dir = QPushButton("Set Source Directory")
        self.btn_set_factory_paperwork_dir.clicked.connect(self._set_factory_paperwork_dir)
        dir_layout.addWidget(self.btn_set_factory_paperwork_dir)
        file_management_layout.addLayout(dir_layout)

        source_files_layout = QVBoxLayout()
        source_files_layout.addWidget(QLabel("Available Files (Source)"))
        self.source_files_list = QListWidget()
        self.source_files_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.source_files_list.itemSelectionChanged.connect(self._update_button_states)
        self.source_files_list.itemDoubleClicked.connect(self._add_to_selected_list_handler)
        self.source_files_list.itemClicked.connect(self._handle_source_file_click)
        source_files_layout.addWidget(self.source_files_list)
        file_management_layout.addLayout(source_files_layout, 3)

        source_action_layout = QHBoxLayout()
        self.btn_add_to_selection = QPushButton("Add to Combine List ->")
        self.btn_add_to_selection.clicked.connect(self._add_to_selected_list_handler)
        source_action_layout.addWidget(self.btn_add_to_selection)
        self.btn_open_inkscape = QPushButton("Open with Inkscape")
        self.btn_open_inkscape.clicked.connect(self._open_selected_pdf_with_inkscape)
        source_action_layout.addWidget(self.btn_open_inkscape)
        self.btn_open_gimp = QPushButton("Open with GIMP")
        self.btn_open_gimp.clicked.connect(self._open_selected_png_with_gimp)
        source_action_layout.addWidget(self.btn_open_gimp)
        source_action_layout.addStretch()
        file_management_layout.addLayout(source_action_layout)

        conversion_action_layout = QHBoxLayout()
        self.btn_convert_pdf_to_png = QPushButton("PDF to PNG")
        self.btn_convert_pdf_to_png.clicked.connect(self._convert_selected_pdf_to_png)
        conversion_action_layout.addWidget(self.btn_convert_pdf_to_png)
        self.btn_convert_png_to_pdf = QPushButton("PNG to PDF")
        self.btn_convert_png_to_pdf.clicked.connect(self._convert_selected_png_to_pdf)
        conversion_action_layout.addWidget(self.btn_convert_png_to_pdf)
        conversion_action_layout.addStretch()
        file_management_layout.addLayout(conversion_action_layout)

        selected_files_layout = QVBoxLayout()
        selected_files_layout.addWidget(QLabel("Files Selected for Combination (Ordered)"))
        self.selected_files_list = QListWidget()
        self.selected_files_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.selected_files_list.itemSelectionChanged.connect(self._update_button_states)
        self.selected_files_list.itemClicked.connect(self._handle_selected_file_click)
        selected_files_layout.addWidget(self.selected_files_list)
        
        selected_action_layout = QHBoxLayout()
        self.btn_remove_from_selection = QPushButton("<- Remove from Combine List")
        self.btn_remove_from_selection.clicked.connect(self._remove_from_selected_list_handler)
        selected_action_layout.addWidget(self.btn_remove_from_selection)
        selected_action_layout.addStretch()
        selected_files_layout.addLayout(selected_action_layout)

        self.btn_combine_files = QPushButton("Combine Selected Files into PDF")
        self.btn_combine_files.clicked.connect(self._perform_pdf_combination)
        selected_files_layout.addWidget(self.btn_combine_files)

        # Toggles for opening after combination
        combination_options_layout = QHBoxLayout()
        self.open_in_inkscape_checkbox = QCheckBox("Open combined PDF in Inkscape")
        self.open_in_inkscape_checkbox.setChecked(self.config.get('open_combined_in_inkscape', False))
        combination_options_layout.addWidget(self.open_in_inkscape_checkbox)

        self.open_in_libreoffice_checkbox = QCheckBox("Open combined PDF in LibreOffice Draw")
        self.open_in_libreoffice_checkbox.setChecked(self.config.get('open_combined_in_libreoffice', False))
        combination_options_layout.addWidget(self.open_in_libreoffice_checkbox)
        selected_files_layout.addLayout(combination_options_layout)

        file_management_layout.addLayout(selected_files_layout, 3)
        self.main_splitter.addWidget(file_management_widget)

        self.preview_widget = QWidget()
        preview_layout = QVBoxLayout(self.preview_widget)
        self.toggle_preview_checkbox = QCheckBox("Show File Preview")
        self.toggle_preview_checkbox.setChecked(self.config.get('show_preview', True))
        self.toggle_preview_checkbox.stateChanged.connect(self._toggle_preview_visibility)
        preview_layout.addWidget(self.toggle_preview_checkbox)

        self.preview_label = QLabel("Select a file to preview")
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setMinimumSize(200, 200)
        self.preview_label.setStyleSheet("border: 1px solid gray; background-color: #f0f0f0;")
        self.preview_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored)
        preview_layout.addWidget(self.preview_label, 1)
        self.main_splitter.addWidget(self.preview_widget)
        self.main_splitter.setStretchFactor(0, 2)
        self.main_splitter.setStretchFactor(1, 1)
        self._toggle_preview_visibility()

        main_layout.addWidget(self.main_splitter)

        settings_button_layout = QHBoxLayout()
        settings_button_layout.addStretch()
        self.btn_settings = QPushButton("Settings")
        self.btn_settings.clicked.connect(self._show_settings_dialog)
        settings_button_layout.addWidget(self.btn_settings)

        self.btn_open_blender = QPushButton("Open Image Blender") # New Button
        self.btn_open_blender.clicked.connect(self._open_image_blender) # New Slot
        settings_button_layout.addWidget(self.btn_open_blender) # Add to layout

        main_layout.addLayout(settings_button_layout)

        self.statusBar().showMessage("Ready")
        self.setAcceptDrops(True)

    def _open_image_blender(self):
        # Check if an instance already exists, if so, bring to front
        # For simplicity, we create a new one each time. 
        # A more robust app might manage instances.
        if not hasattr(self, 'image_blender_window') or not self.image_blender_window.isVisible():
            # Dynamically import ImageBlenderWindow to avoid circular dependency if it were in the same file
            # and to ensure it's only imported when needed.
            try:
                from image_blender_gui import ImageBlenderWindow
                self.image_blender_window = ImageBlenderWindow(self) # Pass parent
                self.image_blender_window.show()
            except ImportError as e:
                QMessageBox.critical(self, "Import Error", f"Could not load Image Blender: {e}")
                logger.error("Failed to import ImageBlenderWindow: %s", e)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not open Image Blender: {e}")
                logger.error("Failed to open ImageBlenderWindow: %s", e, exc_info=True)
        else:
            self.image_blender_window.activateWindow() # Bring to front if already open
            self.image_blender_window.raise_() # For some window managers

    def _on_resize_event(self, event: QResizeEvent):
        super().resizeEvent(event)
        if self.preview_widget.isVisible() and self.last_previewed_path:
            self._update_preview(self.last_previewed_path, force_regenerate=True)

    def _handle_source_file_click(self, item):
        if item:
            file_path = item.data(Qt.ItemDataRole.UserRole)
            self._update_preview(file_path)

    def _handle_selected_file_click(self, item):
        if item:
            widget = self.selected_files_list.itemWidget(item)
            if widget:
                file_path = widget.full_path
                self._update_preview(file_path)

    def _load_source_files_list(self):
        self.source_files_list.clear()
        self.preview_cache.clear()
        self.last_previewed_path = None
        self.preview_label.setText("Select a file to preview")
        self.preview_label.setPixmap(QPixmap())
        count = 0
        if self.current_factory_paperwork_dir and os.path.isdir(self.current_factory_paperwork_dir):
            try:
                all_files_in_dir = []
                for entry in os.scandir(self.current_factory_paperwork_dir):
                    if entry.is_file() and entry.name.lower().endswith( (".pdf", ".png") ):
                        all_files_in_dir.append(entry.path)
                
                # Sort files based on keywords in their names, then by name
                ordering_keywords = self.config.get('ordering_keywords', DEFAULT_ORDERING_KEYWORDS)
                # Ensure ordering_keywords is a dictionary; if not, fallback to default.
                if not isinstance(ordering_keywords, dict):
                    logger.warning("Ordering keywords in config is not a dictionary. Falling back to defaults.")
                    ordering_keywords = DEFAULT_ORDERING_KEYWORDS
                
                def get_sort_key(filepath):
                    filename_lower = os.path.basename(filepath).lower()
                    for keyword, order_val in ordering_keywords.items():
                        if keyword in filename_lower:
                            return (order_val, filename_lower)
                    return (DEFAULT_ORDERING_KEYWORDS.get('other', 999), filename_lower) # Default for unmatched

                sorted_files = sorted(all_files_in_dir, key=get_sort_key)

                for file_path in sorted_files:
                    item = QListWidgetItem(os.path.basename(file_path))
                    item.setData(Qt.ItemDataRole.UserRole, file_path) # Store full path
                    item.setToolTip(file_path) # Show full path on hover
                    self.source_files_list.addItem(item)
                    count += 1
            except OSError as e:
                QMessageBox.warning(self, "Directory Error", f"Error reading source directory {self.current_factory_paperwork_dir}: {e}")
                logger.error("Error reading source directory %s: %s", self.current_factory_paperwork_dir, e)
        
        self.factory_paperwork_dir_label.setText(f"Factory Paperwork ({count} items): {self.current_factory_paperwork_dir or 'Not set'}")
        self._update_button_states()

    def _load_all_lists(self):
        self._load_source_files_list()
        self._update_button_states()

    def _set_factory_paperwork_dir(self):
        new_dir = QFileDialog.getExistingDirectory(self, "Select Factory Paperwork Directory", self.current_factory_paperwork_dir or os.path.expanduser("~"))
        if new_dir:
            self.current_factory_paperwork_dir = new_dir
            self.config['factory_paperwork_dir'] = new_dir
            save_config(self.config)
            self.statusBar().showMessage(f"Factory Paperwork directory set to: {new_dir}", 3000)
            self._load_all_lists()
            self.selected_files_list.clear()
        self._update_button_states()

    def _add_to_selected_list_handler(self):
        selected_source_items = self.source_files_list.selectedItems()
        if not selected_source_items:
            return

        current_selected_paths = []
        for i in range(self.selected_files_list.count()):
            item = self.selected_files_list.item(i)
            widget = self.selected_files_list.itemWidget(item)
            if widget:
                current_selected_paths.append(widget.full_path)
        
        initial_next_order = self.selected_files_list.count() + 1
        items_added = False
        for source_item in selected_source_items:
            full_path = source_item.data(Qt.ItemDataRole.UserRole)
            filename = source_item.text()
            if full_path not in current_selected_paths:
                list_item = QListWidgetItem(self.selected_files_list)
                widget = OrderableListItemWidget(filename, full_path, initial_next_order)
                widget.orderChanged.connect(self._selected_item_order_changed)
                list_item.setSizeHint(widget.sizeHint())
                self.selected_files_list.addItem(list_item)
                self.selected_files_list.setItemWidget(list_item, widget)
                current_selected_paths.append(full_path) # Add to list to prevent duplicates in same batch
                initial_next_order +=1
                items_added = True
        
        if items_added:
            self._sort_selected_files_list_by_spinbox()
        self._update_button_states()

    def _remove_from_selected_list_handler(self):
        selected_to_remove = self.selected_files_list.selectedItems()
        if not selected_to_remove:
            return
        for item in selected_to_remove:
            row = self.selected_files_list.row(item)
            self.selected_files_list.takeItem(row)
        self._update_button_states()

    def _selected_item_order_changed(self):
        self._sort_selected_files_list_by_spinbox()
        self._update_button_states()

    def _sort_selected_files_list_by_spinbox(self):
        self.selected_files_list.blockSignals(True)
        recreation_data = []
        for i in range(self.selected_files_list.count()):
            list_item = self.selected_files_list.item(i)
            widget = self.selected_files_list.itemWidget(list_item)
            if widget:
                recreation_data.append({
                    'order': widget.get_order(),
                    'filename': widget.filename,
                    'full_path': widget.full_path
                })
        recreation_data.sort(key=lambda x: (x['order'], x['filename']))
        self.selected_files_list.clear()
        for data in recreation_data:
            new_list_item = QListWidgetItem(self.selected_files_list)
            new_widget = OrderableListItemWidget(data['filename'], data['full_path'], data['order'])
            new_widget.orderChanged.connect(self._selected_item_order_changed)
            new_list_item.setSizeHint(new_widget.sizeHint())
            self.selected_files_list.addItem(new_list_item)
            self.selected_files_list.setItemWidget(new_list_item, new_widget)
        self.selected_files_list.blockSignals(False)

    def _update_button_states(self):
        source_selected = bool(self.source_files_list.selectedItems())
        self.btn_add_to_selection.setEnabled(source_selected)
        can_open_inkscape = False
        can_open_gimp = False
        can_convert_pdf_to_png = False
        can_convert_png_to_pdf = False

        if source_selected and len(self.source_files_list.selectedItems()) == 1:
            selected_item_text = self.source_files_list.selectedItems()[0].text().lower()
            if selected_item_text.endswith(".pdf"):
                if self.config.get('inkscape_path') and os.path.exists(self.config.get('inkscape_path')):
                    can_open_inkscape = True
                can_convert_pdf_to_png = True
            elif selected_item_text.endswith(".png"):
                if self.config.get('gimp_path') and os.path.exists(self.config.get('gimp_path')):
                    can_open_gimp = True
                can_convert_png_to_pdf = True
        
        self.btn_open_inkscape.setEnabled(can_open_inkscape)
        self.btn_open_gimp.setEnabled(can_open_gimp)
        self.btn_convert_pdf_to_png.setEnabled(can_convert_pdf_to_png)
        self.btn_convert_png_to_pdf.setEnabled(can_convert_png_to_pdf)

        selected_for_combination_selected = bool(self.selected_files_list.selectedItems())
        self.btn_remove_from_selection.setEnabled(selected_for_combination_selected)
        can_combine = self.selected_files_list.count() > 0
        self.btn_combine_files.setEnabled(can_combine)

    def _show_settings_dialog(self):
        dialog = SettingsDialog(
            self.config.get('inkscape_path', ''),
            self.config.get('gimp_path', ''),
            self.config.get('delete_originals_on_action', False),
            self.config.get('ordering_keywords', DEFAULT_ORDERING_KEYWORDS),
            self.config.get('libreoffice_draw_path', ''), # New arg
            self
        )
        if dialog.exec():
            self.config['inkscape_path'] = dialog.new_inkscape_path
            self.config['gimp_path'] = dialog.new_gimp_path
            self.config['libreoffice_draw_path'] = dialog.new_libreoffice_draw_path # Save new path
            self.config['delete_originals_on_action'] = dialog.new_delete_originals_on_action
            self.config['ordering_keywords'] = dialog.new_ordering_keywords
            save_config(self.config)
            self.statusBar().showMessage("Settings updated.", 3000)
            self._load_source_files_list() # Reload to reflect potential keyword changes
            self._update_button_states()

    def _open_selected_pdf_with_inkscape(self):
        if not self.config.get('inkscape_path') or not os.path.exists(self.config['inkscape_path']):
            QMessageBox.warning(self, "Inkscape Not Configured", "Inkscape path is not set or invalid in Settings.")
            return
        selected_items = self.source_files_list.selectedItems()
        if selected_items and selected_items[0].text().lower().endswith(".pdf"):
            file_path = selected_items[0].data(Qt.ItemDataRole.UserRole)
            try:
                subprocess.Popen([self.config['inkscape_path'], file_path])
                self.statusBar().showMessage(f"Opening {os.path.basename(file_path)} with Inkscape...", 2000)
            except OSError as e:
                QMessageBox.critical(self, "Error Opening File", f"Could not open {file_path} with Inkscape: {e}")
        else:
            QMessageBox.information(self, "Selection Error", "Please select a single PDF file from the 'Available Files' list.")

    def _open_selected_png_with_gimp(self):
        if not self.config.get('gimp_path') or not os.path.exists(self.config['gimp_path']):
            QMessageBox.warning(self, "GIMP Not Configured", "GIMP path is not set or invalid in Settings.")
            return
        selected_items = self.source_files_list.selectedItems()
        if selected_items and selected_items[0].text().lower().endswith(".png"):
            file_path = selected_items[0].data(Qt.ItemDataRole.UserRole)
            try:
                subprocess.Popen([self.config['gimp_path'], file_path])
                self.statusBar().showMessage(f"Opening {os.path.basename(file_path)} with GIMP...", 2000)
            except OSError as e:
                QMessageBox.critical(self, "Error Opening File", f"Could not open {file_path} with GIMP: {e}")
        else:
            QMessageBox.information(self, "Selection Error", "Please select a single PNG file from the 'Available Files' list.")

    def _convert_selected_pdf_to_png(self):
        selected_items = self.source_files_list.selectedItems()
        if not selected_items or not selected_items[0].text().lower().endswith(".pdf"):
            QMessageBox.warning(self, "Selection Error", "Please select a single PDF file to convert to PNG.")
            return
        
        source_pdf_path = selected_items[0].data(Qt.ItemDataRole.UserRole)
        base, _ = os.path.splitext(source_pdf_path)
        output_png_path = base + ".png"
        
        if os.path.exists(output_png_path):
            reply = QMessageBox.question(self, "File Exists", f"{os.path.basename(output_png_path)} already exists. Overwrite?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No:
                self.statusBar().showMessage("Conversion cancelled.", 2000)
                return

        progress = QProgressDialog("Converting PDF to PNG...", "Cancel", 0, 1, self)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)
        
        try:
            doc = fitz.open(source_pdf_path)
            if not doc.page_count:
                QMessageBox.warning(self, "Empty PDF", "The selected PDF has no pages to convert.")
                doc.close()
                progress.close()
                return

            page = doc.load_page(0)
            pix = page.get_pixmap()
            pix.save(output_png_path)
            doc.close()
            progress.setValue(1)
            QMessageBox.information(self, "Conversion Successful", f"Converted {os.path.basename(source_pdf_path)} to {os.path.basename(output_png_path)}")

            if self.config.get('delete_originals_on_action', False):
                try:
                    os.remove(source_pdf_path)
                    self.statusBar().showMessage(f"Original PDF {os.path.basename(source_pdf_path)} deleted.", 3000)
                except OSError as e_del:
                    QMessageBox.warning(self, "Deletion Error", f"Failed to delete original PDF: {e_del}")
            self._load_source_files_list()
        except fitz.RuntimeError as e_fitz: # More specific fitz error
            progress.setValue(1)
            QMessageBox.critical(self, "Conversion Error", f"Failed to convert PDF to PNG (PyMuPDF error): {e_fitz}")
        except IOError as e_io:
            progress.setValue(1)
            QMessageBox.critical(self, "Conversion Error", f"Failed to save PNG (IOError): {e_io}")
        except Exception as e: # General fallback
            progress.setValue(1)
            logger.error("Unexpected error during PDF to PNG conversion: %s", e, exc_info=True)
            QMessageBox.critical(self, "Conversion Error", f"An unexpected error occurred: {e}")
        finally:
            if progress.isVisible(): progress.close()

    def _convert_selected_png_to_pdf(self):
        selected_items = self.source_files_list.selectedItems()
        if not selected_items or not selected_items[0].text().lower().endswith(".png"):
            QMessageBox.warning(self, "Selection Error", "Please select a single PNG file to convert to PDF.")
            return

        source_png_path = selected_items[0].data(Qt.ItemDataRole.UserRole)
        base, _ = os.path.splitext(source_png_path)
        output_pdf_path = base + ".pdf"

        if os.path.exists(output_pdf_path):
            reply = QMessageBox.question(self, "File Exists", f"{os.path.basename(output_pdf_path)} already exists. Overwrite?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No:
                self.statusBar().showMessage("Conversion cancelled.", 2000)
                return
        
        progress = QProgressDialog("Converting PNG to PDF...", "Cancel", 0, 1, self)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)

        try:
            img_doc = fitz.open(source_png_path)
            pdf_bytes = img_doc.convert_to_pdf()
            img_doc.close()
            with open(output_pdf_path, "wb") as f_out:
                f_out.write(pdf_bytes)
            progress.setValue(1)
            QMessageBox.information(self, "Conversion Successful", f"Converted {os.path.basename(source_png_path)} to {os.path.basename(output_pdf_path)}")

            if self.config.get('delete_originals_on_action', False):
                try:
                    os.remove(source_png_path)
                    self.statusBar().showMessage(f"Original PNG {os.path.basename(source_png_path)} deleted.", 3000)
                except OSError as e_del:
                    QMessageBox.warning(self, "Deletion Error", f"Failed to delete original PNG: {e_del}")
            self._load_source_files_list()
        except fitz.RuntimeError as e_fitz: # More specific fitz error
            progress.setValue(1)
            QMessageBox.critical(self, "Conversion Error", f"Failed to convert PNG to PDF (PyMuPDF error): {e_fitz}")
        except IOError as e_io:
            progress.setValue(1)
            QMessageBox.critical(self, "Conversion Error", f"Failed to write PDF (IOError): {e_io}")
        except Exception as e: # General fallback
            progress.setValue(1)
            logger.error("Unexpected error during PNG to PDF conversion: %s", e, exc_info=True)
            QMessageBox.critical(self, "Conversion Error", f"An unexpected error occurred: {e}")
        finally:
            if progress.isVisible(): progress.close()

    def _perform_pdf_combination(self):
        if self.selected_files_list.count() == 0:
            QMessageBox.warning(self, "No Files Selected", "Please add files to the 'Files Selected for Combination' list.")
            return

        ordered_files_to_combine = []
        for i in range(self.selected_files_list.count()):
            list_item = self.selected_files_list.item(i)
            widget = self.selected_files_list.itemWidget(list_item)
            if widget:
                ordered_files_to_combine.append(widget.full_path)
        
        if not ordered_files_to_combine:
            QMessageBox.warning(self, "Error", "Could not retrieve files from the selection list.")
            return

        output_base_dir_name = "Combined_Output"
        if not self.current_factory_paperwork_dir or not os.path.isdir(self.current_factory_paperwork_dir):
            fallback_dir = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
            QMessageBox.warning(self, "Output Directory Issue", 
                                f"Factory Paperwork directory is not set or invalid. Combined file will be saved in '{fallback_dir}'.")
            output_dir_base = fallback_dir
        else:
            output_dir_base = self.current_factory_paperwork_dir
        output_dir = os.path.join(output_dir_base, output_base_dir_name)

        try:
            os.makedirs(output_dir, exist_ok=True)
        except OSError as e:
            QMessageBox.critical(self, "Directory Error", f"Failed to create output directory {output_dir}: {e}")
            return

        default_output_name = "Combined_Document.pdf"
        output_filename, _ = QFileDialog.getSaveFileName(self, "Save Combined PDF As...",
                                                         os.path.join(output_dir, default_output_name),
                                                         "PDF Files (*.pdf)")
        if not output_filename:
            return

        merger = PdfWriter()
        temp_pdf_files_created = []
        original_files_processed_for_deletion = []

        progress = QProgressDialog("Combining files...", "Cancel", 0, len(ordered_files_to_combine), self)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        
        try:
            for i, file_path in enumerate(ordered_files_to_combine):
                progress.setValue(i)
                if progress.wasCanceled():
                    self.statusBar().showMessage("Combination canceled by user.", 3000)
                    break
                self.statusBar().showMessage(f"Processing {os.path.basename(file_path)}...", 0)
                actual_pdf_to_merge = None
                if not os.path.exists(file_path):
                    QMessageBox.warning(self, "File Not Found", f"File {os.path.basename(file_path)} no longer exists and will be skipped.")
                    continue
                if file_path.lower().endswith(".png"):
                    try:
                        img_doc = fitz.open(file_path)
                        pdf_bytes = img_doc.convert_to_pdf()
                        img_doc.close()
                        temp_pdf_path = os.path.join(output_dir, f"_temp_conversion_{os.path.basename(file_path)}.pdf")
                        with open(temp_pdf_path, "wb") as temp_f:
                            temp_f.write(pdf_bytes)
                        actual_pdf_to_merge = temp_pdf_path
                        temp_pdf_files_created.append(temp_pdf_path)
                        original_files_processed_for_deletion.append(file_path)
                    except fitz.RuntimeError as e_conv_fitz:
                        QMessageBox.warning(self, "PNG Conversion Error", f"Could not convert {os.path.basename(file_path)} to PDF (PyMuPDF error): {e_conv_fitz}")
                        continue
                    except IOError as e_conv_io:
                        QMessageBox.warning(self, "PNG Conversion Error", f"Could not write temporary PDF for {os.path.basename(file_path)} (IOError): {e_conv_io}")
                        continue
                    except Exception as e_conv:
                        logger.error("Unexpected error converting %s to PDF: %s", file_path, e_conv, exc_info=True)
                        QMessageBox.warning(self, "PNG Conversion Error", f"Unexpected error converting {os.path.basename(file_path)}: {e_conv}")
                        continue
                elif file_path.lower().endswith(".pdf"):
                    actual_pdf_to_merge = file_path
                    original_files_processed_for_deletion.append(file_path)
                else:
                    self.statusBar().showMessage(f"Skipping unsupported file: {os.path.basename(file_path)}", 2000)
                    continue
                if actual_pdf_to_merge:
                    try:
                        merger.append(actual_pdf_to_merge)
                    except pypdf_errors.PdfReadError as e_append_pdf:
                        QMessageBox.warning(self, "PDF Append Error", f"Could not append {os.path.basename(file_path)} (pypdf error): {e_append_pdf}")
                        if file_path in original_files_processed_for_deletion:
                            original_files_processed_for_deletion.remove(file_path)
                        continue
                    except Exception as e_append:
                        logger.error("Unexpected error appending %s: %s", file_path, e_append, exc_info=True)
                        QMessageBox.warning(self, "PDF Append Error", f"Unexpected error appending {os.path.basename(file_path)}: {e_append}")
                        if file_path in original_files_processed_for_deletion:
                            original_files_processed_for_deletion.remove(file_path)
                        continue
            progress.setValue(len(ordered_files_to_combine))
            if progress.wasCanceled():
                QMessageBox.information(self, "Combination Canceled", "PDF combination was canceled.")
            elif not merger.pages:
                QMessageBox.information(self, "No Content", "No pages were successfully processed to combine.")
            else:
                merger.write(output_filename)
                QMessageBox.information(self, "Combination Successful", f"Combined PDF saved as {os.path.basename(output_filename)}")
                
                # Save checkbox states to config
                self.config['open_combined_in_inkscape'] = self.open_in_inkscape_checkbox.isChecked()
                self.config['open_combined_in_libreoffice'] = self.open_in_libreoffice_checkbox.isChecked()
                save_config(self.config)

                # Open in Inkscape if checked and path is valid
                if self.open_in_inkscape_checkbox.isChecked():
                    inkscape_path = self.config.get('inkscape_path')
                    if inkscape_path and os.path.exists(inkscape_path):
                        try:
                            subprocess.Popen([inkscape_path, output_filename])
                            self.statusBar().showMessage(f"Opening {os.path.basename(output_filename)} with Inkscape...", 2000)
                        except OSError as e_open:
                            QMessageBox.critical(self, "Error Opening File", f"Could not open {output_filename} with Inkscape: {e_open}")
                    else:
                        QMessageBox.warning(self, "Inkscape Not Configured", "Inkscape path is not set or invalid in Settings. Cannot open combined file.")

                # Open in LibreOffice Draw if checked and path is valid
                if self.open_in_libreoffice_checkbox.isChecked():
                    libreoffice_draw_path = self.config.get('libreoffice_draw_path')
                    if libreoffice_draw_path and os.path.exists(libreoffice_draw_path):
                        try:
                            subprocess.Popen([libreoffice_draw_path, output_filename])
                            self.statusBar().showMessage(f"Opening {os.path.basename(output_filename)} with LibreOffice Draw...", 2000)
                        except OSError as e_open:
                            QMessageBox.critical(self, "Error Opening File", f"Could not open {output_filename} with LibreOffice Draw: {e_open}")
                    else:
                        QMessageBox.warning(self, "LibreOffice Draw Not Configured", "LibreOffice Draw path is not set or invalid in Settings. Cannot open combined file.")

                if self.config.get('delete_originals_on_action', False) and not progress.wasCanceled():
                    deleted_count = 0
                    errors_deleting = []
                    for f_path_orig in original_files_processed_for_deletion:
                        try:
                            if os.path.exists(f_path_orig) and f_path_orig not in temp_pdf_files_created:
                                os.remove(f_path_orig)
                                deleted_count += 1
                        except OSError as e_del_os:
                            errors_deleting.append(f"{os.path.basename(f_path_orig)} ({e_del_os.strerror})")
                    if deleted_count > 0:
                        self.statusBar().showMessage(f"{deleted_count} original file(s) deleted.", 3000)
                    if errors_deleting:
                        QMessageBox.warning(self, "Deletion Error", f"Failed to delete some original files: {', '.join(errors_deleting)}")
                    self._load_source_files_list()
                    self._reconcile_selected_files_list()
        except pypdf_errors.PdfWriteError as e_main_pdf_write:
            QMessageBox.critical(self, "Combination Error", f"Failed to write combined PDF (pypdf error): {e_main_pdf_write}")
            if progress.isVisible(): progress.setValue(len(ordered_files_to_combine))
        except IOError as e_main_io:
            QMessageBox.critical(self, "Combination Error", f"An IO error occurred during PDF combination: {e_main_io}")
            if progress.isVisible(): progress.setValue(len(ordered_files_to_combine))
        except Exception as e_main:
            logger.error("An unexpected error occurred during PDF combination: %s", e_main, exc_info=True)
            QMessageBox.critical(self, "Combination Error", f"An unexpected error occurred: {e_main}")
            if progress.isVisible(): progress.setValue(len(ordered_files_to_combine))
        finally:
            merger.close()
            for temp_file in temp_pdf_files_created:
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                except OSError as e_clean:
                    self.statusBar().showMessage(f"Error cleaning temporary file {os.path.basename(temp_file)}: {e_clean.strerror}", 3000)
            if progress.isVisible(): progress.close()
            self._update_button_states()

    def _reconcile_selected_files_list(self):
        items_to_remove_indices = []
        for i in range(self.selected_files_list.count()):
            list_item = self.selected_files_list.item(i)
            widget = self.selected_files_list.itemWidget(list_item)
            if widget and not os.path.exists(widget.full_path):
                items_to_remove_indices.append(i)
        for index in sorted(items_to_remove_indices, reverse=True):
            self.selected_files_list.takeItem(index)
        if items_to_remove_indices:
            self.statusBar().showMessage("Some selected files were removed as originals no longer exist.", 3000)
            self._sort_selected_files_list_by_spinbox()
        self._update_button_states()

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        if not self.current_factory_paperwork_dir or not os.path.isdir(self.current_factory_paperwork_dir):
            QMessageBox.warning(self, "Directory Not Set", "Cannot drop files: Factory Paperwork directory is not set.")
            event.ignore()
            return

        mime_data = event.mimeData()
        if mime_data.hasUrls():
            event.acceptProposedAction()
            files_to_add_to_source = []
            import shutil # Moved import here
            for url in mime_data.urls():
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                    if file_path.lower().endswith( (".pdf", ".png") ):
                        dest_path = os.path.join(self.current_factory_paperwork_dir, os.path.basename(file_path))
                        if os.path.abspath(file_path) == os.path.abspath(dest_path):
                            self.statusBar().showMessage(f"{os.path.basename(file_path)} is already in the source directory.", 2000)
                            continue
                        if os.path.exists(dest_path):
                            reply = QMessageBox.question(self, "File Exists", 
                                                         f"{os.path.basename(dest_path)} already exists. Overwrite?",
                                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                            if reply == QMessageBox.StandardButton.No:
                                continue
                        try:
                            shutil.copy2(file_path, dest_path)
                            files_to_add_to_source.append(dest_path)
                            self.statusBar().showMessage(f"Copied {os.path.basename(file_path)} to source directory.", 2000)
                        except OSError as e_copy:
                            QMessageBox.warning(self, "Copy Error", f"Could not copy {os.path.basename(file_path)}: {e_copy}")
                        except Exception as e_copy_general:
                            logger.error("Unexpected error copying %s: %s", file_path, e_copy_general, exc_info=True)
                            QMessageBox.warning(self, "Copy Error", f"An unexpected error occurred copying {os.path.basename(file_path)}: {e_copy_general}")
                    else:
                        logger.info("Skipped non-PDF/PNG file: %s", file_path)
            if files_to_add_to_source:
                self._load_source_files_list()
                for added_path in files_to_add_to_source:
                    items = self.source_files_list.findItems(os.path.basename(added_path), Qt.MatchFlag.MatchExactly)
                    for item in items:
                        if item.data(Qt.ItemDataRole.UserRole) == added_path:
                            item.setSelected(True)
                            self.source_files_list.scrollToItem(item)
                            break
        else:
            event.ignore()

    def _toggle_preview_visibility(self):
        checked = self.toggle_preview_checkbox.isChecked()
        self.preview_widget.setVisible(checked)
        self.config['show_preview'] = checked
        save_config(self.config)
        if checked:
            current_source_item = self.source_files_list.currentItem()
            current_selected_item = self.selected_files_list.currentItem()
            if current_selected_item:
                widget = self.selected_files_list.itemWidget(current_selected_item)
                if widget: self._update_preview(widget.full_path)
            elif current_source_item:
                self._update_preview(current_source_item.data(Qt.ItemDataRole.UserRole))
            else:
                self.preview_label.setText("Select a file to preview")
                self.preview_label.setPixmap(QPixmap())
        else:
            self.preview_label.setText("Preview hidden")
            self.preview_label.setPixmap(QPixmap())
            self.last_previewed_path = None

    def _update_preview(self, file_path, force_regenerate=False):
        if not self.preview_widget.isVisible() or not file_path: # Ensure widget is visible and path is valid
            self.preview_label.setPixmap(QPixmap()) # Clear pixmap if not visible or no path
            self.preview_label.setText("Preview hidden or no file selected" if not file_path else "Preview hidden")
            return

        logger.debug(f"_update_preview called for: {file_path}, force_regenerate: {force_regenerate}")
        current_label_size_for_debug = self.preview_label.size()
        logger.debug(f"Preview widget visible: {self.preview_widget.isVisible()}, Preview label WxH: {current_label_size_for_debug.width()}x{current_label_size_for_debug.height()}")

        # Cache key includes file path and a quantized version of the target label size for more accurate caching
        current_label_size = self.preview_label.size()
        quantize_factor = 20 # Group sizes by 20px blocks to reduce cache churn for minor size fluctuations
        # Ensure a minimum size for cache key parts to avoid issues if label is transiently tiny during layout
        cache_label_width = max(current_label_size.width(), quantize_factor)
        cache_label_height = max(current_label_size.height(), quantize_factor)
        # Quantize by rounding down to the nearest multiple of quantize_factor
        label_size_tuple_for_cache = ( (cache_label_width // quantize_factor) * quantize_factor,
                                       (cache_label_height // quantize_factor) * quantize_factor )
        cache_key = (file_path, label_size_tuple_for_cache)
        logger.debug(f"Using cache key: {cache_key}")

        try:
            if not force_regenerate and cache_key in self.preview_cache:
                scaled_pixmap = self.preview_cache[cache_key]
                logger.debug(f"Using cached preview for {cache_key}. Cached pixmap valid: {not scaled_pixmap.isNull()}, size: {scaled_pixmap.size().width()}x{scaled_pixmap.size().height()}")
            else:
                logger.debug(f"Generating new preview for {file_path} (target cache key {cache_key})")
                pixmap = None # Initialize pixmap variable
                if file_path.lower().endswith(".pdf"):
                    doc = fitz.open(file_path)
                    if not doc.page_count:
                        self.preview_label.setText(f"Empty PDF: {os.path.basename(file_path)}")
                        self.preview_label.setPixmap(QPixmap())
                        doc.close()
                        return
                    page = doc.load_page(0) # Preview first page
                    zoom = 2 # Increase DPI for better quality
                    mat = fitz.Matrix(zoom, zoom)
                    fitz_pix = page.get_pixmap(matrix=mat) # Renamed to fitz_pix to avoid confusion
                    image_format = QImage.Format.Format_RGB888
                    if fitz_pix.alpha:
                         image_format = QImage.Format.Format_RGBA8888
                    
                    qimage = QImage(fitz_pix.samples, fitz_pix.width, fitz_pix.height, fitz_pix.stride, image_format)
                    pixmap = QPixmap.fromImage(qimage) # Assign to general pixmap variable
                    doc.close()
                elif file_path.lower().endswith(".png"):
                    pixmap = QPixmap(file_path) # Assign to general pixmap variable
                else:
                    self.preview_label.setText(f"Unsupported file type: {os.path.basename(file_path)}")
                    self.preview_label.setPixmap(QPixmap())
                    return

                if pixmap is None or pixmap.isNull(): # Check the general pixmap variable
                    logger.warning(f"Pixmap is null after loading attempt for {file_path}")
                    self.preview_label.setText(f"Could not load preview for {os.path.basename(file_path)}")
                    self.preview_label.setPixmap(QPixmap())
                    if cache_key in self.preview_cache: # Remove potentially bad cache entry
                        self.preview_cache.pop(cache_key, None)
                    return
                
                logger.debug(f"Pixmap loaded from file: valid: {not pixmap.isNull()}, size: {pixmap.size().width()}x{pixmap.size().height()}")

                # Scale pixmap to fit the label while maintaining aspect ratio
                label_size_for_scaling = self.preview_label.size() # Use current actual size for scaling
                # Subtract a small margin to avoid scrollbars if the label has a border/padding
                available_width = max(1, label_size_for_scaling.width() - 4) 
                available_height = max(1, label_size_for_scaling.height() - 4)
                logger.debug(f"Scaling original pixmap (size {pixmap.size().width()}x{pixmap.size().height()}) to fit available space: {available_width}x{available_height}")

                scaled_pixmap = pixmap.scaled(available_width, available_height,
                                              Qt.AspectRatioMode.KeepAspectRatio,
                                              Qt.TransformationMode.SmoothTransformation)
                self.preview_cache[cache_key] = scaled_pixmap
                logger.debug(f"Cached new preview. Scaled pixmap valid: {not scaled_pixmap.isNull()}, size: {scaled_pixmap.size().width()}x{scaled_pixmap.size().height()}")

            self.preview_label.setPixmap(scaled_pixmap)
            self.preview_label.setText("") # Clear any previous text like "Select a file"
            self.last_previewed_path = file_path

        except fitz.RuntimeError as e_fitz:
            logger.error("PyMuPDF error generating preview for %s: %s", file_path, e_fitz, exc_info=False)
            self.preview_label.setText(f"PyMuPDF error for {os.path.basename(file_path)}")
            if file_path in self.preview_cache: del self.preview_cache[file_path]
            self.last_previewed_path = None
        except Exception as e:
            logger.error("Error generating preview for %s: %s", file_path, e, exc_info=True)
            self.preview_label.setText(f"Error previewing {os.path.basename(file_path)}")
            if file_path in self.preview_cache: del self.preview_cache[file_path]
            self.last_previewed_path = None


def main():
    app = QApplication(sys.argv)
    # Apply a style if desired, e.g., Fusion for a modern look
    # app.setStyle("Fusion") 
    window = PDFToolApp()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()