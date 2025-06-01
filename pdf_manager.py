import sys
import os

# Add lib directory to Python path
script_dir = os.path.dirname(os.path.abspath(__file__))
lib_dir = os.path.join(script_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

import shutil
import json
import re
import subprocess # For Inkscape & GIMP
import io # For in-memory PDF conversion of PNGs

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox,
    QDialog, QScrollArea, QSpinBox, QDialogButtonBox, QCheckBox,
    QMenuBar, QGroupBox, QStackedWidget, QSizePolicy
)
from PySide6.QtGui import QPixmap, QImage, QAction
from PySide6.QtCore import Qt, QSize
from pypdf import PdfWriter, PdfReader

# Attempt to import PyMuPDF (fitz)
try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False
    print("PyMuPDF (fitz) not found. PDF previews and PNG-to-PDF conversion will be disabled. Please install with 'pip install PyMuPDF'")

CONFIG_FILE = 'pdf_tool_config.json'
DEFAULT_ORDERING_KEYWORDS = {
    "front": 1, "install": 2, "cad": 3, "model": 4, "checklist": 100
}
PREVIEW_WIDTH = 100
PREVIEW_HEIGHT = 140


# --- Configuration Handling ---
def load_config():
    defaults = {
        'source_pdf_dir': '', 'factory_paperwork_dir': '',
        'ordering_keywords': DEFAULT_ORDERING_KEYWORDS.copy(),
        'inkscape_path': '', 'gimp_path': ''
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                if 'project_dir' in config and 'factory_paperwork_dir' not in config:
                    config['factory_paperwork_dir'] = config.pop('project_dir')
                for key, value in defaults.items(): config.setdefault(key, value)
                return config
        except Exception as e: print(f"Error loading {CONFIG_FILE}: {e}. Using defaults.")
    return defaults

def save_config(config):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump(config, f, indent=4)
    except Exception as e: QMessageBox.warning(None, "Config Error", f"Could not save config: {e}")

# --- File Preview Generation ---
def generate_file_preview(file_full_path):
    filename_lower = file_full_path.lower()
    if filename_lower.endswith(".pdf"):
        if not PYMUPDF_AVAILABLE: return None
        try:
            doc = fitz.open(file_full_path)
            if not doc.page_count > 0: doc.close(); return None
            page = doc.load_page(0); zoom = 1.5; mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            fmt = QImage.Format.Format_RGB888 if pix.alpha == 0 else QImage.Format.Format_ARGB32
            qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
            qpix = QPixmap.fromImage(qimg); doc.close()
            return qpix.scaled(QSize(PREVIEW_WIDTH, PREVIEW_HEIGHT), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        except Exception as e: print(f"Error PDF preview {os.path.basename(file_full_path)}: {e}"); return None
    elif filename_lower.endswith(".png"):
        try:
            qpix = QPixmap(file_full_path)
            if qpix.isNull(): return None
            return qpix.scaled(QSize(PREVIEW_WIDTH, PREVIEW_HEIGHT), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        except Exception as e: print(f"Error PNG preview {os.path.basename(file_full_path)}: {e}"); return None
    return None

# --- Helper for Sorting ---
def get_file_sort_priority(filename_short, ordering_keywords_config):
    fn_lower = filename_short.lower()
    num_match = re.search(r'\d+', filename_short)
    num_val = int(num_match.group(0)) if num_match else float('inf')
    num_presence = 0 if num_match else 1 # Sort items with numbers first
    kw_val = float('inf') # Default for items without keywords
    for kw, prio in ordering_keywords_config.items():
        if kw.lower() in fn_lower: kw_val = min(kw_val, prio)
    return (num_presence, num_val, kw_val, filename_short) # Sort key

# --- Settings Dialog for Executable Paths ---
class SettingsDialog(QDialog):
    def __init__(self, current_inkscape_path, current_gimp_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configure Executable Paths"); self.new_inkscape_path = current_inkscape_path; self.new_gimp_path = current_gimp_path
        layout = QVBoxLayout(self)
        for app_name, current_path_attr, edit_attr, browse_method_name, placeholder in [
            ("Inkscape", "new_inkscape_path", "inkscape_path_edit", "_browse_inkscape_path", "e.g., .../inkscape.exe"),
            ("GIMP", "new_gimp_path", "gimp_path_edit", "_browse_gimp_path", "e.g., .../gimp-2.10.exe")]:
            group_box = QGroupBox(app_name); app_layout = QVBoxLayout(); path_layout = QHBoxLayout()
            path_layout.addWidget(QLabel("Executable Path:"))
            path_edit = QLineEdit(getattr(self, current_path_attr)); path_edit.setPlaceholderText(placeholder)
            setattr(self, edit_attr, path_edit)
            path_layout.addWidget(path_edit, 1)
            browse_btn = QPushButton("Browse..."); browse_btn.clicked.connect(getattr(self, browse_method_name))
            path_layout.addWidget(browse_btn); app_layout.addLayout(path_layout); group_box.setLayout(app_layout); layout.addWidget(group_box)
        
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(self._accept_settings); btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box); self.setMinimumWidth(550)

    def _browse_executable_path(self, line_edit_widget, title):
        filt = "Executable files (*.exe)" if sys.platform == "win32" else "All files (*)"
        curr_dir = os.path.dirname(line_edit_widget.text()) if line_edit_widget.text() else os.path.expanduser("~")
        fp, _ = QFileDialog.getOpenFileName(self, title, curr_dir, filt)
        if fp: line_edit_widget.setText(fp)
    def _browse_inkscape_path(self): self._browse_executable_path(self.inkscape_path_edit, "Select Inkscape Executable")
    def _browse_gimp_path(self): self._browse_executable_path(self.gimp_path_edit, "Select GIMP Executable")
    def _accept_settings(self):
        self.new_inkscape_path = self.inkscape_path_edit.text().strip()
        self.new_gimp_path = self.gimp_path_edit.text().strip()
        if self.new_inkscape_path and not os.path.isfile(self.new_inkscape_path):
            QMessageBox.warning(self, "Invalid Inkscape Path", "Path is not a valid file."); return
        if self.new_gimp_path and not os.path.isfile(self.new_gimp_path):
            QMessageBox.warning(self, "Invalid GIMP Path", "Path is not a valid file."); return
        self.accept()

# --- Main Application Window ---
class PDFToolApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("File Management Tool"); self.setGeometry(100, 100, 900, 700)
        self.config = load_config(); self.last_created_pdf_path = None
        self.main_gui_file_items = [] 
        self._create_menu()
        central_widget = QWidget(self); self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget); main_layout.setSpacing(10); central_widget.setContentsMargins(10,10,10,10)
        
        dir_setup_group = QGroupBox("Directory Setup"); dir_setup_layout = QVBoxLayout()
        source_layout = QHBoxLayout(); source_layout.addWidget(QLabel("Originals Folder (PDFs, PNGs):"))
        self.source_dir_edit = QLineEdit(self.config.get('source_pdf_dir', ''))
        self.source_dir_edit.editingFinished.connect(self._update_config_and_refresh_display) # Changed
        source_layout.addWidget(self.source_dir_edit, 1); btn_src = QPushButton("Browse..."); btn_src.clicked.connect(self._browse_source_dir)
        source_layout.addWidget(btn_src); dir_setup_layout.addLayout(source_layout)
        factory_layout = QHBoxLayout(); factory_layout.addWidget(QLabel("Factory Paperwork Folder:"))
        self.factory_dir_edit = QLineEdit(self.config.get('factory_paperwork_dir', ''))
        self.factory_dir_edit.editingFinished.connect(self._update_config_and_refresh_display) # Changed
        factory_layout.addWidget(self.factory_dir_edit, 1); btn_fact = QPushButton("Browse..."); btn_fact.clicked.connect(self._browse_factory_dir)
        factory_layout.addWidget(btn_fact); dir_setup_layout.addLayout(factory_layout)
        dir_setup_group.setLayout(dir_setup_layout); main_layout.addWidget(dir_setup_group)

        contents_group = QGroupBox("Factory Paperwork Folder Contents"); contents_layout = QVBoxLayout()
        self.view_mode_stack = QStackedWidget()
        self._create_list_view_page(); self._create_preview_view_page()
        contents_layout.addWidget(self.view_mode_stack)
        btn_toggle_view = QPushButton("Toggle View Mode"); btn_toggle_view.clicked.connect(self._toggle_view_mode)
        contents_layout.addWidget(btn_toggle_view, 0, Qt.AlignmentFlag.AlignRight)
        contents_group.setLayout(contents_layout)
        main_layout.addWidget(contents_group)
        contents_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        actions_group = QGroupBox("Actions"); actions_main_layout = QVBoxLayout() 
        copy_layout = QHBoxLayout(); copy_layout.addStretch()
        self.copy_button = QPushButton("Copy Files to Factory Folder"); self.copy_button.clicked.connect(self._copy_files)
        copy_layout.addWidget(self.copy_button); copy_layout.addStretch(); actions_main_layout.addLayout(copy_layout)
        compile_actions_group = QGroupBox("Compilation Actions"); compile_actions_layout = QHBoxLayout()
        self.combine_all_button = QPushButton("Combine ALL (Main View Order)"); self.combine_all_button.clicked.connect(self._combine_all_files_from_main_gui)
        self.selective_compile_button = QPushButton("Combine SELECTED (Main View)"); self.selective_compile_button.clicked.connect(self._selective_compile_files_from_main_gui)
        compile_actions_layout.addWidget(self.combine_all_button); compile_actions_layout.addWidget(self.selective_compile_button)
        compile_actions_group.setLayout(compile_actions_layout); actions_main_layout.addWidget(compile_actions_group)
        edit_actions_group = QGroupBox("External Editing Actions"); edit_actions_layout = QHBoxLayout()
        self.open_last_pdf_inkscape_button = QPushButton("Open Last Output PDF in Inkscape"); self.open_last_pdf_inkscape_button.clicked.connect(self._open_last_pdf_in_inkscape); self.open_last_pdf_inkscape_button.setEnabled(False)
        self.open_selected_inkscape_button = QPushButton("Open SELECTED in Inkscape"); self.open_selected_inkscape_button.clicked.connect(self._open_selected_in_inkscape); self.open_selected_inkscape_button.setEnabled(False)
        self.open_gimp_button = QPushButton("Open SELECTED in GIMP"); self.open_gimp_button.clicked.connect(self._open_selected_in_gimp); self.open_gimp_button.setEnabled(False)
        edit_actions_layout.addWidget(self.open_last_pdf_inkscape_button); edit_actions_layout.addWidget(self.open_selected_inkscape_button); edit_actions_layout.addWidget(self.open_gimp_button)
        
        # Add the new button for Image Blender
        self.open_image_blender_button = QPushButton("Open Image Blender")
        self.open_image_blender_button.clicked.connect(self._open_image_blender)
        edit_actions_layout.addWidget(self.open_image_blender_button) # Add to the same layout as other edit actions

        edit_actions_group.setLayout(edit_actions_layout); actions_main_layout.addWidget(edit_actions_group)
        actions_group.setLayout(actions_main_layout); main_layout.addWidget(actions_group)
        
        status_msg = "Ready."
        if not PYMUPDF_AVAILABLE: status_msg += " PyMuPDF not found: PDF previews & PNG inclusion disabled."
        self.statusBar().showMessage(status_msg)
        self._update_button_states(); self._refresh_factory_folder_display()

    def _create_list_view_page(self):
        self.list_view_page = QWidget()
        list_scroll_area = QScrollArea(); list_scroll_area.setWidgetResizable(True)
        self.list_view_widget_content = QWidget() 
        self.list_view_layout = QVBoxLayout(self.list_view_widget_content)
        self.list_view_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        list_scroll_area.setWidget(self.list_view_widget_content)
        page_layout = QVBoxLayout(self.list_view_page); page_layout.addWidget(list_scroll_area)
        self.view_mode_stack.addWidget(self.list_view_page)

    def _create_preview_view_page(self):
        self.preview_view_page = QWidget()
        preview_scroll_area = QScrollArea(); preview_scroll_area.setWidgetResizable(True)
        self.preview_view_widget_content = QWidget()
        self.preview_view_layout = QVBoxLayout(self.preview_view_widget_content) 
        self.preview_view_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        preview_scroll_area.setWidget(self.preview_view_widget_content)
        page_layout = QVBoxLayout(self.preview_view_page); page_layout.addWidget(preview_scroll_area)
        self.view_mode_stack.addWidget(self.preview_view_page)

    def _clear_layout(self, layout):
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0); widget = item.widget()
                if widget is not None: widget.setParent(None); widget.deleteLater() # Ensure proper cleanup
                else:
                    sub_layout = item.layout()
                    if sub_layout is not None: self._clear_layout(sub_layout)
    
    def _render_main_gui_file_display(self):
        """Sorts main_gui_file_items and re-populates both view layouts."""
        self._clear_layout(self.list_view_layout)
        self._clear_layout(self.preview_view_layout)

        if not self.main_gui_file_items: # Handle case where list might be empty after operations
            no_files_msg = "No PDF or PNG files found in folder."
            self.list_view_layout.addWidget(QLabel(no_files_msg))
            self.preview_view_layout.addWidget(QLabel(no_files_msg))
            self._update_button_states()
            return

        # Sort the items based on their 'order' attribute
        self.main_gui_file_items.sort(key=lambda item: item['order'])

        # Re-populate both views with the sorted items
        for item_data in self.main_gui_file_items:
            # List View Item
            list_item_layout = QHBoxLayout()
            list_item_layout.addWidget(item_data['list_view_widgets']['checkbox'])
            list_fn_label = QLabel(item_data['filename']); list_fn_label.setWordWrap(True)
            list_item_layout.addWidget(list_fn_label, 1)
            list_item_layout.addWidget(item_data['list_view_widgets']['spinbox'])
            self.list_view_layout.addLayout(list_item_layout)

            # Preview View Item
            preview_item_layout = QHBoxLayout()
            preview_item_layout.addWidget(item_data['preview_view_widgets']['checkbox'])
            
            preview_label_widget = QLabel() # Create new label for preview image each render
            preview_label_widget.setFixedSize(PREVIEW_WIDTH, PREVIEW_HEIGHT + 10)
            preview_label_widget.setStyleSheet("border: 1px solid #ccc; background-color: #f0f0f0; padding: 5px;")
            preview_label_widget.setAlignment(Qt.AlignmentFlag.AlignCenter)
            qpixmap = generate_file_preview(item_data['full_path'])
            if qpixmap: preview_label_widget.setPixmap(qpixmap)
            else: preview_label_widget.setText("Preview N/A")
            
            preview_fn_label = QLabel(f"{item_data['filename']}"); preview_fn_label.setWordWrap(True)
            details_layout = QVBoxLayout(); details_layout.addWidget(preview_fn_label)
            details_layout.addWidget(item_data['preview_view_widgets']['spinbox']) # Add existing spinbox
            
            preview_item_layout.addWidget(preview_label_widget)
            preview_item_layout.addLayout(details_layout,1)
            self.preview_view_layout.addLayout(preview_item_layout)
        self._update_button_states()


    def _refresh_factory_folder_display(self):
        # Clear existing items and their widgets before fetching new ones
        for item_data in self.main_gui_file_items:
            item_data['list_view_widgets']['checkbox'].setParent(None)
            item_data['list_view_widgets']['spinbox'].setParent(None)
            item_data['preview_view_widgets']['checkbox'].setParent(None)
            item_data['preview_view_widgets']['spinbox'].setParent(None)
        self.main_gui_file_items.clear()
        
        # Note: _clear_layout is called by _render_main_gui_file_display, so not needed here if we call render.

        factory_dir = self.config.get('factory_paperwork_dir', '')
        if not factory_dir or not os.path.isdir(factory_dir):
            self._render_main_gui_file_display() # Render with empty items to show message
            return
        try: files = self._get_processable_files_from_factory_folder(show_msg_if_empty=False)
        except Exception:
            self._render_main_gui_file_display() # Render with empty items to show message
            return
        if not files:
            self._render_main_gui_file_display() # Render with empty items to show message
            return

        # Initial sort based on filename properties
        sorted_files_by_name = sorted(files, key=lambda fn: get_file_sort_priority(fn, self.config['ordering_keywords']))
        num_files = len(sorted_files_by_name); max_order_val = max(1, num_files)

        for i, filename_short in enumerate(sorted_files_by_name):
            full_path = os.path.join(factory_dir, filename_short)
            
            # Create new widgets for this item
            cb_list = QCheckBox(); sb_list = QSpinBox(); sb_list.setMinimum(1); sb_list.setMaximum(max_order_val); sb_list.setValue(i + 1)
            cb_preview = QCheckBox(); sb_preview = QSpinBox(); sb_preview.setMinimum(1); sb_preview.setMaximum(max_order_val); sb_preview.setValue(i + 1)

            item_data = {
                'filename': filename_short, 'full_path': full_path, 
                'order': i + 1, # Initial order matches visual sort by name
                'is_checked': False,
                'list_view_widgets': {'checkbox': cb_list, 'spinbox': sb_list},
                'preview_view_widgets': {'checkbox': cb_preview, 'spinbox': sb_preview}
            }
            self.main_gui_file_items.append(item_data)
            
            item_index = len(self.main_gui_file_items) - 1
            cb_list.toggled.connect(lambda state, idx=item_index, src_view='list': self._handle_checkbox_toggled(idx, state, src_view))
            sb_list.valueChanged.connect(lambda value, idx=item_index, src_view='list': self._handle_spinbox_value_changed(idx, value, src_view))
            cb_preview.toggled.connect(lambda state, idx=item_index, src_view='preview': self._handle_checkbox_toggled(idx, state, src_view))
            sb_preview.valueChanged.connect(lambda value, idx=item_index, src_view='preview': self._handle_spinbox_value_changed(idx, value, src_view))
        
        self._render_main_gui_file_display() # This will sort by 'order' and then populate UI

    def _handle_checkbox_toggled(self, item_index, state, source_view_type):
        if 0 <= item_index < len(self.main_gui_file_items):
            item_data = self.main_gui_file_items[item_index]
            item_data['is_checked'] = state
            target_cb = item_data['preview_view_widgets']['checkbox'] if source_view_type == 'list' else item_data['list_view_widgets']['checkbox']
            blocked = target_cb.blockSignals(True); target_cb.setChecked(state); target_cb.blockSignals(blocked)
            self._update_button_states()

    def _handle_spinbox_value_changed(self, item_index, new_value, source_view_type):
        if not (0 <= item_index < len(self.main_gui_file_items)): return
        self.main_gui_file_items[item_index]['order'] = new_value
        item_data = self.main_gui_file_items[item_index]
        target_sb = item_data['preview_view_widgets']['spinbox'] if source_view_type == 'list' else item_data['list_view_widgets']['spinbox']
        blocked = target_sb.blockSignals(True); target_sb.setValue(new_value); target_sb.blockSignals(blocked)

        for i, other_item_data in enumerate(self.main_gui_file_items):
            if i == item_index: continue
            if other_item_data['order'] == new_value:
                current_orders = {item['order'] for idx, item in enumerate(self.main_gui_file_items) if idx != i}
                placeholder_val = 1; num_items = len(self.main_gui_file_items)
                while placeholder_val <= num_items and placeholder_val in current_orders : placeholder_val += 1
                if placeholder_val > num_items:
                    temp_val = 1; all_orders = {item['order'] for item in self.main_gui_file_items}
                    while temp_val in all_orders: temp_val +=1
                    placeholder_val = temp_val
                
                other_item_data['order'] = placeholder_val
                sb_list_disp = other_item_data['list_view_widgets']['spinbox']
                sb_prev_disp = other_item_data['preview_view_widgets']['spinbox']
                
                bl_list = sb_list_disp.blockSignals(True); sb_list_disp.setValue(placeholder_val); sb_list_disp.blockSignals(bl_list)
                bl_prev = sb_prev_disp.blockSignals(True); sb_prev_disp.setValue(placeholder_val); sb_prev_disp.blockSignals(bl_prev)
                break 
        self._render_main_gui_file_display() # Re-sort and re-render the UI

    def _toggle_view_mode(self):
        self.view_mode_stack.setCurrentIndex((self.view_mode_stack.currentIndex() + 1) % self.view_mode_stack.count())

    def _create_menu(self):
        menubar = self.menuBar(); settings_menu = menubar.addMenu("&Settings")
        cfg_paths_action = QAction("&Configure Executable Paths...", self); cfg_paths_action.triggered.connect(self._show_settings_dialog)
        settings_menu.addAction(cfg_paths_action)

    def _show_settings_dialog(self):
        dialog = SettingsDialog(self.config.get('inkscape_path', ''), self.config.get('gimp_path', ''), self)
        if dialog.exec():
            self.config['inkscape_path'] = dialog.new_inkscape_path; self.config['gimp_path'] = dialog.new_gimp_path
            save_config(self.config); self.statusBar().showMessage("Executable paths updated.", 3000)
            self._update_button_states()

    def _update_button_states(self):
        inkscape_path_set = bool(self.config.get('inkscape_path', ''))
        gimp_path_set = bool(self.config.get('gimp_path', ''))
        any_item_selected = any(item['is_checked'] for item in self.main_gui_file_items)
        can_open_last_inkscape = self.last_created_pdf_path is not None and inkscape_path_set
        self.open_last_pdf_inkscape_button.setEnabled(can_open_last_inkscape)
        if can_open_last_inkscape: self.open_last_pdf_inkscape_button.setToolTip(f"Opens '{os.path.basename(self.last_created_pdf_path)}' in Inkscape")
        elif inkscape_path_set: self.open_last_pdf_inkscape_button.setToolTip("Create a PDF first.")
        else: self.open_last_pdf_inkscape_button.setToolTip("Set Inkscape path in Settings.")
        self.open_selected_inkscape_button.setEnabled(inkscape_path_set and any_item_selected)
        self.open_selected_inkscape_button.setToolTip("Open selected in Inkscape" if inkscape_path_set else "Set Inkscape path in Settings.")
        self.open_gimp_button.setEnabled(gimp_path_set and any_item_selected)
        self.open_gimp_button.setToolTip("Open selected in GIMP" if gimp_path_set else "Set GIMP path in Settings.")
        self.selective_compile_button.setEnabled(any_item_selected)
        self.combine_all_button.setEnabled(len(self.main_gui_file_items) > 0)

    def _update_config_and_refresh_display(self):
        self.config['source_pdf_dir'] = self.source_dir_edit.text().strip()
        self.config['factory_paperwork_dir'] = self.factory_dir_edit.text().strip()
        save_config(self.config)
        self.statusBar().showMessage("Configuration updated. Refreshing display...", 3000)
        self._refresh_factory_folder_display()

    def _browse_dir(self, line_edit, title):
        path = line_edit.text().strip() or os.path.expanduser("~")
        directory = QFileDialog.getExistingDirectory(self, title, path)
        if directory: line_edit.setText(directory); self._update_config_and_refresh_display()

    def _browse_source_dir(self): self._browse_dir(self.source_dir_edit, "Select Source Files Directory")
    def _browse_factory_dir(self): self._browse_dir(self.factory_dir_edit, "Select Factory Paperwork Directory")

    def _copy_files(self):
        source, factory = self.config.get('source_pdf_dir',''), self.config.get('factory_paperwork_dir','')
        if not (os.path.isdir(source) and os.path.isdir(factory)): QMessageBox.warning(self, "Input Error", "Both source and factory directories must be valid."); return
        copied, errors, found = 0, [], False; exts = ('.pdf', '.png')
        try:
            for f_name in os.listdir(source):
                if f_name.lower().endswith(exts):
                    found = True
                    try: shutil.copy2(os.path.join(source, f_name), os.path.join(factory, f_name)); copied += 1
                    except Exception as e: errors.append(f"'{f_name}': {e}")
        except OSError as e: QMessageBox.critical(self, "Dir Error", f"Accessing source: {e}"); return
        msg = ""
        if not found: msg = "No PDF or PNG files in source."
        elif errors: msg = f"Copied {copied} with errors: {', '.join(errors)}"
        else: msg = f"Copied {copied} file(s) to factory folder."
        self.statusBar().showMessage(msg); QMessageBox.information(self, "Copy Result", msg)
        self._refresh_factory_folder_display()

    def _get_processable_files_from_factory_folder(self, show_msg_if_empty=True):
        factory_dir = self.config.get('factory_paperwork_dir', '')
        if not os.path.isdir(factory_dir): 
            if show_msg_if_empty: QMessageBox.warning(self, "Input Error", "Factory folder invalid."); 
            return []
        exts = ('.pdf', '.png')
        try: files = [f for f in os.listdir(factory_dir) if f.lower().endswith(exts)]
        except OSError as e: 
            if show_msg_if_empty: QMessageBox.critical(self, "Dir Error", f"Accessing factory folder: {e}"); 
            return []
        if not files and show_msg_if_empty: QMessageBox.information(self, "No Files", "No PDF or PNG files in factory folder."); 
        return files

    def _combine_all_files_from_main_gui(self):
        factory_dir = self.config.get('factory_paperwork_dir', '')
        if not os.path.isdir(factory_dir): QMessageBox.warning(self, "Error", "Factory Paperwork Folder not set."); return
        if not self.main_gui_file_items: QMessageBox.information(self, "No Files", "No files to combine."); return
        orders_seen = set()
        for item_data in self.main_gui_file_items:
            order = item_data['order']
            if order in orders_seen: QMessageBox.warning(self, "Order Conflict", f"Order {order} duplicated. Correct."); return
            orders_seen.add(order)
        sorted_items = sorted(self.main_gui_file_items, key=lambda x: x['order'])
        ordered_paths = [item['full_path'] for item in sorted_items]
        if not ordered_paths: QMessageBox.information(self, "No Files Ordered", "No files to combine."); return
        self._execute_file_merge(ordered_paths, factory_dir, "Combined_All_Factory_Files.pdf")

    def _selective_compile_files_from_main_gui(self):
        factory_dir = self.config.get('factory_paperwork_dir', '')
        if not os.path.isdir(factory_dir): QMessageBox.warning(self, "Input Error", "Factory folder invalid."); return
        selected_items_data = [{'full_path': item['full_path'], 'order': item['order']} for item in self.main_gui_file_items if item['is_checked']]
        if not selected_items_data: QMessageBox.information(self, "No Selection", "No files selected for compilation."); return
        orders_seen = set()
        for item in selected_items_data:
            if item['order'] in orders_seen: QMessageBox.warning(self, "Order Conflict", f"Order {item['order']} duplicated among selected. Correct."); return
            orders_seen.add(item['order'])
        selected_items_data.sort(key=lambda x: x['order'])
        ordered_paths = [item['full_path'] for item in selected_items_data]
        self.statusBar().showMessage(f"{len(ordered_paths)} file(s) selected for compilation.")
        self._execute_file_merge(ordered_paths, factory_dir, "Selective_Compilation_Output.pdf")
            
    def _execute_file_merge(self, paths_to_merge, save_dir, default_filename_base):
        if not paths_to_merge: QMessageBox.information(self, "Merge Info", "No files to merge."); return
        out_fn, _ = QFileDialog.getSaveFileName(self, "Save Combined PDF", os.path.join(save_dir, default_filename_base), "PDF Files (*.pdf)")
        if not out_fn: self.statusBar().showMessage("Save cancelled."); return
        if not out_fn.lower().endswith(".pdf"): out_fn += ".pdf"
        merger = PdfWriter(); processed_one_file = False
        try:
            self.statusBar().showMessage(f"Combining {len(paths_to_merge)} files...")
            QApplication.processEvents()
            for i, p_path in enumerate(paths_to_merge):
                self.statusBar().showMessage(f"Processing {i+1}/{len(paths_to_merge)}: {os.path.basename(p_path)}...")
                QApplication.processEvents()
                try:
                    if p_path.lower().endswith(".pdf"):
                        merger.append(p_path); processed_one_file = True
                    elif p_path.lower().endswith(".png"):
                        if not PYMUPDF_AVAILABLE: QMessageBox.warning(self, "Missing Dependency", f"PNG '{os.path.basename(p_path)}' skipped. PyMuPDF required."); continue
                        img_doc = fitz.open(p_path); pdf_bytes = img_doc.convert_to_pdf(); img_doc.close()
                        if pdf_bytes: merger.append(PdfReader(io.BytesIO(pdf_bytes))); processed_one_file = True
                        else: QMessageBox.warning(self, "PNG Error", f"Could not convert PNG '{os.path.basename(p_path)}'. Skipped.")
                except Exception as e_app: QMessageBox.warning(self, "File Error", f"Could not process '{os.path.basename(p_path)}': {e_app}\nSkipped.");
            if not processed_one_file or not merger.pages:
                 QMessageBox.warning(self, "Combine Error", "No pages/files processed. PDF not created.")
                 self.statusBar().showMessage("No content to combine."); self.last_created_pdf_path = None; self._update_button_states(); return
            with open(out_fn, "wb") as f_out: merger.write(f_out)
            self.last_created_pdf_path = out_fn
            self.statusBar().showMessage(f"Successfully combined into {os.path.basename(out_fn)}", 7000)
            QMessageBox.information(self, "Combine Successful", f"{len(merger.pages)} page(s) combined into:\n{out_fn}")
        except Exception as e:
            self.last_created_pdf_path = None; self.statusBar().showMessage(f"Error combining: {e}")
            QMessageBox.critical(self, "Combine Error", f"Unexpected error: {e}")
        finally: merger.close(); self._update_button_states()

    def _open_image_blender(self):
        from image_blender_gui import ImageBlenderWindow
        blender_dialog = ImageBlenderWindow(self)
        blender_dialog.exec()

    def _open_last_pdf_in_inkscape(self):
        inkscape_path = self.config.get('inkscape_path', '')
        if not inkscape_path or not os.path.isfile(inkscape_path): QMessageBox.warning(self, "Inkscape Not Configured", "Inkscape path invalid. Configure in Settings."); return
        if not self.last_created_pdf_path or not os.path.exists(self.last_created_pdf_path): QMessageBox.information(self, "No PDF", "No recent PDF or file gone."); return
        try:
            self.statusBar().showMessage(f"Opening {os.path.basename(self.last_created_pdf_path)} in Inkscape...")
            subprocess.Popen([inkscape_path, self.last_created_pdf_path])
            self.statusBar().showMessage(f"Sent to Inkscape.", 5000)
        except Exception as e: QMessageBox.critical(self, "Inkscape Error", f"Could not open in Inkscape: {e}")

    def _get_selected_files_from_main_gui(self):
        return [item['full_path'] for item in self.main_gui_file_items if item['is_checked']]

    def _open_selected_in_inkscape(self):
        inkscape_path = self.config.get('inkscape_path', '')
        if not inkscape_path or not os.path.isfile(inkscape_path): QMessageBox.warning(self, "Inkscape Not Configured", "Inkscape path invalid. Configure in Settings."); return
        selected_files = self._get_selected_files_from_main_gui()
        if not selected_files: QMessageBox.information(self, "No Selection", "No files selected to open in Inkscape."); return
        try:
            for f_path in selected_files:
                 self.statusBar().showMessage(f"Opening {os.path.basename(f_path)} in Inkscape...")
                 subprocess.Popen([inkscape_path, f_path]) # Inkscape typically opens one file per new window/instance unless configured otherwise
            self.statusBar().showMessage(f"Sent {len(selected_files)} file(s) to Inkscape.", 5000)
        except Exception as e: QMessageBox.critical(self, "Inkscape Error", f"Could not open files in Inkscape: {e}")

    def _open_selected_in_gimp(self):
        gimp_path = self.config.get('gimp_path', '')
        if not gimp_path or not os.path.isfile(gimp_path): QMessageBox.warning(self, "GIMP Not Configured", "GIMP path invalid. Configure in Settings."); return
        selected_files = self._get_selected_files_from_main_gui()
        if not selected_files: QMessageBox.information(self, "No Selection", "No files selected to open in GIMP."); return
        try:
            command = [gimp_path] + selected_files
            self.statusBar().showMessage(f"Opening {len(selected_files)} file(s) in GIMP...")
            subprocess.Popen(command)
            self.statusBar().showMessage(f"Sent to GIMP.", 5000)
        except Exception as e: QMessageBox.critical(self, "GIMP Error", f"Could not open in GIMP: {e}")

# --- Application Entry Point ---
if __name__ == '__main__':
    app = QApplication(sys.argv)
    if not PYMUPDF_AVAILABLE:
         QMessageBox.warning(None, "Missing Dependency: PyMuPDF", 
                               "PyMuPDF (fitz) not installed.\n- PDF previews disabled.\n- PNGs cannot be included in combined PDFs.\nInstall: pip install PyMuPDF")
    main_window = PDFToolApp(); main_window.show(); sys.exit(app.exec())
