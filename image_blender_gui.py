import sys
import os

# Add lib directory to Python path
script_dir = os.path.dirname(os.path.abspath(__file__))
lib_dir = os.path.join(script_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QComboBox, QFileDialog, QMessageBox, QScrollArea, QWidget
)
from PySide6.QtGui import QPixmap, QImage, QPainter
from PySide6.QtCore import Qt, QSize
import fitz  # PyMuPDF
from PIL import Image, ImageChops, ImageEnhance, ImageOps

# Define a fixed preview size
PREVIEW_WIDTH = 200
PREVIEW_HEIGHT = 200

def convert_pil_to_qimage(pil_img):
    """Convert PIL Image to QImage."""
    if pil_img.mode == "RGB":
        r, g, b = pil_img.split()
        pil_img = Image.merge("RGB", (b, g, r))
    elif pil_img.mode == "RGBA":
        r, g, b, a = pil_img.split()
        pil_img = Image.merge("RGBA", (b, g, r, a))

    if pil_img.mode == "L": # Grayscale
        # Qt expects 8-bit grayscale
        # Pillow 'L' mode is 8-bit, QImage.Format_Grayscale8
        data = pil_img.tobytes("raw", "L")
        qimage = QImage(data, pil_img.width, pil_img.height, QImage.Format_Grayscale8)
        return qimage

    # For RGB/RGBA
    data = pil_img.tobytes("raw", pil_img.mode)
    if pil_img.mode == "RGB":
        qimage = QImage(data, pil_img.width, pil_img.height, QImage.Format_RGB888)
    elif pil_img.mode == "RGBA":
        qimage = QImage(data, pil_img.width, pil_img.height, QImage.Format_RGBA8888)
    else: # Fallback for other modes, might not display correctly
        pil_img = pil_img.convert("RGBA")
        r, g, b, a = pil_img.split()
        pil_img = Image.merge("RGBA", (b, g, r, a))
        data = pil_img.tobytes("raw", "RGBA")
        qimage = QImage(data, pil_img.width, pil_img.height, QImage.Format_RGBA8888)
    return qimage


class ImageBlenderWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Image Blender")
        self.setMinimumSize(600, 700)

        self.image_path1 = None
        self.image_path2 = None
        self.pil_image1 = None
        self.pil_image2 = None
        self.blended_image = None

        main_layout = QVBoxLayout(self)

        # File 1 Selection
        file1_layout = QHBoxLayout()
        self.file1_label = QLabel("Layer 1 (Base):")
        self.file1_edit = QLineEdit()
        self.file1_edit.setReadOnly(True)
        self.file1_browse_btn = QPushButton("Browse...")
        self.file1_browse_btn.clicked.connect(self._browse_file1)
        file1_layout.addWidget(self.file1_label)
        file1_layout.addWidget(self.file1_edit)
        file1_layout.addWidget(self.file1_browse_btn)
        main_layout.addLayout(file1_layout)

        # File 2 Selection
        file2_layout = QHBoxLayout()
        self.file2_label = QLabel("Layer 2 (Blend):")
        self.file2_edit = QLineEdit()
        self.file2_edit.setReadOnly(True)
        self.file2_browse_btn = QPushButton("Browse...")
        self.file2_browse_btn.clicked.connect(self._browse_file2)
        file2_layout.addWidget(self.file2_label)
        file2_layout.addWidget(self.file2_edit)
        file2_layout.addWidget(self.file2_browse_btn)
        main_layout.addLayout(file2_layout)

        # Previews
        preview_layout = QHBoxLayout()
        self.preview1_label = QLabel("Preview Layer 1")
        self.preview1_label.setFixedSize(PREVIEW_WIDTH, PREVIEW_HEIGHT)
        self.preview1_label.setStyleSheet("border: 1px solid #ccc; background-color: #f0f0f0;")
        self.preview1_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        preview_layout.addWidget(self.preview1_label)

        self.preview2_label = QLabel("Preview Layer 2")
        self.preview2_label.setFixedSize(PREVIEW_WIDTH, PREVIEW_HEIGHT)
        self.preview2_label.setStyleSheet("border: 1px solid #ccc; background-color: #f0f0f0;")
        self.preview2_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        preview_layout.addWidget(self.preview2_label)
        main_layout.addLayout(preview_layout)

        # Blending Options
        blend_options_layout = QHBoxLayout()
        blend_options_layout.addWidget(QLabel("Blend Mode:"))
        self.blend_mode_combo = QComboBox()
        self.blend_mode_combo.addItems([
            "Normal", "Multiply", "Screen", "Overlay", "Darken", 
            "Lighten", "Add", "Subtract", "Difference"
        ])
        blend_options_layout.addWidget(self.blend_mode_combo)
        self.blend_button = QPushButton("Blend Images")
        self.blend_button.clicked.connect(self._blend_images)
        blend_options_layout.addWidget(self.blend_button)
        main_layout.addLayout(blend_options_layout)

        # Result Preview
        main_layout.addWidget(QLabel("Blended Result:"))
        self.result_preview_label = QLabel("Result will appear here")
        self.result_preview_label.setMinimumSize(PREVIEW_WIDTH * 2, PREVIEW_HEIGHT * 2) # Larger preview for result
        self.result_preview_label.setStyleSheet("border: 1px solid #ccc; background-color: #e0e0e0;")
        self.result_preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(self.result_preview_label)
        main_layout.addWidget(scroll_area, 1) # Give stretch factor

        # Save Button
        self.save_button = QPushButton("Save Blended Image")
        self.save_button.clicked.connect(self._save_result)
        self.save_button.setEnabled(False)
        main_layout.addWidget(self.save_button)

    def _browse_file1(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Layer 1 Image", "", "Images (*.png *.jpg *.jpeg *.bmp *.pdf)")
        if file_path:
            self.image_path1 = file_path
            self.file1_edit.setText(file_path)
            self.pil_image1 = self._load_and_convert_image(file_path)
            if self.pil_image1:
                self._update_preview(self.preview1_label, self.pil_image1)

    def _browse_file2(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Layer 2 Image", "", "Images (*.png *.jpg *.jpeg *.bmp *.pdf)")
        if file_path:
            self.image_path2 = file_path
            self.file2_edit.setText(file_path)
            self.pil_image2 = self._load_and_convert_image(file_path)
            if self.pil_image2:
                self._update_preview(self.preview2_label, self.pil_image2)

    def _load_and_convert_image(self, file_path):
        try:
            if file_path.lower().endswith(".pdf"):
                img = self._convert_pdf_to_pil_png(file_path)
            else:
                img = Image.open(file_path)
            
            if img.mode != 'RGBA': # Ensure RGBA for blending
                img = img.convert('RGBA')
            return img
        except Exception as e:
            QMessageBox.warning(self, "Load Error", f"Could not load or convert image '{os.path.basename(file_path)}': {e}")
            return None

    def _convert_pdf_to_pil_png(self, pdf_path, page_num=0):
        try:
            doc = fitz.open(pdf_path)
            if not doc.page_count > page_num:
                raise ValueError(f"Page number {page_num} out of range for PDF with {doc.page_count} pages.")
            page = doc.load_page(page_num)
            pix = page.get_pixmap(alpha=True) # Get pixmap with alpha
            doc.close()
            if pix.alpha:
                img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
            else: # If no alpha channel in PDF source, samples are RGB
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples).convert("RGBA")
            return img
        except Exception as e:
            QMessageBox.warning(self, "PDF Conversion Error", f"Could not convert PDF '{os.path.basename(pdf_path)}' to PNG: {e}")
            return None

    def _update_preview(self, preview_label_widget, pil_image):
        if pil_image:
            qimage = convert_pil_to_qimage(pil_image.copy())
            pixmap = QPixmap.fromImage(qimage)
            scaled_pixmap = pixmap.scaled(PREVIEW_WIDTH, PREVIEW_HEIGHT, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            preview_label_widget.setPixmap(scaled_pixmap)
        else:
            preview_label_widget.setText("Preview N/A")

    def _blend_images(self):
        if not self.pil_image1 or not self.pil_image2:
            QMessageBox.warning(self, "Error", "Please select two images first.")
            return

        img1 = self.pil_image1.copy()
        img2 = self.pil_image2.copy()

        # Ensure images are the same size for most blend modes
        # Resize img2 to match img1's dimensions
        if img1.size != img2.size:
            img2 = img2.resize(img1.size, Image.Resampling.LANCZOS)

        # Ensure RGBA mode for consistency
        if img1.mode != 'RGBA': img1 = img1.convert('RGBA')
        if img2.mode != 'RGBA': img2 = img2.convert('RGBA')

        blend_mode = self.blend_mode_combo.currentText()
        
        try:
            if blend_mode == "Normal":
                # Alpha compositing: img2 on top of img1
                self.blended_image = Image.alpha_composite(img1, img2)
            elif blend_mode == "Add":
                self.blended_image = ImageChops.add(img1, img2)
            elif blend_mode == "Subtract":
                self.blended_image = ImageChops.subtract(img1, img2)
            elif blend_mode == "Multiply":
                self.blended_image = ImageChops.multiply(img1, img2)
            elif blend_mode == "Screen":
                self.blended_image = ImageChops.screen(img1, img2)
            elif blend_mode == "Lighten":
                self.blended_image = ImageChops.lighter(img1, img2)
            elif blend_mode == "Darken":
                self.blended_image = ImageChops.darker(img1, img2)
            elif blend_mode == "Difference":
                self.blended_image = ImageChops.difference(img1, img2)
            elif blend_mode == "Overlay":
                # Pillow doesn't have a direct ImageChops.overlay.
                # This is a common formula for overlay.
                # It requires pixel-level access or more complex ImageChops combinations.
                # For simplicity, we'll use a simplified version or one that might be available.
                # A true overlay is (if base < 0.5: 2*base*blend else: 1 - 2*(1-base)*(1-blend))
                # Let's try a simplified approach or skip if too complex for now.
                # Using Image.blend with a mask or alpha might be an approximation.
                # For now, let's implement a basic overlay logic.
                # This is a common way to implement overlay:
                base = img1.convert('RGB') # Overlay usually defined for RGB
                blend = img2.convert('RGB')
                
                # Create an empty image for the result
                overlay_img = Image.new('RGB', base.size)
                
                base_pixels = base.load()
                blend_pixels = blend.load()
                overlay_pixels = overlay_img.load()

                for y in range(base.height):
                    for x in range(base.width):
                        r_base, g_base, b_base = base_pixels[x, y]
                        r_blend, g_blend, b_blend = blend_pixels[x, y]

                        def overlay_channel(b_ch, l_ch):
                            b_ch /= 255.0
                            l_ch /= 255.0
                            if b_ch < 0.5:
                                return int((2 * b_ch * l_ch) * 255.0)
                            else:
                                return int((1 - 2 * (1 - b_ch) * (1 - l_ch)) * 255.0)

                        r_overlay = overlay_channel(r_base, r_blend)
                        g_overlay = overlay_channel(g_base, g_blend)
                        b_overlay = overlay_channel(b_base, b_blend)
                        
                        overlay_pixels[x, y] = (r_overlay, g_overlay, b_overlay)
                self.blended_image = overlay_img.convert('RGBA') # Convert back to RGBA
            else:
                QMessageBox.warning(self, "Blend Error", f"Blend mode '{blend_mode}' not implemented yet.")
                return

            if self.blended_image:
                q_blended_img = convert_pil_to_qimage(self.blended_image.copy())
                pixmap = QPixmap.fromImage(q_blended_img)
                # Scale to fit the label while maintaining aspect ratio
                scaled_pixmap = pixmap.scaled(self.result_preview_label.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                self.result_preview_label.setPixmap(scaled_pixmap)
                self.save_button.setEnabled(True)
            else:
                self.result_preview_label.setText("Blending failed.")
                self.save_button.setEnabled(False)

        except Exception as e:
            QMessageBox.critical(self, "Blending Error", f"An error occurred during blending: {e}")
            self.blended_image = None
            self.result_preview_label.setText("Blending Error.")
            self.save_button.setEnabled(False)


    def _save_result(self):
        if not self.blended_image:
            QMessageBox.warning(self, "Error", "No blended image to save.")
            return

        file_path, _ = QFileDialog.getSaveFileName(self, "Save Blended Image", "", "PNG Image (*.png);;JPEG Image (*.jpg *.jpeg)")
        if file_path:
            try:
                # Ensure the image is in a savable mode (e.g., RGB for JPEG)
                save_image = self.blended_image
                if file_path.lower().endswith((".jpg", ".jpeg")) and save_image.mode == 'RGBA':
                    save_image = save_image.convert('RGB') # JPEG doesn't support alpha
                
                save_image.save(file_path)
                QMessageBox.information(self, "Saved", f"Blended image saved to {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Save Error", f"Could not save image: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    # For testing the dialog independently
    # You might need to create a dummy main window or ensure resources are available
    # if the dialog relies on external configurations passed from a parent.
    # For now, it should be self-contained enough.
    
    # Example of how to load config if needed (not used in this standalone version)
    # from pdf_manager import load_config # Assuming pdf_manager.py is in PYTHONPATH
    # config = load_config() 
    
    blender_dialog = ImageBlenderWindow()
    blender_dialog.show()
    sys.exit(app.exec())

