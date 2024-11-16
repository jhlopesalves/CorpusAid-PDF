import sys
import os
import fitz
from html.parser import HTMLParser
from collections import defaultdict
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QLabel, QFileDialog,
    QTextEdit, QVBoxLayout, QHBoxLayout, QWidget, QProgressBar, QMessageBox,
    QComboBox, QScrollArea, QSplitter, QToolBar, QListWidget, QListWidgetItem,
    QGroupBox, QStatusBar, QTabWidget, QLineEdit, QDialog, QSizePolicy
)
from PySide6.QtCore import Qt, QThread, Signal, QPropertyAnimation, QSize
from PySide6.QtGui import (
    QPixmap, QImage, QIcon, QTextCursor, QAction, QKeySequence,
    QShortcut, QColor, QPainter
)
from docx import Document
import logging

def setup_logging():
    """Configures the logging system."""

    log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(module)s:%(lineno)d - %(message)s')
    log_file = "pdf_extractor.log"  # Name of the log file

    # File handler (writes logs to a file)
    file_handler = logging.FileHandler(log_file, mode='w')  # 'w' mode overwrites the log file each time
    file_handler.setFormatter(log_formatter)
    file_handler.setLevel(logging.DEBUG)  # Log everything to the file

    # Console handler (prints logs to the console) — Optional
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(log_formatter)
    console_handler.setLevel(logging.INFO) # Log INFO and above to the console

    # Root logger configuration
    root_logger = logging.getLogger()
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler) # Add console handler if you want console output
    root_logger.setLevel(logging.DEBUG) # Set the lowest level to DEBUG to capture everything

# Utility function to get resource paths
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# ErrorHandler class for centralized error and warning messages
class ErrorHandler:
    @staticmethod
    def show_error(message, title="Error", parent=None):
        logging.error(message)
        QMessageBox.critical(parent, title, message)
    
    @staticmethod
    def show_warning(message, title="Warning", parent=None):
        logging.warning(message)
        QMessageBox.warning(parent, title, message)

# ThemeManager class to handle dark and light themes
class ThemeManager:
    def __init__(self):
        self.dark_theme = True
        self.custom_colors = {
            "primary": "#518FBC",
            "secondary": "#325F84",
            "background": "#1E1E1E",
            "text": "#FFFFFF",
            "accent": "#FFB900",
            "icon_dark": "#FFFFFF",
            "icon_light": "#000000",
            "border": "#3F3F3F",
            "widget_background": "#2D2D2D",
            "report_background": "#2D2D2D",
            "report_text": "#FFFFFF",
            "report_description": "#A0A0A0"
        }

    def toggle_theme(self):
        self.dark_theme = not self.dark_theme
        if self.dark_theme:
            self.custom_colors.update({
                "background": "#1E1E1E",
                "text": "#FFFFFF",
                "widget_background": "#2D2D2D",
                "report_background": "#2D2D2D",
                "border": "#3F3F3F",
                "report_text": "#FFFFFF",
                "report_description": "#A0A0A0"
            })
        else:
            self.custom_colors.update({
                "background": "#F0F0F0",
                "text": "#000000",
                "widget_background": "#FFFFFF",
                "report_background": "#FFFFFF",
                "border": "#CCCCCC",
                "report_text": "#000000",
                "report_description": "#505050"
            })

    def get_stylesheet(self):
        return f"""
            QMainWindow, QWidget {{
                background-color: {self.custom_colors['background']};
                color: {self.custom_colors['text']};
                font-family: Roboto, sans-serif;
                font-size: 10pt;
            }}
            QLabel {{
                color: {self.custom_colors['text']};
                background-color: transparent;
            }}
            QLabel#sectionLabel {{
                color: {self.custom_colors['primary']};
                font-size: 16px;
                font-weight: bold;
                margin-bottom: 10px;
            }}
            QLabel#modeDescriptionLabel {{
                font-size: 11px;
                color: {self.custom_colors['text']};
                opacity: 0.7;
                background: none;
                padding-left: 5px;
                margin-bottom: 10px;
            }}
            QMenuBar {{
                background-color: {self.custom_colors['background']};
                color: {self.custom_colors['text']};
            }}
            QMenuBar::item {{
                background-color: transparent;
                padding: 4px 10px;
            }}
            QMenuBar::item:selected {{
                background-color: {self.custom_colors['secondary']};
            }}
            QMenu {{
                background-color: {self.custom_colors['background']};
                color: {self.custom_colors['text']};
                border: 1px solid {self.custom_colors['border']};
            }}
            QMenu::item:selected {{
                background-color: {self.custom_colors['secondary']};
            }}
            QToolBar {{
                background-color: {self.custom_colors['widget_background']};
                border-bottom: 1px solid {self.custom_colors['border']};
                spacing: 5px;
            }}
            QPushButton#toolbarButton, QPushButton#zoomButton, QPushButton#fitWidthButton, QPushButton#searchPrevButton, QPushButton#searchNextButton, QPushButton#pagePrevButton, QPushButton#pageNextButton {{
                padding: 6px 12px;
                background-color: {self.custom_colors['primary']};
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 12px;
                min-width: 15px;
                min-height: 15px;
            }}
            QPushButton#toolbarButton:hover, QPushButton#zoomButton:hover, QPushButton#fitWidthButton:hover, QPushButton#searchPrevButton:hover, QPushButton#searchNextButton:hover, QPushButton#pagePrevButton:hover, QPushButton#pageNextButton:hover {{
                background-color: {self.custom_colors['secondary']};
            }}
            QPushButton#toolbarButton:disabled, QPushButton#zoomButton:disabled, QPushButton#fitWidthButton:disabled, QPushButton#searchPrevButton:disabled, QPushButton#searchNextButton:disabled, QPushButton#pagePrevButton:disabled, QPushButton#pageNextButton:disabled {{
                background-color: #555555;
                color: #AAAAAA;
            }}
            QTabWidget::pane {{
                border: 1px solid {self.custom_colors['border']};
                border-radius: 5px;
            }}
            QTabBar::tab {{
                background-color: {self.custom_colors['widget_background']};
                color: {self.custom_colors['text']};
                padding: 5px 10px;
                border-top-left-radius: 3px;
                border-top-right-radius: 3px;
            }}
            QTabBar::tab:selected {{
                background-color: {self.custom_colors['primary']};
                color: {self.custom_colors['text']};
            }}
            QListWidget {{
                background-color: {self.custom_colors['widget_background']};
                color: {self.custom_colors['text']};
                border: 1px solid {self.custom_colors['border']};
                border-radius: 5px;
            }}
            QGroupBox {{
                border: none;
                margin-top: 0;
                padding: 0;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 3px;
                color: {self.custom_colors['primary']};
                font-size: 14px;
                font-weight: bold;
            }}
            QComboBox {{
                background-color: {self.custom_colors['primary']};
                color: {self.custom_colors['text']};
                border: 1px solid {self.custom_colors['border']};
                border-radius: 3px;
                padding: 5px;
                min-width: 100px;
            }}
            QComboBox::drop-down {{
                border: none;
                width: 20px;
            }}
            QComboBox::down-arrow {{
                image: none;
            }}
            QComboBox:hover {{
                background-color: {self.custom_colors['secondary']};
            }}
            QComboBox QAbstractItemView {{
                background-color: {self.custom_colors['widget_background']};
                color: {self.custom_colors['text']};
                selection-background-color: {self.custom_colors['primary']};
                selection-color: {self.custom_colors['text']};
                border: 1px solid {self.custom_colors['border']};
            }}
            QComboBox QAbstractItemView::item:selected {{
                background-color: {self.custom_colors['primary']};
                color: {self.custom_colors['text']};
            }}
            QComboBox QAbstractItemView::item:hover {{
                background-color: {self.custom_colors['secondary']};
                color: {self.custom_colors['text']};
            }}
            QPushButton#navigationButton {{
                padding: 8px 16px;
                background-color: {self.custom_colors['primary']};
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 14px;
                min-width: 13px;
                min-height: 13px;
            }}
            QPushButton#navigationButton:hover {{
                background-color: {self.custom_colors['secondary']};
            }}
            QPushButton#navigationButton:disabled {{
                background-color: {self.custom_colors['primary']};
                opacity: 0.5;
            }}
            QProgressBar {{
                border: 1px solid {self.custom_colors['border']};
                border-radius: 3px;
                text-align: center;
                height: 20px;
            }}
            QProgressBar::chunk {{
                background-color: {self.custom_colors['accent']};
                border-radius: 3px;
            }}
            QLineEdit {{
                border: 1px solid {self.custom_colors['border']};
                background-color: {self.custom_colors['widget_background']};
                color: {self.custom_colors['text']};
                padding: 5px;
                border-radius: 3px;
            }}
            QSplitter::handle {{
                background-color: {self.custom_colors['border']};
            }}
            QSplitter::handle:horizontal {{
                height: 1px;
            }}
            QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {{
                background-color: {self.custom_colors['accent']};
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}
            QStatusBar {{
                background-color: {self.custom_colors['widget_background']};
                color: {self.custom_colors['text']};
                border-top: 1px solid {self.custom_colors['border']};
            }}
            QToolTip {{
                background-color: {self.custom_colors['widget_background']};
                color: {self.custom_colors['text']};
                border: 1px solid {self.custom_colors['border']};
                padding: 5px;
                border-radius: 3px;
            }}
        """
# ToastNotification class for non-intrusive notifications
class ToastNotification(QLabel):
    def __init__(self, message, parent=None):
        super().__init__(parent)
        self.setText(message)
        self.setStyleSheet("""
            QLabel {
                background-color: rgba(0, 0, 0, 180);
                color: white;
                padding: 10px;
                border-radius: 5px;
            }
        """)
        self.setAlignment(Qt.AlignCenter)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.adjustSize()
        self.animation = QPropertyAnimation(self, b"windowOpacity")
        self.animation.setDuration(2000)
        self.animation.setStartValue(1)
        self.animation.setEndValue(0)
        self.animation.finished.connect(self.close)

    def show_notification(self, duration=2000):
        self.show()
        self.animation.setDuration(duration)
        self.animation.start()

# HelpDialog class for displaying help information
class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PDF Text Extractor Help")
        self.setMinimumSize(600, 400)
        layout = QVBoxLayout()
        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setHtml("""
            <h2>PDF Text Extractor Help</h2>
            <p>Welcome to PDF Text Extractor! Here's how to use the application:</p>
            <ol>
                <li><strong>Open PDF(s):</strong> Click the "Open PDF(s)" button or drag and drop PDF files into the application window.</li>
                <li><strong>Select Output Folder:</strong> Click the "Save As" button to choose where to save the extracted text.</li>
                <li><strong>Choose Extraction Mode:</strong>
                    <ul>
                        <li><strong>Column-aware:</strong> Best for documents with multiple columns (academic papers, magazines, newspapers). Intelligently detects and preserves column layout.</li>
                        <li><strong>Layout-preserved:</strong> Maintains exact document formatting, ideal for forms, code listings, or specially formatted text.</li>
                    </ul>
                </li>
                <li><strong>Select Output Format:</strong> Choose from TXT, HTML, Markdown, or DOCX formats for the extracted text.</li>
                <li><strong>Process PDF(s):</strong> Click the "Process PDF(s)" button to start extraction. Progress will be shown in the progress bar.</li>
                <li><strong>View Results:</strong> Switch between the original PDF preview and the extracted text using the tabs.</li>
                <li><strong>Search:</strong> Use the search bar to find specific text within the extracted content.</li>
            </ol>
            <h3>Extraction Modes Explained:</h3>
            <p><strong>Column-aware Mode</strong></p>
            <ul>
                <li>Intelligently detects and handles multi-column layouts</li>
                <li>Preserves the natural reading flow across columns</li>
                <li>Perfect for:
                    <ul>
                        <li>Academic papers and journals</li>
                        <li>Magazines and newspapers</li>
                        <li>Documents with complex column structures</li>
                    </ul>
                </li>
            </ul>
            <p><strong>Layout-preserved Mode</strong></p>
            <ul>
                <li>Maintains exact document formatting and structure</li>
                <li>Preserves special characters, spacing, and indentation</li>
                <li>Ideal for:
                    <ul>
                        <li>Forms and structured documents</li>
                        <li>Technical documentation</li>
                        <li>Code listings or tabulated data</li>
                        <li>Documents with precise formatting requirements</li>
                    </ul>
                </li>
            </ul>
            <h3>Additional Features:</h3>
            <ul>
                <li>Support for multiple file processing</li>
                <li>Real-time preview of PDF pages</li>
                <li>Advanced search functionality</li>
                <li>Multiple export formats</li>
                <li>Dark/Light themes</li>
                <li>High contrast mode for accessibility</li>
                <li>Drag and drop support</li>
                <li>Zoom controls for detailed preview</li>
            </ul>
            <h3>Keyboard Shortcuts:</h3>
            <ul>
                <li>Ctrl+O: Open PDF(s)</li>
                <li>Ctrl+S: Select output folder</li>
                <li>Ctrl+P: Process PDF(s)</li>
                <li>Ctrl+F: Find text</li>
                <li>Ctrl+C: Copy selected text</li>
                <li>Ctrl++ / Ctrl+-: Zoom in/out preview</li>
            </ul>
        """)
        layout.addWidget(help_text)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)
        self.setLayout(layout)

class PreviewWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.zoom_factor = 1.0
        self.current_page = 0
        self.total_pages = 0
        self.current_doc = None
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(False)
        self.scroll_area.setAlignment(Qt.AlignCenter)
        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.scroll_area.setWidget(self.preview_label)
        layout.addWidget(self.scroll_area)

        controls_layout = QHBoxLayout()
        nav_layout = QHBoxLayout()

        self.prev_page_btn = QPushButton()
        self.prev_page_btn.setObjectName("navigationButton")
        self.prev_page_btn.setFixedWidth(25)
        self.prev_page_btn.setIcon(QIcon.fromTheme("go-previous"))
        self.prev_page_btn.clicked.connect(self.previous_page)
        self.prev_page_btn.setEnabled(False)

        self.page_label = QLabel("Page 0 of 0")
        self.page_label.setFixedWidth(100)
        self.page_label.setAlignment(Qt.AlignCenter)

        self.next_page_btn = QPushButton()
        self.next_page_btn.setObjectName("navigationButton")
        self.next_page_btn.setFixedWidth(25)
        self.next_page_btn.setIcon(QIcon.fromTheme("go-next"))
        self.next_page_btn.clicked.connect(self.next_page)
        self.next_page_btn.setEnabled(False)

        nav_layout.addWidget(self.prev_page_btn)
        nav_layout.addWidget(self.page_label)
        nav_layout.addWidget(self.next_page_btn)

        zoom_layout = QHBoxLayout()

        self.zoom_out_btn = QPushButton()
        self.zoom_out_btn.setObjectName("zoomButton")
        self.zoom_out_btn.setFixedWidth(25)
        self.zoom_out_btn.setIcon(QIcon.fromTheme("zoom-out"))
        self.zoom_out_btn.clicked.connect(self.zoom_out)

        self.zoom_level = QLabel("100%")
        self.zoom_level.setFixedWidth(30)
        self.zoom_level.setAlignment(Qt.AlignCenter)

        self.zoom_in_btn = QPushButton()
        self.zoom_in_btn.setObjectName("zoomButton")
        self.zoom_in_btn.setFixedWidth(25)
        self.zoom_in_btn.setIcon(QIcon.fromTheme("zoom-in"))
        self.zoom_in_btn.clicked.connect(self.zoom_in)

        self.fit_width_btn = QPushButton("Fit Width")
        self.fit_width_btn.setObjectName("fitWidthButton")
        self.fit_width_btn.clicked.connect(self.fit_to_width)

        zoom_layout.addWidget(self.zoom_out_btn)
        zoom_layout.addWidget(self.zoom_level)
        zoom_layout.addWidget(self.zoom_in_btn)
        zoom_layout.addWidget(self.fit_width_btn)

        controls_layout.addLayout(nav_layout)
        controls_layout.addStretch()
        controls_layout.addLayout(zoom_layout)

        layout.addLayout(controls_layout)

        QShortcut(QKeySequence(Qt.Key_Left), self, self.previous_page)
        QShortcut(QKeySequence(Qt.Key_Right), self, self.next_page)
        
    def set_document(self, pdf_path):
        try:
            if self.current_doc:
                self.current_doc.close()
            self.current_doc = fitz.open(pdf_path)
            self.total_pages = self.current_doc.page_count
            self.current_page = 0
            self.update_navigation()
            self.load_current_page()
        except Exception as e:
            ErrorHandler.show_error(f"Failed to load PDF: {str(e)}", "Load Error", self)
            self.current_doc = None
            self.total_pages = 0
            self.current_page = 0
            self.update_navigation()
            self.preview_label.clear()

    def load_current_page(self):
        try:
            if self.current_doc and 0 <= self.current_page < self.total_pages:
                page = self.current_doc.load_page(self.current_page)
                zoom_matrix = fitz.Matrix(self.zoom_factor, self.zoom_factor)
                pix = page.get_pixmap(matrix=zoom_matrix)
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(img)
                self.preview_label.setPixmap(pixmap)
                self.preview_label.resize(pixmap.size())
        except Exception as e:
            ErrorHandler.show_error(f"Error loading page: {str(e)}", "Page Load Error", self)
            self.preview_label.clear()

    def update_navigation(self):
        self.prev_page_btn.setEnabled(self.current_page > 0)
        self.next_page_btn.setEnabled(self.current_page < self.total_pages - 1)
        self.page_label.setText(f"Page {self.current_page + 1} of {self.total_pages}")

    def next_page(self):
        if self.current_doc and self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.load_current_page()
            self.update_navigation()

    def previous_page(self):
        if self.current_doc and self.current_page > 0:
            self.current_page -= 1
            self.load_current_page()
            self.update_navigation()

    def zoom_in(self):
        if self.zoom_factor < 5.0:
            self.zoom_factor *= 1.2
            self.zoom_level.setText(f"{int(self.zoom_factor * 100)}%")
            self.load_current_page()

    def zoom_out(self):
        if self.zoom_factor > 0.2:
            self.zoom_factor /= 1.2
            self.zoom_level.setText(f"{int(self.zoom_factor * 100)}%")
            self.load_current_page()

    def fit_to_width(self):
        if self.current_doc and self.current_doc.page_count > 0:
            page = self.current_doc.load_page(self.current_page)
            page_width = page.rect.width
            available_width = self.scroll_area.viewport().width() - 20
            self.zoom_factor = available_width / page_width
            self.zoom_level.setText(f"{int(self.zoom_factor * 100)}%")
            self.load_current_page()

    def resizeEvent(self, event):
        if self.current_doc and self.current_page < self.total_pages:
            page = self.current_doc.load_page(self.current_page)
            page_width = page.rect.width
            available_width = self.scroll_area.viewport().width() - 20
            self.zoom_factor = available_width / page_width
            self.zoom_level.setText(f"{int(self.zoom_factor * 100)}%")
            self.load_current_page()
        super().resizeEvent(event)

# ExtractionThread class for handling PDF extraction in a separate thread
class ExtractionThread(QThread):
    progress = Signal(int, str)
    finished = Signal(str)
    error = Signal(str)
    preview_ready = Signal(object)
    toast = Signal(str)
    extracted_text = Signal(str, str)

    class SpecialCharParser(HTMLParser):
        def __init__(self):
            super().__init__()
            self.text = []
            self.in_sup = False
            self.in_sub = False

        def handle_starttag(self, tag, attrs):
            if tag == 'sup':
                self.in_sup = True
            elif tag == 'sub':
                self.in_sub = True

        def handle_endtag(self, tag):
            if tag == 'sup':
                self.in_sup = False
            elif tag == 'sub':
                self.in_sub = False

        def handle_data(self, data):
            if self.in_sup:
                converted = ''.join([chr(ord('⁰') + int(c)) if c.isdigit() else c for c in data])
                self.text.append(converted)
            elif self.in_sub:
                converted = ''.join([chr(ord('₀') + int(c)) if c.isdigit() else c for c in data])
                self.text.append(converted)
            else:
                self.text.append(data)

        def get_text(self):
            return ''.join(self.text)

    def __init__(self, pdf_paths, output_path, extraction_mode, output_format):
        super().__init__()
        self.pdf_paths = pdf_paths
        self.output_path = output_path
        self.extraction_mode = extraction_mode
        self.output_format = output_format

    def analyze_layout(self, page):
        blocks = page.get_text("dict")
        if not isinstance(blocks, dict) or "blocks" not in blocks:
            return {"columns": 1, "boundaries": []}
        blocks = blocks["blocks"]
        page_width = page.rect.width
        x_positions = []
        for block in blocks:
            if isinstance(block, dict) and "bbox" in block:
                bbox = block["bbox"]
                if len(bbox) >= 4:
                    x_positions.extend([bbox[0], bbox[2]])
        if not x_positions:
            return {"columns": 1, "boundaries": []}
        x_positions = sorted(set(x_positions))
        gaps = []
        for i in range(len(x_positions) - 1):
            gap = x_positions[i + 1] - x_positions[i]
            if gap > page_width * 0.08:
                gaps.append((gap, x_positions[i]))
        significant_gaps = [g for g in gaps if g[0] > page_width * 0.08]
        return {
            "columns": len(significant_gaps) + 1,
            "boundaries": sorted([g[1] for g in significant_gaps])
        }

    def extract_with_columns(self, page):
        try:
            html_text = page.get_text("html")
            blocks = page.get_text("dict", sort=True)
            self.analyze_layout(page)
            parser = self.SpecialCharParser()
            parser.feed(html_text)
            special_chars_text = parser.get_text()
            columns = defaultdict(list)
            page_width = page.rect.width
            page_height = page.rect.height
            x_coordinates = []
            for block in blocks["blocks"]:
                if "bbox" in block:
                    x_coordinates.extend([block["bbox"][0], block["bbox"][2]])
            if x_coordinates:
                x_coordinates.sort()
                gaps = []
                for i in range(len(x_coordinates) - 1):
                    gap = x_coordinates[i + 1] - x_coordinates[i]
                    if gap > page_width * 0.05:
                        gaps.append((gap, (x_coordinates[i] + x_coordinates[i + 1]) / 2))
                gaps.sort(reverse=True)
                significant_gaps = [g for g in gaps if g[0] > page_width * 0.05]
                column_boundaries = sorted([g[1] for g in significant_gaps[:2]])
                for block in blocks["blocks"]:
                    if "bbox" not in block or "lines" not in block:
                        continue
                    bbox = block["bbox"]
                    block_center = (bbox[0] + bbox[2]) / 2
                    col_idx = 0
                    for boundary in column_boundaries:
                        if block_center > boundary:
                            col_idx += 1
                    columns[col_idx].append((bbox[1], block))
                final_text = ""
                for col_idx in sorted(columns.keys()):
                    col_blocks = sorted(columns[col_idx], key=lambda x: x[0])
                    column_text = ""
                    last_y = None
                    for y_pos, block in col_blocks:
                        if last_y is not None:
                            gap = y_pos - last_y
                            if gap > page_height * 0.02:
                                column_text += "\n"
                            if gap > page_height * 0.05:
                                column_text += "\n"
                        block_text = ""
                        for line in block["lines"]:
                            line_text = " ".join(span.get("text", "") for span in line.get("spans", []))
                            if line_text.strip():
                                block_text += line_text.strip() + " "
                        column_text += block_text.strip() + "\n"
                        last_y = y_pos + block["bbox"][3] - block["bbox"][1]
                    if column_text.strip():
                        if final_text:
                            final_text += "\n\n"
                        final_text += column_text.strip()
                final_text = self.merge_special_characters(final_text, special_chars_text)
                return final_text
            return special_chars_text
        except Exception as e:
            return f"Error in column extraction: {str(e)}"

    def extract_with_layout(self, page):
        try:
            html_text = page.get_text("html")
            parser = self.SpecialCharParser()
            parser.feed(html_text)
            text = parser.get_text()
            text = text.replace('\u200b', '')
            text = text.strip()
            return text
        except Exception as e:
            return f"Error extracting text with layout: {str(e)}"

    def merge_special_characters(self, column_text, special_text):
        if not column_text.strip():
            return special_text
        if not special_text.strip():
            return column_text
        words_special = special_text.split()
        words_column = column_text.split()
        merged = []
        i = 0
        j = 0
        while i < len(words_column) and j < len(words_special):
            if words_column[i].lower() == words_special[j].lower():
                merged.append(words_special[j])
                i += 1
                j += 1
            else:
                merged.append(words_column[i])
                i += 1
        merged.extend(words_column[i:])
        return ' '.join(merged)

    def save_as_format(self, text, output_path, pdf_path):
        try:
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            if self.output_format == "TXT":
                file_path = os.path.join(output_path, f"{base_name}.txt")
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(text)
            elif self.output_format == "HTML":
                file_path = os.path.join(output_path, f"{base_name}.html")
                html_content = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <style>
                        body {{ 
                            font-family: Arial, sans-serif; 
                            line-height: 1.6; 
                            margin: 2em;
                        }}
                        .page-break {{ 
                            border-top: 2px dashed #999;
                            margin: 20px 0;
                            padding-top: 20px;
                        }}
                    </style>
                </head>
                <body>
                    <pre>{text}</pre>
                </body>
                </html>
                """
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(html_content)
            elif self.output_format == "Markdown":
                file_path = os.path.join(output_path, f"{base_name}.md")
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(text)
            elif self.output_format == "DOCX":
                file_path = os.path.join(output_path, f"{base_name}.docx")
                document = Document()
                for line in text.split('\n'):
                    document.add_paragraph(line)
                document.save(file_path)
            return file_path
        except Exception as e:
            raise Exception(f"Error saving file: {str(e)}")

    def run(self):
        try:
            total_pdfs = len(self.pdf_paths)
            for idx, pdf_path in enumerate(self.pdf_paths):
                self.progress.emit(int((idx / total_pdfs) * 100), f"Processing {os.path.basename(pdf_path)} ({idx + 1}/{total_pdfs})")
                doc = fitz.open(pdf_path)
                self.generate_preview(doc, pdf_path)
                total_pages = doc.page_count
                extracted_text = ""
                for page_num in range(total_pages):
                    page = doc.load_page(page_num)
                    if self.extraction_mode == "Column-aware":
                        text = self.extract_with_columns(page)
                    else:
                        text = self.extract_with_layout(page)
                    if text is None:
                        text = ""
                    if text.strip():
                        extracted_text += f"\n--- Page {page_num + 1} ---\n\n{text}\n\n"
                    progress_percent = int(((idx + (page_num + 1)/total_pages) / total_pdfs) * 100)
                    self.progress.emit(progress_percent, f"Processed page {page_num + 1} of {total_pages}")
                output_file = self.save_as_format(extracted_text, self.output_path, pdf_path)
                self.finished.emit(output_file)
                self.extracted_text.emit(pdf_path, extracted_text)
                self.toast.emit(f"Extraction complete for {os.path.basename(pdf_path)}")
                doc.close()
        except Exception as e:
            self.error.emit(str(e))

    def generate_preview(self, doc, pdf_path):
        try:
            if doc.page_count > 0:
                page = doc.load_page(0)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                self.preview_ready.emit(img)
        except Exception as e:
            self.error.emit(f"Error generating preview for {pdf_path}: {str(e)}")

# MainWindow class for the main application window
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Text Extractor")
        self.setGeometry(100, 100, 1400, 900)
        self.setAcceptDrops(True)
        self.pdf_paths = []
        self.extracted_texts = {}
        self.output_path = ""
        self.current_theme = "dark"
        self.search_positions = []
        self.current_search_index = 0
        self.theme_manager = ThemeManager()
        self.create_menu_bar()
        self.create_toolbar()
        self.create_ui()
        self.create_status_bar()
        self.apply_theme(self.current_theme)
        self.init_shortcuts()
        self.init_connections()

    def apply_theme(self, theme):
        logging.debug(f"Attempting to apply theme: {theme}")
        if theme == 'dark' and not self.theme_manager.dark_theme:
            logging.debug("Toggling to dark theme.")
            self.theme_manager.toggle_theme()
            self.current_theme = 'dark'
        elif theme == 'light' and self.theme_manager.dark_theme:
            logging.debug("Toggling to light theme.")
            self.theme_manager.toggle_theme()
            self.current_theme = 'light'
        else:
            logging.debug("No theme change required.")
        stylesheet = self.theme_manager.get_stylesheet()
        self.setStyleSheet(stylesheet)
        self.status_bar.setStyleSheet(stylesheet)
        self.update_icon_colors()
        self.mode_description.setStyleSheet(f"""
            font-size: 11px;
            color: {self.theme_manager.custom_colors['text']};
            opacity: 0.7;
            background: none;
            padding-left: 5px;
            margin-bottom: 10px;
        """)

    def create_menu_bar(self):
        menubar = self.menuBar()
        file_menu = menubar.addMenu('&File')
        open_action = QAction(QIcon.fromTheme("document-open"), "Open PDF(s)...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.setToolTip("Open PDF files (Ctrl + O)")
        open_action.triggered.connect(self.select_pdf)
        save_action = QAction(QIcon.fromTheme("document-save"), "Save As...", self)
        save_action.setShortcut("Ctrl+S")
        save_action.setToolTip("Select Output Folder (Ctrl + S)")
        save_action.triggered.connect(self.select_output_folder)
        exit_action = QAction(QIcon.fromTheme("application-exit"), "Exit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.setToolTip("Exit Application (Ctrl + Q)")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(open_action)
        file_menu.addAction(save_action)
        file_menu.addSeparator()
        file_menu.addAction(exit_action)

        edit_menu = menubar.addMenu('&Edit')
        copy_action = QAction(QIcon.fromTheme("edit-copy"), "Copy", self)
        copy_action.setShortcut("Ctrl+C")
        copy_action.setToolTip("Copy Selected Text (Ctrl + C)")
        copy_action.triggered.connect(self.copy_text)
        find_action = QAction(QIcon.fromTheme("edit-find"), "Find...", self)
        find_action.setShortcut("Ctrl+F")
        find_action.setToolTip("Find Text (Ctrl + F)")
        find_action.triggered.connect(self.focus_search)
        edit_menu.addAction(copy_action)
        edit_menu.addAction(find_action)

        settings_menu = menubar.addMenu('&Settings')
        theme_menu = settings_menu.addMenu("Theme")
        
        dark_action = QAction("Dark", self)
        dark_action.triggered.connect(lambda: self.apply_theme('dark'))
        
        light_action = QAction("Light", self)
        light_action.triggered.connect(lambda: self.apply_theme('light'))
        
        theme_menu.addAction(dark_action)
        theme_menu.addAction(light_action)

        help_menu = menubar.addMenu('&Help')
        about_action = QAction(QIcon.fromTheme("help-about"), "About", self)
        about_action.triggered.connect(self.show_about)
        help_action = QAction(QIcon.fromTheme("help-contents"), "Help", self)
        help_action.triggered.connect(self.show_help)
        help_menu.addAction(help_action)
        help_menu.addAction(about_action)

    def create_toolbar(self):
        self.toolbar = QToolBar("Main Toolbar")
        self.toolbar.setIconSize(QSize(16, 16))
        self.addToolBar(self.toolbar)
        
        open_btn = QPushButton("Open PDF(s)")
        open_btn.setIcon(QIcon.fromTheme("document-open"))
        open_btn.setToolTip("Open PDF files (Ctrl + O)")
        open_btn.setObjectName("toolbarButton")
        open_btn.clicked.connect(self.select_pdf)
        
        save_btn = QPushButton("Save As")
        save_btn.setIcon(QIcon.fromTheme("document-save"))
        save_btn.setToolTip("Select Output Folder (Ctrl + S)")
        save_btn.setObjectName("toolbarButton")
        save_btn.clicked.connect(self.select_output_folder)
        
        process_btn = QPushButton("Process PDF(s)")
        process_btn.setIcon(QIcon.fromTheme("system-run"))
        process_btn.setEnabled(False)
        process_btn.setToolTip("Start Extraction (Ctrl + P)")
        process_btn.setObjectName("toolbarButton")
        process_btn.clicked.connect(self.start_extraction)
        
        self.process_btn = process_btn
        
        self.toolbar.addWidget(open_btn)
        self.toolbar.addWidget(save_btn)
        self.toolbar.addWidget(process_btn)

    def create_status_bar(self):
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")

    def create_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        splitter = QSplitter(Qt.Horizontal)
        
        # Left Widget Setup
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(10, 10, 10, 10)
        left_layout.setSpacing(10)
        
        files_label = QLabel("Selected PDF(s):")
        files_label.setObjectName("sectionLabel")
        self.files_list = QListWidget()
        self.files_list.setSelectionMode(QListWidget.MultiSelection)
        self.files_list.setToolTip("List of selected PDF files. You can select multiple files.")
        self.files_list.itemClicked.connect(self.on_file_selected)
        left_layout.addWidget(files_label)
        left_layout.addWidget(self.files_list)
        
        extraction_label = QLabel("Extraction Options:")
        extraction_label.setObjectName("sectionLabel")
        left_layout.addWidget(extraction_label)
        
        options_group = QGroupBox()
        options_group.setObjectName("extractionOptionsGroup")
        options_layout = QVBoxLayout()
        
        mode_layout = QHBoxLayout()
        mode_label = QLabel("Mode:")
        self.extraction_mode = QComboBox()
        self.extraction_mode.addItems(["Column-aware", "Layout-preserved"])
        self.extraction_mode.setToolTip("Column-aware: Best for documents with columns (papers, magazines)\nLayout-preserved: Maintains exact document formatting")
        mode_layout.addWidget(mode_label)
        mode_layout.addWidget(self.extraction_mode)
        
        self.mode_description = QLabel()
        self.mode_description.setObjectName("modeDescriptionLabel")
        self.mode_description.setWordWrap(True)
        self.update_mode_description(self.extraction_mode.currentText())
        self.extraction_mode.currentTextChanged.connect(self.update_mode_description)
        
        format_layout = QHBoxLayout()
        format_label = QLabel("Format:")
        self.output_format = QComboBox()
        self.output_format.addItems(["TXT", "HTML", "Markdown", "DOCX"])
        self.output_format.setToolTip("Select the output format for extracted text")
        format_layout.addWidget(format_label)
        format_layout.addWidget(self.output_format)
        
        options_layout.addLayout(mode_layout)
        options_layout.addWidget(self.mode_description)
        options_layout.addLayout(format_layout)
        options_group.setLayout(options_layout)
        left_layout.addWidget(options_group)
        
        splitter.addWidget(left_widget)
        
        # Right Widget Setup
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(10, 10, 10, 10)
        right_layout.setSpacing(10)
        
        self.tabs = QTabWidget()
        preview_tab = QWidget()
        preview_layout = QVBoxLayout(preview_tab)
        self.preview_widget = PreviewWidget()
        preview_layout.addWidget(self.preview_widget)
        
        text_tab = QWidget()
        text_layout = QVBoxLayout(text_tab)
        self.text_area = QTextEdit()
        self.text_area.setReadOnly(True)
        self.text_area.setToolTip("Extracted text will appear here")
        text_layout.addWidget(self.text_area)
        
        self.tabs.addTab(preview_tab, "Original PDF")
        self.tabs.addTab(text_tab, "Processed Text")
        right_layout.addWidget(self.tabs)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setToolTip("Extraction progress")
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setFormat("%p%")
        right_layout.addWidget(self.progress_bar)
        
        # Search Bar Setup
        search_widget = QWidget()
        search_layout = QHBoxLayout(search_widget)
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.setSpacing(5)  # Adjust spacing as needed
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search in text...")
        self.search_input.setToolTip("Enter text to search within the extracted content")
        self.search_input.textChanged.connect(self.search_text)
        self.search_input.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.search_input.setFixedHeight(30)  # Match the button height
        
        self.prev_btn = QPushButton()
        self.prev_btn.setObjectName("searchPrevButton")
        self.prev_btn.setFixedSize(30, 30)  # Smaller size
        self.prev_btn.setIcon(QIcon.fromTheme("go-previous"))
        self.prev_btn.setToolTip("Previous Match (Ctrl + Shift + P)")
        self.prev_btn.clicked.connect(self.search_previous)
        
        self.next_btn = QPushButton()
        self.next_btn.setObjectName("searchNextButton")
        self.next_btn.setFixedSize(30, 30)  # Smaller size
        self.next_btn.setIcon(QIcon.fromTheme("go-next"))
        self.next_btn.setToolTip("Next Match (Ctrl + Shift + N)")
        self.next_btn.clicked.connect(self.search_next)
        
        self.search_count = QLabel("0/0")
        self.search_count.setFixedWidth(50)
        self.search_count.setAlignment(Qt.AlignCenter)
        self.search_count.setFixedHeight(30)  # Match the buttons
        
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.prev_btn)
        search_layout.addWidget(self.next_btn)
        search_layout.addWidget(self.search_count)
        
        right_layout.addWidget(search_widget)
        
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        main_layout.addWidget(splitter)
        
        self.setCentralWidget(central_widget)

    def update_mode_description(self, mode):
        descriptions = {
            "Column-aware": "Best for documents with multiple columns.\nIntelligently detects and preserves column layout while maintaining proper reading order.",
            "Layout-preserved": "Maintains exact document formatting.\nPreserves spaces, indentation, and special characters.\nIdeal for forms and technical documents."
        }
        self.mode_description.setText(descriptions.get(mode, ""))

    def init_connections(self):
        self.files_list.itemClicked.connect(self.on_file_selected)

    def init_shortcuts(self):
        QShortcut(QKeySequence("Ctrl+P"), self, self.start_extraction)
        QShortcut(QKeySequence("Ctrl++"), self, self.preview_widget.zoom_in)
        QShortcut(QKeySequence("Ctrl+-"), self, self.preview_widget.zoom_out)
        QShortcut(QKeySequence("Ctrl+F"), self, self.focus_search)
        QShortcut(QKeySequence("Ctrl+C"), self, self.copy_text)

    def update_icon_colors(self):
        icon_color = QColor(self.theme_manager.custom_colors['icon_dark'] if self.theme_manager.dark_theme else self.theme_manager.custom_colors['icon_light'])
        for action in self.toolbar.actions():
            if not action.isSeparator():
                icon = action.icon()
                if not icon.isNull():
                    pixmap = icon.pixmap(QSize(16, 16))
                    if not pixmap.isNull():
                        painter = QPainter(pixmap)
                        painter.setCompositionMode(QPainter.CompositionMode_SourceIn)
                        painter.fillRect(pixmap.rect(), icon_color)
                        painter.end()
                        action.setIcon(QIcon(pixmap))

    def select_pdf(self):
        try:
            file_dialog = QFileDialog(self)
            file_dialog.setNameFilter("PDF Files (*.pdf)")
            file_dialog.setFileMode(QFileDialog.ExistingFiles)
            file_dialog.setWindowTitle("Select PDF File(s)")
            if file_dialog.exec():
                selected_files = file_dialog.selectedFiles()
                if selected_files:
                    for file in selected_files:
                        if file not in self.pdf_paths:
                            self.pdf_paths.append(file)
                            file_name = os.path.basename(file)
                            item = QListWidgetItem(file_name)
                            item.setToolTip(file)
                            self.files_list.addItem(item)
                    self.files_list.setCurrentRow(self.files_list.count() - len(selected_files))
                    self.preview_widget.set_document(self.pdf_paths[-1])
                    self.check_ready_to_extract()
                    self.show_toast("PDF(s) added successfully")
        except Exception as e:
            ErrorHandler.show_error(f"Error selecting PDF(s): {str(e)}", "Selection Error", self)

    def select_output_folder(self):
        try:
            folder_dialog = QFileDialog(self)
            folder_dialog.setFileMode(QFileDialog.Directory)
            folder_dialog.setWindowTitle("Select Output Folder")
            if folder_dialog.exec():
                selected_folders = folder_dialog.selectedFiles()
                if selected_folders:
                    self.output_path = selected_folders[0]
                    self.check_ready_to_extract()
                    self.show_toast("Output folder selected")
        except Exception as e:
            ErrorHandler.show_error(f"Error selecting output folder: {str(e)}", "Selection Error", self)

    def check_ready_to_extract(self):
        is_ready = bool(self.pdf_paths and self.output_path)
        self.process_btn.setEnabled(is_ready)
        if is_ready:
            self.status_bar.showMessage("Ready to process")
        else:
            missing = []
            if not self.pdf_paths:
                missing.append("PDF file(s)")
            if not self.output_path:
                missing.append("output folder")
            self.status_bar.showMessage(f"Missing: {', '.join(missing)}")

    def start_extraction(self):
        if not self.pdf_paths or not self.output_path:
            ErrorHandler.show_warning("Please select both PDF file(s) and output folder.", "Invalid Selection", self)
            return
        self.set_ui_enabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.thread = ExtractionThread(
            self.pdf_paths,
            self.output_path,
            self.extraction_mode.currentText(),
            self.output_format.currentText()
        )
        self.thread.progress.connect(self.update_progress)
        self.thread.finished.connect(self.extraction_finished)
        self.thread.error.connect(self.handle_error)
        self.thread.preview_ready.connect(self.preview_widget.load_current_page)
        self.thread.toast.connect(self.show_toast)
        self.thread.extracted_text.connect(self.store_extracted_text)
        self.thread.start()

    def set_ui_enabled(self, enabled):
        self.process_btn.setEnabled(enabled and bool(self.pdf_paths and self.output_path))
        self.extraction_mode.setEnabled(enabled)
        self.output_format.setEnabled(enabled)
        self.files_list.setEnabled(enabled)

    def update_progress(self, percentage, message):
        self.progress_bar.setValue(percentage)
        self.status_bar.showMessage(message)

    def extraction_finished(self, output_file):
        self.set_ui_enabled(True)
        self.progress_bar.setVisible(False)
        self.status_bar.showMessage(f"Extraction complete: {output_file}")
        self.show_toast("Extraction completed successfully")

    def handle_error(self, error_message):
        self.set_ui_enabled(True)
        self.progress_bar.setVisible(False)
        self.status_bar.showMessage("Extraction failed")
        ErrorHandler.show_error(f"An error occurred during extraction:\n{error_message}", "Extraction Error", self)
        self.show_toast("Extraction failed", error=True)

    def search_text(self):
        search_term = self.search_input.text().strip()
        if not search_term:
            self.clear_search_highlights()
            self.search_count.setText("0/0")
            return
        cursor = self.text_area.textCursor()
        highlight_format = QTextCursor().charFormat()
        highlight_format.setBackground(QColor("#4DA6FF"))
        cursor.movePosition(QTextCursor.Start)
        self.text_area.setTextCursor(cursor)
        self.clear_search_highlights()
        self.search_positions = []
        while True:
            cursor = self.text_area.document().find(search_term, cursor)
            if cursor.isNull():
                break
            self.search_positions.append((cursor.selectionStart(), cursor.selectionEnd()))
            cursor.mergeCharFormat(highlight_format)
        total_matches = len(self.search_positions)
        if total_matches > 0:
            self.current_search_index = 0
            self.search_count.setText(f"1/{total_matches}")
            self.navigate_to_match(0)
        else:
            self.search_count.setText("0/0")

    def search_next(self):
        if not self.search_positions:
            return
        self.current_search_index = (self.current_search_index + 1) % len(self.search_positions)
        self.navigate_to_match(self.current_search_index)

    def search_previous(self):
        if not self.search_positions:
            return
        self.current_search_index = (self.current_search_index - 1) % len(self.search_positions)
        self.navigate_to_match(self.current_search_index)

    def clear_search_highlights(self):
        cursor = self.text_area.textCursor()
        cursor.select(QTextCursor.Document)
        format_orig = QTextCursor().charFormat()
        format_orig.setBackground(Qt.transparent)
        cursor.mergeCharFormat(format_orig)

    def navigate_to_match(self, index):
        if 0 <= index < len(self.search_positions):
            start, end = self.search_positions[index]
            cursor = self.text_area.textCursor()
            cursor.setPosition(start)
            cursor.setPosition(end, QTextCursor.KeepAnchor)
            self.text_area.setTextCursor(cursor)
            self.search_count.setText(f"{index + 1}/{len(self.search_positions)}")

    def copy_text(self):
        cursor = self.text_area.textCursor()
        if cursor.hasSelection():
            QApplication.clipboard().setText(cursor.selectedText())
            self.show_toast("Text copied to clipboard")

    def focus_search(self):
        self.search_input.setFocus()
        self.search_input.selectAll()

    def show_about(self):
        QMessageBox.about(
            self,
            "About PDF Text Extractor",
            "PDF Text Extractor\n\nA powerful tool for extracting and processing text from PDF files.\nFeatures intelligent column detection and layout preservation.\n\nVersion 2.0\n© 2024 All rights reserved."
        )

    def show_help(self):
        help_dialog = HelpDialog(self)
        help_dialog.exec()

    def show_toast(self, message, error=False):
        toast = ToastNotification(message, self)
        if error:
            toast.setStyleSheet("""
                QLabel {
                    background-color: rgba(255, 0, 0, 180);
                    color: white;
                    padding: 10px;
                    border-radius: 5px;
                }
            """)
        toast.move(
            self.geometry().center().x() - toast.width() // 2,
            self.geometry().height() - 100
        )
        toast.show_notification()

    def get_current_theme_color(self):
        return self.theme_manager.custom_colors['background'] if self.current_theme == "dark" else "#FFFFFF"

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(url.toLocalFile().endswith('.pdf') for url in urls):
                event.acceptProposedAction()
            else:
                event.ignore()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        added = False
        for url in urls:
            file_path = url.toLocalFile()
            if file_path.endswith('.pdf') and file_path not in self.pdf_paths:
                self.pdf_paths.append(file_path)
                file_name = os.path.basename(file_path)
                item = QListWidgetItem(file_name)
                item.setToolTip(file_path)
                self.files_list.addItem(item)
                added = True
        if added:
            self.files_list.setCurrentRow(self.files_list.count() - 1)
            last_pdf = self.pdf_paths[-1]
            self.preview_widget.set_document(last_pdf)
            self.check_ready_to_extract()
            self.show_toast("PDF(s) added via drag-and-drop")

    def on_file_selected(self, item):
        pdf_path = item.toolTip()
        self.preview_widget.set_document(pdf_path)
        extracted_text = self.extracted_texts.get(pdf_path, "")
        self.text_area.setPlainText(extracted_text)

    def store_extracted_text(self, pdf_path, text):
        self.extracted_texts[pdf_path] = text

# Main function to run the application
def main():
    setup_logging()
    app = QApplication(sys.argv)
    app.setApplicationName("PDF Text Extractor")
    try:
        window = MainWindow()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        logging.exception("A critical error occurred:")

if __name__ == "__main__":
    main()