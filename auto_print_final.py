import sys
import os
import win32print
import win32api
from pathlib import Path
from datetime import datetime
import json
import time
import subprocess
import threading
import psutil
import tempfile
import winreg

from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *

# Используем PyMuPDF для работы с PDF
try:
    import fitz  # PyMuPDF
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False
    print("⚠️ PyMuPDF nicht installiert.")


class AutoPrintTool(QMainWindow):
    status_signal = Signal(str)
    printing_done_signal = Signal()
    log_print_signal = Signal(str)
    
    def __init__(self):
        super().__init__()
        self.current_file = None
        self.print_copies = 1
        self.config_file = Path("autoprint_config.json")
        self.printing_in_progress = False
        self.files_directory = "W:\\live\\Buttons"
        
        if not os.path.exists(self.files_directory):
            self.files_directory = str(Path.home())
        
        self.setup_ui()
        self.load_printers()
        self.load_config()
        
        self.status_signal.connect(self.update_status)
        self.printing_done_signal.connect(self.on_printing_done)
        self.log_print_signal.connect(self.do_log_print)
        
        print("AutoPrintTool gestartet")
    
    def setup_ui(self):
        self.setWindowTitle("🖨️ AutoPrintTool - Automatisches Drucksystem")
        self.setGeometry(100, 100, 900, 700)
        
        central = QWidget()
        central.setObjectName("central_widget")
        self.setCentralWidget(central)
        
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(25, 25, 25, 25)
        
        # Header
        header = QLabel("🖨️ AutoPrintTool")
        header.setAlignment(Qt.AlignCenter)
        font = header.font()
        font.setPointSize(24)
        font.setBold(True)
        header.setFont(font)
        header.setStyleSheet("color: #2c3e50; margin-bottom: 5px;")
        
        subheader = QLabel("Datei ziehen → Loslassen → Automatischer Druck")
        subheader.setAlignment(Qt.AlignCenter)
        subheader.setStyleSheet("""
            color: #7f8c8d; 
            font-size: 12px; 
            margin-bottom: 15px;
        """)
        
        main_layout.addWidget(header)
        main_layout.addWidget(subheader)
        
        # Превью и зона загрузки
        preview_container = QWidget()
        preview_layout = QHBoxLayout(preview_container)
        preview_layout.setContentsMargins(0, 0, 0, 0)
        
        # Превью
        self.preview_label = QLabel()
        self.preview_label.setObjectName("preview_label")
        self.preview_label.setFixedSize(220, 320)
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setStyleSheet("""
            QLabel#preview_label {
                border: 2px solid #dee2e6;
                border-radius: 8px;
                background: #f8f9fa;
                font-size: 14px;
                color: #6c757d;
            }
        """)
        self.preview_label.setText("Vorschau\nwird hier\nangezeigt")
        
        # Зона загрузки
        self.drop_zone = QLabel()
        self.drop_zone.setObjectName("drop_zone")
        self.drop_zone.setAlignment(Qt.AlignCenter)
        self.drop_zone.setText(
            "📁<br><br>"
            "<span style='font-size: 16px; font-weight: bold;'>Datei hier ablegen</span><br>"
            "<span style='color: #6c757d;'>oder klicken zum Auswählen</span><br><br>"
            "<small>Unterstützt: PDF, JPG, PNG, BMP</small><br>"
            "<small>PDF: Adobe Color Profiles werden verwendet</small>"
        )
        self.drop_zone.setAcceptDrops(True)
        
        preview_layout.addWidget(self.preview_label)
        preview_layout.addWidget(self.drop_zone, 1)
        
        main_layout.addWidget(preview_container)
        
        # Группа принтера с информацией
        printer_group = QGroupBox("📋 Drucker & Einstellungen")
        printer_group.setStyleSheet("""
            QGroupBox {
                font-size: 13px;
                font-weight: bold;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                margin-top: 8px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 6px 0 6px;
            }
        """)
        
        printer_layout = QVBoxLayout()
        
        # Выбор принтера
        printer_selection_layout = QHBoxLayout()
        
        self.printer_combo = QComboBox()
        self.printer_combo.setMinimumHeight(32)
        self.printer_combo.currentTextChanged.connect(self.update_printer_info_display)
        
        self.btn_refresh = QPushButton("🔄")
        self.btn_refresh.setFixedSize(32, 32)
        self.btn_refresh.setToolTip("Druckerliste aktualisieren")
        self.btn_refresh.clicked.connect(self.load_printers)
        
        printer_selection_layout.addWidget(self.printer_combo)
        printer_selection_layout.addWidget(self.btn_refresh)
        
        # Информация о принтере
        self.printer_info_display = QLabel("Wählen Sie einen Drucker")
        self.printer_info_display.setStyleSheet("""
            QLabel {
                color: #495057;
                font-size: 11px;
                padding: 5px;
                background: #f8f9fa;
                border-radius: 4px;
                border-left: 3px solid #667eea;
                margin-top: 3px;
            }
        """)
        self.printer_info_display.setWordWrap(True)
        
        printer_layout.addLayout(printer_selection_layout)
        printer_layout.addWidget(self.printer_info_display)
        
        printer_group.setLayout(printer_layout)
        main_layout.addWidget(printer_group)
        
        # Настройки копий
        copies_widget = QWidget()
        copies_layout = QHBoxLayout(copies_widget)
        copies_layout.setContentsMargins(0, 8, 0, 8)
        
        self.copy_label = QLabel("🔢 Anzahl der Kopien:")
        self.copy_label.setStyleSheet("font-size: 13px; font-weight: bold;")
        
        self.copy_spinbox = QSpinBox()
        self.copy_spinbox.setMinimum(1)
        self.copy_spinbox.setMaximum(9999)
        self.copy_spinbox.setValue(self.print_copies)
        self.copy_spinbox.setFixedWidth(90)
        self.copy_spinbox.valueChanged.connect(self.update_copy_count)
        
        self.btn_apply_copies = QPushButton("Übernehmen")
        self.btn_apply_copies.clicked.connect(self.apply_copy_settings)
        self.btn_apply_copies.setFixedWidth(90)
        
        copies_layout.addWidget(self.copy_label)
        copies_layout.addWidget(self.copy_spinbox)
        copies_layout.addWidget(self.btn_apply_copies)
        copies_layout.addStretch()
        
        main_layout.addWidget(copies_widget)
        
        # Кнопки управления
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(8)
        
        self.btn_save_config = QPushButton("💾 Standard")
        self.btn_save_config.setToolTip("Speichert die aktuellen Einstellungen")
        self.btn_save_config.clicked.connect(self.save_printer_config)
        self.btn_save_config.setFixedWidth(120)
        
        self.btn_print = QPushButton("🚀 DRUCKEN")
        self.btn_print.setObjectName("print_button")
        self.btn_print.setMinimumHeight(40)
        self.btn_print.setEnabled(False)
        self.btn_print.clicked.connect(self.start_printing)
        
        self.btn_reset = QPushButton("🔄 Zurücksetzen")
        self.btn_reset.clicked.connect(self.reset_ui)
        self.btn_reset.setFixedWidth(110)
        
        buttons_layout.addWidget(self.btn_save_config)
        buttons_layout.addWidget(self.btn_print, 1)
        buttons_layout.addWidget(self.btn_reset)
        
        main_layout.addLayout(buttons_layout)
        
        # Статус
        self.status_label = QLabel("🔵 Bereit. Datei per Drag & Drop hinzufügen.")
        self.status_label.setObjectName("status_label")
        
        main_layout.addWidget(self.status_label)
        
        main_layout.addStretch()
        
        self.setStyleSheet("""
            QMainWindow {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                          stop:0 #667eea, stop:1 #764ba2);
            }
            
            QWidget#central_widget {
                background: white;
                border-radius: 10px;
            }
            
            QLabel#drop_zone {
                border: 3px dashed #667eea;
                border-radius: 10px;
                background: #f8f9ff;
                padding: 20px;
                font-size: 12px;
                color: #495057;
                min-height: 300px;
            }
            
            QLabel#drop_zone:hover {
                background: #eef1ff;
                border-color: #764ba2;
            }
            
            QPushButton#print_button {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                          stop:0 #28a745, stop:1 #20c997);
                color: white;
                font-size: 14px;
                font-weight: bold;
                border-radius: 6px;
                border: none;
                padding: 8px 16px;
            }
            
            QPushButton#print_button:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                          stop:0 #218838, stop:1 #1e7e34);
            }
            
            QPushButton#print_button:disabled {
                background: #adb5bd;
            }
            
            QLabel#status_label {
                background: #f8f9fa;
                border-radius: 5px;
                padding: 8px;
                border-left: 3px solid #17a2b8;
                font-weight: bold;
                color: #0c5460;
                font-size: 12px;
            }
            
            QSpinBox {
                padding: 5px 8px;
                border: 2px solid #667eea;
                border-radius: 4px;
                font-size: 12px;
                font-weight: bold;
            }
            
            QSpinBox:hover {
                border-color: #764ba2;
            }
            
            QPushButton {
                background: #667eea;
                color: white;
                border: none;
                padding: 6px 12px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 12px;
            }
            
            QPushButton:hover {
                background: #764ba2;
            }
        """)
        
        self.drop_zone.mousePressEvent = self.select_file
        self.setAcceptDrops(True)
    
    def update_status(self, message):
        self.status_label.setText(message)
        QApplication.processEvents()
    
    def on_printing_done(self):
        self.btn_print.setEnabled(True)
        self.printing_in_progress = False
        self.btn_print.setText("🚀 DRUCKEN")
    
    def do_log_print(self, printer_name):
        self.log_print(printer_name)
    
    def load_config(self):
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    
                saved_printer = config.get('default_printer')
                if saved_printer:
                    index = self.printer_combo.findText(saved_printer)
                    if index >= 0:
                        self.printer_combo.setCurrentIndex(index)
                        self.update_printer_info_display()
                    
                saved_copies = config.get('default_copies', 1)
                self.copy_spinbox.setValue(saved_copies)
                self.print_copies = saved_copies
                
                saved_dir = config.get('files_directory')
                if saved_dir and os.path.exists(saved_dir):
                    self.files_directory = saved_dir
                
                self.status_label.setText(f"💾 Konfiguration geladen")
                
        except Exception as e:
            print(f"Fehler beim Laden der Konfiguration: {e}")
    
    def save_printer_config(self):
        printer = self.printer_combo.currentText()
        if not printer:
            QMessageBox.warning(self, "Warnung", "Bitte wählen Sie zuerst einen Drucker aus")
            return
            
        try:
            config = {
                'default_printer': printer,
                'default_copies': self.print_copies,
                'last_saved': datetime.now().isoformat(),
                'files_directory': self.files_directory
            }
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            
            self.status_label.setText(f"✅ Standard gespeichert: {printer}")
            
        except Exception as e:
            self.status_label.setText("❌ Fehler beim Speichern")
    
    def update_printer_info_display(self):
        """Отображает информацию о принтере"""
        printer_name = self.printer_combo.currentText()
        if not printer_name:
            self.printer_info_display.setText("Wählen Sie einen Drucker")
            return
        
        try:
            handle = win32print.OpenPrinter(printer_name)
            printer_info = win32print.GetPrinter(handle, 2)
            win32print.ClosePrinter(handle)
            
            info_parts = []
            
            # Статус
            status = printer_info['Status']
            status_text = {
                0: "✅ Bereit",
                1: "⏸️ Pausiert",
                2: "❌ Fehler",
                4: "🖨️ Druckt",
                5: "🔌 Offline"
            }.get(status, f"Status {status}")
            info_parts.append(status_text)
            
            # Настройки из DEVMODE
            if 'pDevMode' in printer_info and printer_info['pDevMode']:
                devmode = printer_info['pDevMode']
                
                # Размер бумаги
                if hasattr(devmode, 'PaperSize'):
                    paper_sizes = {
                        1: "Letter",
                        5: "Legal",
                        8: "A3",
                        9: "A4",
                        11: "A5",
                        80: "Custom"
                    }
                    paper = paper_sizes.get(devmode.PaperSize, f"Size {devmode.PaperSize}")
                    info_parts.append(f"📄 {paper}")
                
                # Ориентация
                if hasattr(devmode, 'Orientation'):
                    orientation = "H" if devmode.Orientation == 1 else "Q"
                    info_parts.append(f"🔄 {orientation}")
                
                # Разрешение
                if hasattr(devmode, 'PrintQuality') and devmode.PrintQuality > 0:
                    info_parts.append(f"⭐ {devmode.PrintQuality}dpi")
                
                # Цвет
                if hasattr(devmode, 'Color'):
                    color = "Farb" if devmode.Color == 2 else "S/W"
                    info_parts.append(f"🎨 {color}")
            
            self.printer_info_display.setText(" | ".join(info_parts))
            
        except Exception as e:
            print(f"Fehler beim Abrufen der Druckerinfo: {e}")
            self.printer_info_display.setText(f"✅ {printer_name}")
    
    def generate_preview(self, file_path):
        """Создает превью файла"""
        try:
            file_ext = Path(file_path).suffix.lower()
            preview_size = QSize(210, 300)
            
            if file_ext in ['.jpg', '.jpeg', '.png', '.bmp']:
                pixmap = QPixmap(file_path)
                if not pixmap.isNull():
                    pixmap = pixmap.scaled(preview_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    self.preview_label.setPixmap(pixmap)
                else:
                    self.set_preview_icon(file_ext)
                    
            elif file_ext == '.pdf' and PYPDF_AVAILABLE:
                try:
                    doc = fitz.open(file_path)
                    page = doc.load_page(0)
                    pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                    
                    img_data = pix.tobytes("ppm")
                    qimage = QImage()
                    qimage.loadFromData(img_data)
                    
                    pixmap = QPixmap.fromImage(qimage)
                    if not pixmap.isNull():
                        pixmap = pixmap.scaled(preview_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                        self.preview_label.setPixmap(pixmap)
                    else:
                        self.set_preview_icon(file_ext)
                    
                    doc.close()
                        
                except Exception as e:
                    print(f"PDF Vorschau Fehler: {e}")
                    self.set_preview_icon(file_ext)
                    
            else:
                self.set_preview_icon(file_ext)
                
        except Exception as e:
            print(f"Vorschau Fehler: {e}")
            self.set_preview_icon('.unknown')
    
    def set_preview_icon(self, file_ext):
        """Устанавливает иконку вместо превью"""
        icons = {
            '.pdf': '📄 PDF',
            '.jpg': '🖼️ JPG',
            '.jpeg': '🖼️ JPEG',
            '.png': '🖼️ PNG',
            '.bmp': '🖼️ BMP',
        }
        icon_text = icons.get(file_ext, '📁 Datei')
        self.preview_label.setText(icon_text)
    
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drop_zone.setStyleSheet("""
                QLabel#drop_zone {
                    border: 3px solid #28a745;
                    border-radius: 10px;
                    background: #d4edda;
                    padding: 20px;
                    font-size: 12px;
                    color: #155724;
                    min-height: 300px;
                }
            """)
    
    def dragLeaveEvent(self, event):
        self.drop_zone.setStyleSheet("""
            QLabel#drop_zone {
                border: 3px dashed #667eea;
                border-radius: 10px;
                background: #f8f9ff;
                padding: 20px;
                font-size: 12px;
                color: #495057;
                min-height: 300px;
            }
        """)
    
    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            self.load_file(file_path)
        
        self.drop_zone.setStyleSheet("""
            QLabel#drop_zone {
                border: 3px dashed #667eea;
                border-radius: 10px;
                background: #f8f9ff;
                padding: 20px;
                font-size: 12px;
                color: #495057;
                min-height: 300px;
            }
        """)
    
    def select_file(self, event):
        if os.path.exists(self.files_directory):
            start_dir = self.files_directory
        else:
            start_dir = str(Path.home())
        
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Datei auswählen",
            start_dir,
            "Unterstützte Dateien (*.pdf *.jpg *.jpeg *.png *.bmp)"
        )
        
        if file_path:
            self.files_directory = os.path.dirname(file_path)
            self.load_file(file_path)
    
    def load_file(self, file_path):
        if os.path.exists(file_path):
            self.current_file = file_path
            self.generate_preview(file_path)
            
            filename = os.path.basename(file_path)
            self.status_label.setText(f"✅ Datei geladen: {filename}")
            self.btn_print.setEnabled(True)
            self.btn_print.setText(f"🚀 {self.print_copies} KOPIE(N) DRUCKEN")
    
    def load_printers(self):
        try:
            printers = win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            )
            
            default_printer = win32print.GetDefaultPrinter()
            self.printer_combo.clear()
            
            for printer in printers:
                name = printer[2]
                self.printer_combo.addItem(name)
                if name == default_printer:
                    self.printer_combo.setCurrentText(name)
            
            if len(printers) > 0:
                self.status_label.setText(f"📋 {len(printers)} Drucker geladen")
                self.update_printer_info_display()
            else:
                self.status_label.setText("⚠ Keine Drucker gefunden")
                
        except Exception as e:
            print(f"Fehler beim Laden der Drucker: {e}")
            self.status_label.setText("⚠ Fehler beim Laden der Drucker")
    
    def update_copy_count(self, value):
        self.print_copies = value
        if self.current_file:
            self.btn_print.setText(f"🚀 {value} KOPIE(N) DRUCKEN")
    
    def apply_copy_settings(self):
        self.status_label.setText(f"✅ Kopien: {self.print_copies} eingestellt")
        if self.current_file:
            self.btn_print.setText(f"🚀 {self.print_copies} KOPIE(N) DRUCKEN")
    
    def start_printing(self):
        if not self.current_file or not os.path.exists(self.current_file):
            self.status_label.setText("❌ Keine Datei ausgewählt")
            return
        
        printer = self.printer_combo.currentText()
        if not printer:
            self.status_label.setText("❌ Bitte Drucker auswählen")
            return
        
        if self.printing_in_progress:
            self.status_label.setText("⚠️ Druckvorgang läuft bereits")
            return
        
        # Подтверждение для большого количества
        if self.print_copies > 50:
            reply = QMessageBox.question(
                self,
                "Bestätigung",
                f"Möchten Sie wirklich {self.print_copies} Kopien drucken?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.No:
                return
        
        self.printing_in_progress = True
        self.btn_print.setEnabled(False)
        self.btn_print.setText("🖨️ DRUCKE...")
        
        thread = threading.Thread(
            target=self.print_file_thread,
            args=(self.current_file, printer, self.print_copies),
            daemon=True
        )
        thread.start()
    
    def print_file_thread(self, file_path, printer_name, copies):
        """Основной поток печати"""
        try:
            file_ext = Path(file_path).suffix.lower()
            
            if file_ext == '.pdf':
                # Для PDF - используем УЛУЧШЕННУЮ печать через Adobe
                self.status_signal.emit(f"🚀 Starte Adobe-Druck für {copies} Kopie(n)...")
                success = self.print_pdf_adobe_simple(file_path, printer_name, copies)
                
                if not success:
                    # Резервный метод - Windows печать
                    self.status_signal.emit("⚠️ Adobe nicht verfügbar, verwende Windows-Druck...")
                    self.print_with_windows(file_path, printer_name, copies)
            else:
                # Для изображений - Windows печать
                self.status_signal.emit(f"🚀 Starte Windows-Druck für {copies} Kopie(n)...")
                self.print_with_windows(file_path, printer_name, copies)
            
            self.status_signal.emit(f"✅ {copies} Kopie(n) gesendet an {printer_name}")
            self.log_print_signal.emit(printer_name)
            
            time.sleep(1)
            QTimer.singleShot(100, self.reset_ui_after_print)
            
        except Exception as e:
            error_msg = f"❌ Druckfehler: {str(e)}"
            self.status_signal.emit(error_msg)
            print(f"Fehler: {e}")
        finally:
            self.printing_done_signal.emit()
    
    def print_pdf_adobe_simple(self, file_path, printer_name, copies):
        """ПРОСТОЙ и надежный метод печати через Adobe"""
        try:
            # Находим Adobe Reader
            adobe_path = self.find_adobe_reader()
            if not adobe_path:
                return False
            
            print(f"Adobe gefunden: {adobe_path}")
            
            # Закрываем все предыдущие Adobe
            self.force_kill_adobe()
            time.sleep(0.5)
            
            # Используем параметры командной строки Adobe:
            # /t <filename> <printername> - печатать и закрыть
            # /p <filename> - печатать в портретном режиме (необязательно)
            
            base_cmd = f'"{adobe_path}" /t "{file_path}" "{printer_name}"'
            print(f"Adobe Befehl: {base_cmd}")
            
            # Запускаем процессы печати
            success_count = 0
            
            for i in range(copies):
                if copies > 1 and i % 10 == 0:
                    self.status_signal.emit(f"📤 Sende Kopie {i+1}/{copies}")
                
                try:
                    # Настройки для скрытого запуска
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE
                    
                    # Запускаем Adobe для печати
                    process = subprocess.Popen(
                        base_cmd,
                        shell=True,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        startupinfo=startupinfo,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    
                    success_count += 1
                    
                    # Минимальная задержка между запусками
                    time.sleep(0.5)
                    
                except Exception as e:
                    print(f"Fehler bei Kopie {i+1}: {e}")
                    continue
            
            # Даем время на обработку
            time.sleep(2)
            
            # Закрываем Adobe
            self.force_kill_adobe()
            
            return success_count > 0
            
        except Exception as e:
            print(f"Adobe Druckfehler: {e}")
            self.force_kill_adobe()
            return False
    
    def print_with_windows(self, file_path, printer_name, copies):
        """Печать через Windows ShellExecute"""
        original_printer = win32print.GetDefaultPrinter()
        
        try:
            win32print.SetDefaultPrinter(printer_name)
        except:
            pass
        
        try:
            # Быстрая отправка всех копий
            for i in range(copies):
                if copies > 1 and i % 20 == 0:
                    self.status_signal.emit(f"📤 Sende Kopie {i+1}/{copies}")
                
                win32api.ShellExecute(
                    0,
                    "print",
                    file_path,
                    None,
                    ".",
                    0
                )
                
                # Минимальная пауза
                if i < copies - 1:
                    time.sleep(0.02)
                
        finally:
            # Восстанавливаем принтер
            if original_printer:
                try:
                    win32print.SetDefaultPrinter(original_printer)
                except:
                    pass
    
    def find_adobe_reader(self):
        """Находит Adobe Reader в системе"""
        # Стандартные пути
        paths = [
            r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files\Adobe\Acrobat Reader\Acrobat Reader.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat Reader\Acrobat Reader.exe",
        ]
        
        for path in paths:
            if os.path.exists(path):
                return path
        
        # Ищем через реестр
        try:
            # 64-bit
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe")
            path, _ = winreg.QueryValueEx(key, "")
            winreg.CloseKey(key)
            if os.path.exists(path):
                return path
        except:
            pass
        
        try:
            # 32-bit
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe")
            path, _ = winreg.QueryValueEx(key, "")
            winreg.CloseKey(key)
            if os.path.exists(path):
                return path
        except:
            pass
        
        return None
    
    def force_kill_adobe(self):
        """Закрывает Adobe Reader"""
        try:
            # Основные процессы Adobe
            processes = ["AcroRd32.exe", "Acrobat.exe", "AcroDist.exe"]
            
            for proc in processes:
                subprocess.run(
                    f"taskkill /F /IM {proc} /T 2>nul",
                    shell=True,
                    capture_output=True,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            
            time.sleep(0.3)
            
        except Exception as e:
            print(f"Fehler beim Beenden von Adobe: {e}")
    
    def log_print(self, printer_name):
        """Логирование печати"""
        try:
            log_file = Path("print_log.txt")
            with open(log_file, 'a', encoding='utf-8') as f:
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                filename = os.path.basename(self.current_file) if self.current_file else "Unbekannt"
                f.write(f"{timestamp} | {filename} | {printer_name} | {self.print_copies} Kopien\n")
        except:
            pass
    
    def reset_ui_after_print(self):
        """Сброс после печати"""
        self.current_file = None
        self.preview_label.clear()
        self.set_preview_icon('.unknown')
        self.btn_print.setEnabled(False)
        self.btn_print.setText("🚀 DRUCKEN")
        self.printing_in_progress = False
        self.status_label.setText("✅ Fertig. Neue Datei wählen.")
    
    def reset_ui(self):
        """Ручной сброс"""
        self.current_file = None
        self.preview_label.clear()
        self.set_preview_icon('.unknown')
        self.btn_print.setEnabled(False)
        self.btn_print.setText("🚀 DRUCKEN")
        self.printing_in_progress = False
        self.status_label.setText("🔵 Bereit für nächste Datei")


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = AutoPrintTool()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
