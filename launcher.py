# pip install pyqt6
import sys
import os
import subprocess
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, 
                            QTextEdit, QMessageBox, QLabel, QLineEdit,
                            QHBoxLayout, QSpacerItem, QSizePolicy, QProgressBar, 
                            QStyleOptionProgressBar, QDialog, QFormLayout)
from PyQt6.QtGui import (QPalette, QColor, QFont, QPainter, 
                         QLinearGradient, QBrush, QPen, QPainterPath)
from PyQt6.QtCore import (Qt, QSettings, QTimer, QPoint, QPropertyAnimation, 
                          QEasingCurve, pyqtProperty, QRect)


def recurso_path(ruta_relativa):
    """Devuelve la ruta absoluta compatible con PyInstaller"""
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, ruta_relativa)

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.settings = QSettings("MiEmpresa", "ProgramLauncher")
        self.setWindowTitle("Configuración")
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint)
        self.setModal(True)
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout()
        form_layout = QFormLayout()
        
        # Campo para ID
        self.id_input = QLineEdit(self)
        self.id_input.setPlaceholderText("Ingrese su ID")
        saved_id = self.settings.value("user_id", "")
        self.id_input.setText(saved_id)
        
        # Campo para contraseña
        self.password_input = QLineEdit(self)
        self.password_input.setPlaceholderText("Ingrese su contraseña")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        saved_password = self.settings.value("user_password", "")
        self.password_input.setText(saved_password)
        
        # Campo para iniciales del analista
        self.initials_input = QLineEdit(self)
        self.initials_input.setPlaceholderText("Ej: CJ")
        self.initials_input.setMaxLength(3)
        saved_initials = self.settings.value("user_initials", "")
        self.initials_input.setText(saved_initials)
        
        form_layout.addRow("ID:", self.id_input)
        form_layout.addRow("Contraseña:", self.password_input)
        form_layout.addRow("Iniciales Analista:", self.initials_input)
        
        # Botones
        button_layout = QHBoxLayout()
        self.save_button = QPushButton("Guardar", self)
        self.save_button.clicked.connect(self.save_settings)
        self.cancel_button = QPushButton("Cancelar", self)
        self.cancel_button.clicked.connect(self.close)
        
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.cancel_button)
        
        layout.addLayout(form_layout)
        layout.addLayout(button_layout)
        self.setLayout(layout)
        
    def save_settings(self):
        """Guarda las configuraciones y cierra el diálogo"""
        self.settings.setValue("user_id", self.id_input.text())
        self.settings.setValue("user_password", self.password_input.text())
        self.settings.setValue("user_initials", self.initials_input.text().upper())
        self.accept()

class GlowProgressBar(QProgressBar):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_style = 0  # 0:Sakura, 1:Dark, 2:Dorado, 3:Rojo Vino
        self._glow_pos = -0.2
        self._setup_animation()

    def _setup_animation(self):
        self.animation = QPropertyAnimation(self, b"glow_pos")
        self.animation.setDuration(1500)
        self.animation.setLoopCount(-1)
        self.animation.setStartValue(-0.2)
        self.animation.setEndValue(1.2)
        self.animation.setEasingCurve(QEasingCurve.Type.Linear)
        self.animation.start()

    def setCurrentStyle(self, style):
        self._current_style = style
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        option = QStyleOptionProgressBar()
        self.initStyleOption(option)
        rect = option.rect
        
        # Colores de fondo para cada tema
        bg_colors = [
            QColor("#FFE5ED"),  # Sakura - Rosa claro
            QColor("#1A2A4A"),  # Dark - Azul oscuro
            QColor("#FFF4D5"),  # Dorado - Dorado claro
            QColor("#2A0A0A")   # Rojo Vino - Fondo oscuro
        ]
        
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QBrush(bg_colors[self._current_style]))
        painter.drawRoundedRect(rect, 5, 5)
        
        # Calcular progreso
        progress = max(self.minimum(), min(self.value(), self.maximum()))
        progress_width = int((progress - self.minimum()) / (self.maximum() - self.minimum()) * rect.width())
        
        if progress_width > 0:
            # Colores base para cada tema
            base_colors = [
                QColor("#FF6B8B"),  # Sakura - Rosa
                QColor("#4A90E2"),  # Dark - Azul
                QColor("#FFD700"),  # Dorado - Oro
                QColor("#8B0000")   # Rojo Vino - Rojo oscuro
            ]
            
            glow_colors = [
                QColor(255, 255, 255, 180),  # Sakura
                QColor(255, 255, 255, 150),  # Dark
                QColor(255, 255, 255, 200),  # Dorado (brillo más intenso)
                QColor(255, 200, 200, 150)   # Rojo Vino (brillo rojizo)
            ]
            
            # Dibujar progreso base
            progress_rect = QRect(rect)
            progress_rect.setWidth(progress_width)
            painter.setBrush(QBrush(base_colors[self._current_style]))
            painter.drawRoundedRect(progress_rect, 5, 5)
            
            # Efecto de brillo
            if progress_width > 20:
                glow_width = progress_width * 0.4
                glow_x = int(self._glow_pos * progress_width)
                
                gradient = QLinearGradient(glow_x - glow_width, 0, glow_x + glow_width, 0)
                gradient.setColorAt(0, Qt.GlobalColor.transparent)
                gradient.setColorAt(0.5, glow_colors[self._current_style])
                gradient.setColorAt(1, Qt.GlobalColor.transparent)
                
                glow_rect = QRect(int(glow_x - glow_width), 0, int(glow_width * 2), rect.height())
                painter.setBrush(QBrush(gradient))
                painter.setPen(Qt.PenStyle.NoPen)
                painter.drawRoundedRect(glow_rect.intersected(progress_rect), 5, 5)
        
        # Colores del borde
        border_colors = [
            QColor("#FF9EB5"),  # Sakura
            QColor("#4A90E2"),  # Dark
            QColor("#D4AF37"),  # Dorado (oro metálico)
            QColor("#6D071A")   # Rojo Vino (borde vino)
        ]
        
        painter.setPen(QPen(border_colors[self._current_style], 1.5))
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRoundedRect(rect, 5, 5)

    def getGlowPos(self):
        return self._glow_pos

    def setGlowPos(self, pos):
        self._glow_pos = pos
        self.update()

    glow_pos = pyqtProperty(float, getGlowPos, setGlowPos)

class TitleBar(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent_window = parent  
        self.setFixedHeight(35)
        self.mouse_pos = QPoint(0, 0)
        self.initUI()

    def initUI(self):
        self.layout = QHBoxLayout()
        self.layout.setContentsMargins(10, 0, 10, 0)
        self.layout.setSpacing(10)
        
        # Icono de la aplicación
        self.icon_label = QLabel(self)
        self.icon_label.setFixedSize(20, 20)
        self.layout.addWidget(self.icon_label)
        
        # Título de la ventana
        self.title_label = QLabel("Iris Automatizer", self)
        self.title_label.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        self.title_label.setStyleSheet("background: transparent; border: none;")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.layout.addWidget(self.title_label, stretch=1)
        
        # Espaciador
        self.layout.addSpacerItem(QSpacerItem(20, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        
        # Botones de control
        self.min_button = QPushButton("─", self)
        self.max_button = QPushButton("□", self)
        self.close_button = QPushButton("×", self)
        
        for btn in [self.min_button, self.max_button, self.close_button]:
            btn.setFixedSize(25, 25)
            btn.setFont(QFont("Arial", 12))
        
        self.min_button.clicked.connect(self.parent_window.showMinimized)
        self.max_button.clicked.connect(self.toggle_maximize)
        self.close_button.clicked.connect(self.parent_window.close)
        
        self.layout.addWidget(self.min_button)
        self.layout.addWidget(self.max_button)
        self.layout.addWidget(self.close_button)
        
        self.setLayout(self.layout)
    
    def toggle_maximize(self):
        if self.parent_window.isMaximized():
            self.parent_window.showNormal()
        else:
            self.parent_window.showMaximized()
    
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.mouse_pos = event.globalPosition().toPoint() - self.parent_window.pos()
    
    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.MouseButton.LeftButton:
            self.parent_window.move(event.globalPosition().toPoint() - self.mouse_pos)
    
    def paintEvent(self, event):
        painter = QPainter(self)
        
        # Configuraciones para cada tema
        themes = [
            {  # Sakura
                "gradient": ["#FFD6E0", "#FFF0F5"],
                "line_color": "#FF9EB5",
                "control_style": """
                    background-color: #FFC0CB;
                    color: #5A3A3A;
                    border: 1px solid #FF9EB5;
                """,
                "hover_style": """
                    background-color: #FFB6C1;
                    border: 1px solid #FF6B8B;
                """,
                "close_style": """
                    background-color: #FF6B8B;
                    color: white;
                    border: 1px solid #FF3D6D;
                """,
                "icon_style": "background-color: #FF9EB5;",
                "title_color": "#5A3A3A"
            },
            {  # Dark
                "gradient": ["#1A2A4A", "#2C3E66"],
                "line_color": "#4A90E2",
                "control_style": """
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                              stop:0 #2C3E66, stop:1 #1A2A4A);
                    color: #A8C4FF;
                    border: 1px solid #4A90E2;
                """,
                "hover_style": """
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                              stop:0 #3A4E76, stop:1 #2A3A5A);
                    border: 1px solid #6AA8FF;
                """,
                "close_style": """
                    background-color: #FF6B6B;
                    color: white;
                    border: 1px solid #FF3D3D;
                """,
                "icon_style": """
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                              stop:0 #4A90E2, stop:1 #6AA8FF);
                """,
                "title_color": "#A8C4FF"
            },
            {  # Dorado
                "gradient": ["#FFE8B0", "#FFF4D5"],
                "line_color": "#D4AF37",
                "control_style": """
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                              stop:0 #FFE8B0, stop:1 #FFD700);
                    color: #5A3A3A;
                    border: 1px solid #D4AF37;
                """,
                "hover_style": """
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                              stop:0 #FFE8C0, stop:1 #FFDF00);
                    border: 1px solid #FFC800;
                """,
                "close_style": """
                    background-color: #D4AF37;
                    color: white;
                    border: 1px solid #B8860B;
                """,
                "icon_style": """
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                              stop:0 #FFD700, stop:1 #D4AF37);
                """,
                "title_color": "#5A3A3A"
            },
            {  # Rojo Vino
                "gradient": ["#3A0A0A", "#5A1A1A"],
                "line_color": "#6D071A",
                "control_style": """
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                              stop:0 #5A1A1A, stop:1 #3A0A0A);
                    color: #E0C0C0;
                    border: 1px solid #6D071A;
                """,
                "hover_style": """
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                              stop:0 #6A2A2A, stop:1 #4A1A1A);
                    border: 1px solid #8D071A;
                """,
                "close_style": """
                    background-color: #8B0000;
                    color: white;
                    border: 1px solid #6D071A;
                """,
                "icon_style": """
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                              stop:0 #8B0000, stop:1 #6D071A);
                """,
                "title_color": "#E0C0C0"
            }
        ]
        
        theme = themes[self.parent_window.current_style]
        
        # Aplicar gradiente de fondo
        gradient = QLinearGradient(0, 0, self.width(), 0)
        gradient.setColorAt(0, QColor(theme["gradient"][0]))
        gradient.setColorAt(1, QColor(theme["gradient"][1]))
        painter.fillRect(self.rect(), QBrush(gradient))
        
        # Línea inferior
        pen = QPen(QColor(theme["line_color"]), 1.5)
        painter.setPen(pen)
        painter.drawLine(0, self.height()-1, self.width(), self.height()-1)
        
        # Estilos para los botones de control
        control_style = f"""
            QPushButton {{
                {theme["control_style"]}
                border-radius: 12px;
                font-size: 14px;
            }}
            QPushButton:hover {{
                {theme["hover_style"]}
            }}
        """
        
        close_style = f"""
            QPushButton {{
                {theme["close_style"]}
                border-radius: 12px;
                font-size: 14px;
            }}
            QPushButton:hover {{
                background-color: #FF3D3D;
            }}
        """
        
        # Aplicar estilos
        self.min_button.setStyleSheet(control_style)
        self.max_button.setStyleSheet(control_style)
        self.close_button.setStyleSheet(close_style)
        
        # Estilo del icono
        self.icon_label.setStyleSheet(f"""
            {theme["icon_style"]}
            border-radius: 10px;
        """)
        
        # Estilo del título
        self.title_label.setStyleSheet(f"""
            QLabel {{
                background: transparent;
                color: {theme["title_color"]};
                border: none;
                padding: 0px;
                margin: 0px;
            }}
        """)

class ProgramLauncher(QWidget):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("MiEmpresa", "ProgramLauncher")
        self.current_style = self.settings.value("style", 0, int) % 4  # Aseguramos que esté entre 0-3
        self.user_input_var = self.settings.value("user_initials", "")  # Cargar iniciales guardadas
        
        # Configuración de rutas
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.progress_dir = os.path.join(self.base_dir, "data", "cache")
        self.progress_file = os.path.join(self.progress_dir, "progress.txt")
        os.makedirs(self.progress_dir, exist_ok=True)
        
        self.current_process = None
        self.is_aborted = False
        
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.initUI()
        self.apply_style()
    
    def initUI(self):
        self.setWindowTitle("Iris Automatizer")
        self.setGeometry(100, 100, 450, 600)  # Aumenté un poco la altura
        
        self.main_layout = QVBoxLayout()
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        
        self.title_bar = TitleBar(self)
        self.main_layout.addWidget(self.title_bar)
        
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout()
        self.content_layout.setContentsMargins(15, 15, 15, 15)
        
        # Botón Extraer Datos OIRS
        btn_extractor = QPushButton("Extraer Datos OIRS", self)
        btn_extractor.setFont(QFont("Arial", 12))
        btn_extractor.setMinimumHeight(35)
        btn_extractor.clicked.connect(lambda: self.run_script("extractor.py"))
        self.content_layout.addWidget(btn_extractor)
        
        # Botón Crear Ficha Retiro
        btn_retiros = QPushButton("Crear Ficha Retiro", self)
        btn_retiros.setFont(QFont("Arial", 12))
        btn_retiros.setMinimumHeight(35)
        btn_retiros.clicked.connect(lambda: self.run_script("retiros.py"))
        self.content_layout.addWidget(btn_retiros)
        
        # Contenedor para botones de responder
        responder_container = QWidget()
        responder_layout = QHBoxLayout()
        responder_layout.setContentsMargins(0, 0, 0, 0)
        responder_layout.setSpacing(5)
        
        # Botón Responder OIRS
        btn_responder_oirs = QPushButton("Responder OIRS", self)
        btn_responder_oirs.setFont(QFont("Arial", 10))
        btn_responder_oirs.setMinimumHeight(35)
        btn_responder_oirs.clicked.connect(lambda: self.run_script("responderoirs.py"))
        
        # Botón Responder Ref y NC
        btn_responder_refync = QPushButton("Responder Ref y NC", self)
        btn_responder_refync.setFont(QFont("Arial", 10))
        btn_responder_refync.setMinimumHeight(35)
        btn_responder_refync.clicked.connect(lambda: self.run_script("responderoirsrefync.py"))
        
        responder_layout.addWidget(btn_responder_oirs)
        responder_layout.addWidget(btn_responder_refync)
        responder_container.setLayout(responder_layout)
        self.content_layout.addWidget(responder_container)
        
        # Botones pequeños (configuración y tema)
        small_buttons_layout = QHBoxLayout()
        
        self.btn_settings = QPushButton("⚙️", self)
        self.btn_settings.setFixedSize(30, 30)
        self.btn_settings.setToolTip("Configuración")
        self.btn_settings.clicked.connect(self.show_settings)
        
        theme_icons = ["🌸", "🌙", "🌟", "🍷"]
        self.btn_toggle_style = QPushButton(theme_icons[self.current_style], self)
        self.btn_toggle_style.setFixedSize(30, 30)
        self.btn_toggle_style.clicked.connect(self.toggle_style)
        
        small_buttons_layout.addWidget(self.btn_settings)
        small_buttons_layout.addStretch()
        small_buttons_layout.addWidget(self.btn_toggle_style)
        self.content_layout.addLayout(small_buttons_layout)
        
        # Área de texto principal
        self.text_edit = QTextEdit(self)
        self.text_edit.setPlaceholderText("Escribe aquí el texto a enviar a los scripts...")
        self.text_edit.setFont(QFont("Arial", 12))
        self.content_layout.addWidget(self.text_edit)
        
        # Botón Limpiar
        self.btn_clear = QPushButton("Limpiar", self)
        self.btn_clear.setFont(QFont("Arial", 12))
        self.btn_clear.setMinimumHeight(35)
        self.btn_clear.clicked.connect(self.text_edit.clear)
        self.content_layout.addWidget(self.btn_clear)
        
        # Barra de progreso con efecto de brillo
        self.progress_bar = GlowProgressBar(self)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(25)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setCurrentStyle(self.current_style)
        
        # Botón de abortar
        self.btn_abort = QPushButton("✖", self)
        self.btn_abort.setFixedSize(25, 25)
        self.btn_abort.setToolTip("Abortar ejecución")
        self.btn_abort.clicked.connect(self.abort_script)
        self.btn_abort.setVisible(False)
        
        progress_layout = QHBoxLayout()
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.btn_abort)
        self.content_layout.addLayout(progress_layout)
        
        self.content_widget.setLayout(self.content_layout)
        self.main_layout.addWidget(self.content_widget)
        self.setLayout(self.main_layout)
    
    def show_settings(self):
        """Muestra el diálogo de configuración"""
        settings_dialog = SettingsDialog(self)
        if settings_dialog.exec() == QDialog.DialogCode.Accepted:
            # Actualizar las iniciales si se guardaron cambios
            self.user_input_var = self.settings.value("user_initials", "")
    
    def run_script(self, script_name):
        """Ejecuta un script y monitorea su progreso"""
        self.reset_script_state()
        self.is_aborted = False
        
        if os.path.exists(self.progress_file):
            os.remove(self.progress_file)
        
        script_path = recurso_path(os.path.join("data", "python", script_name))
        if not os.path.exists(script_path):
            self.show_error(f"⚠️ No se encontró el script:\n{script_path}")
            return
        ruta_progress = recurso_path(os.path.join("data", "cache", "progress.txt"))
        python_exec = os.path.join(self.base_dir, "venv", "Scripts", "python.exe")
        
        try:
            # Obtener credenciales guardadas
            user_id = self.settings.value("user_id", "")
            user_password = self.settings.value("user_password", "")
            
            self.current_process = subprocess.Popen(
                [
                    python_exec,
                    script_path,
                    self.text_edit.toPlainText(),  # Texto principal
                    ruta_progress,                # Ruta del archivo de progreso
                    self.user_input_var,          # Iniciales del analista
                    user_id,                      # ID de usuario
                    user_password                 # Contraseña
                ],
                creationflags=subprocess.CREATE_NEW_CONSOLE
            )
            
            self.btn_abort.setVisible(True)
            self.progress_timer = QTimer(self)
            self.progress_timer.timeout.connect(self.update_real_progress)
            self.progress_timer.start(100)
            
        except Exception as e:
            self.show_error(f"Error al ejecutar: {e}")
            self.reset_script_state()
    
    def abort_script(self):
        """Cancela la ejecución del script actual"""
        if self.current_process:
            self.is_aborted = True
            self.current_process.terminate()
            self.progress_bar.setValue(0)
            self.show_error("Ejecución abortada por el usuario")
            self.reset_script_state()
    
    def reset_script_state(self):
        """Reinicia el estado de la ejecución"""
        if hasattr(self, 'progress_timer') and self.progress_timer.isActive():
            self.progress_timer.stop()
        self.current_process = None
        self.btn_abort.setVisible(False)
        if os.path.exists(self.progress_file):
            os.remove(self.progress_file)
    
    def update_real_progress(self):
        """Actualiza la barra de progreso"""
        if self.is_aborted:
            return
            
        if os.path.exists(self.progress_file):
            try:
                with open(self.progress_file, "r") as f:
                    progress = int(f.read().strip())
                    self.progress_bar.setValue(progress)
                    
                    if progress > 100:
                        self.reset_script_state()
                    elif progress==100:
                        self.is_aborted = True
                        self.current_process.terminate()
                        self.progress_bar.setValue(0)
                        self.show_success("Completado Exitosamente")
                        self.reset_script_state()

            except:
                pass
        elif self.current_process and self.current_process.poll() is not None:
            self.reset_script_state()
    
    def toggle_style(self):
        self.current_style = (self.current_style + 1) % 4  # Cicla entre 0-3
        self.settings.setValue("style", self.current_style)
        self.apply_style()
        
        # Actualizar icono del botón según el tema
        theme_icons = ["🌸", "🌙", "🌟", "🍷"]
        self.btn_toggle_style.setText(theme_icons[self.current_style])
        self.progress_bar.setCurrentStyle(self.current_style)
    
    def apply_style(self):
        palette = self.palette()
        
        # Configuraciones para cada tema
        themes = [
            {  # Sakura
                "window": "#FFF0F5",
                "stylesheet": """
                    QWidget {
                        background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #FFD6E0);
                    }
                    QPushButton {
                        background-color: #FFC0CB;
                        color: #5A3A3A;
                        border-radius: 10px;
                        padding: 8px;
                        font-size: 14px;
                        border: 1px solid #FF9EB5;
                    }
                    QPushButton:hover {
                        background-color: #FFB6C1;
                        border: 1px solid #FF6B8B;
                    }
                    QTextEdit {
                        background-color: #FFF9FB;
                        border-radius: 10px;
                        padding: 5px;
                        border: 1px solid #FFC0CB;
                        color: #5A3A3A;
                    }
                    QLabel {
                        background: transparent;
                        color: #5A3A3A;
                        border: none;
                        padding: 0px;
                    }
                """,
                "abort_style": """
                    QPushButton {
                        background-color: #FF6B8B;
                        color: white;
                        border-radius: 12px;
                        font-size: 14px;
                        border: 1px solid #FF3D6D;
                    }
                    QPushButton:hover {
                        background-color: #FF3D6D;
                    }
                """
            },
            {  # Dark
                "window": "#1A2A4A",
                "stylesheet": """
                    QWidget {
                        background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #1A2A4A);
                    }
                    QPushButton {
                        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2C3E66, stop:1 #1A2A4A);
                        color: #A8C4FF;
                        border-radius: 10px;
                        padding: 8px;
                        font-size: 14px;
                        border: 1px solid #4A90E2;
                    }
                    QPushButton:hover {
                        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #3A4E76, stop:1 #2A3A5A);
                        border: 1px solid #6AA8FF;
                    }
                    QTextEdit {
                        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2A3A5A, stop:1 #1C2C4C);
                        color: #E0E9FF;
                        border-radius: 10px;
                        padding: 5px;
                        border: 1px solid #4A90E2;
                    }
                    QLabel {
                        background: transparent;
                        color: #A8C4FF;
                        border: none;
                        padding: 0px;
                    }
                """,
                "abort_style": """
                    QPushButton {
                        background-color: #FF6B6B;
                        color: white;
                        border-radius: 12px;
                        font-size: 14px;
                        border: 1px solid #FF3D3D;
                    }
                    QPushButton:hover {
                        background-color: #FF3D3D;
                    }
                """
            },
            {  # Dorado
                "window": "#FFF4D5",
                "stylesheet": """
                    QWidget {
                        background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #FFE8B0);
                    }
                    QPushButton {
                        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #FFE8B0, stop:1 #FFD700);
                        color: #5A3A3A;
                        border-radius: 10px;
                        padding: 8px;
                        font-size: 14px;
                        border: 1px solid #D4AF37;
                    }
                    QPushButton:hover {
                        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #FFE8C0, stop:1 #FFDF00);
                        border: 1px solid #FFC800;
                    }
                    QTextEdit {
                        background-color: #FFF9E5;
                        border-radius: 10px;
                        padding: 5px;
                        border: 1px solid #D4AF37;
                        color: #5A3A3A;
                    }
                    QLabel {
                        background: transparent;
                        color: #5A3A3A;
                        border: none;
                        padding: 0px;
                    }
                """,
                "abort_style": """
                    QPushButton {
                        background-color: #D4AF37;
                        color: white;
                        border-radius: 12px;
                        font-size: 14px;
                        border: 1px solid #B8860B;
                    }
                    QPushButton:hover {
                        background-color: #B8860B;
                    }
                """
            },
            {  # Rojo Vino
                "window": "#2A0A0A",
                "stylesheet": """
                    QWidget {
                        background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #3A0A0A);
                    }
                    QPushButton {
                        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #5A1A1A, stop:1 #3A0A0A);
                        color: #E0C0C0;
                        border-radius: 10px;
                        padding: 8px;
                        font-size: 14px;
                        border: 1px solid #6D071A;
                    }
                    QPushButton:hover {
                        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #6A2A2A, stop:1 #4A1A1A);
                        border: 1px solid #8D071A;
                    }
                    QTextEdit {
                        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #4A1A1A, stop:1 #2A0A0A);
                        color: #E0C0C0;
                        border-radius: 10px;
                        padding: 5px;
                        border: 1px solid #6D071A;
                    }
                    QLabel {
                        background: transparent;
                        color: #E0C0C0;
                        border: none;
                        padding: 0px;
                    }
                """,
                "abort_style": """
                    QPushButton {
                        background-color: #8B0000;
                        color: white;
                        border-radius: 12px;
                        font-size: 14px;
                        border: 1px solid #6D071A;
                    }
                    QPushButton:hover {
                        background-color: #6D071A;
                    }
                """
            }
        ]
        
        theme = themes[self.current_style]
        
        # Aplicar tema
        palette.setColor(QPalette.ColorRole.Window, QColor(theme["window"]))
        self.setStyleSheet(theme["stylesheet"])
        self.btn_abort.setStyleSheet(theme["abort_style"])
        self.progress_bar.setCurrentStyle(self.current_style)
        self.setPalette(palette)
        self.title_bar.update()

    def show_error(self, message):
        QMessageBox.critical(self, "Error", message)
    
    def show_success(self, message):
        QMessageBox.information(self, "Completado", message)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ProgramLauncher()
    window.show()
    sys.exit(app.exec())