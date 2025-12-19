from PyQt5.QtWidgets import QWidget, QLabel, QProgressBar, QApplication
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont
from theme import var_theme
class LoadingScreen(QWidget):
    """Simple loading screen shown during application startup with pixmap-based layout"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        self.setWindowTitle("Loading")
        self.setAttribute(Qt.WA_TranslucentBackground, False)
        self.setup_ui()
        self.center_on_screen()
    def setup_ui(self):
        """Setup the loading screen UI with precise pixel positioning"""
        window_width = 500
        window_height = 220
        self.setFixedSize(window_width, window_height)
        self.title = QLabel("Universal Email Sender", self)
        self.title.setAlignment(Qt.AlignCenter)
        self.title.setFont(QFont('Segoe UI', 16, QFont.Bold))
        self.title.setStyleSheet(f"""
            QLabel {{
                color: {var_theme.colors['button_primary']};
                background-color: transparent;
                padding: 0px;
                margin: 0px;
            }}
        """)
        self.title.setGeometry(0, 25, window_width, 35)
        self.status_label = QLabel("Initializing application...", self)
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setFont(QFont('Segoe UI', 10))
        self.status_label.setStyleSheet(f"""
            QLabel {{
                color: {var_theme.colors['text_primary']};
                background-color: transparent;
                padding: 0px;
                margin: 0px;
            }}
        """)
        self.status_label.setGeometry(0, 75, window_width, 25)
        self.progress = QProgressBar(self)
        self.progress.setMinimum(0)
        self.progress.setMaximum(100)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        self.progress.setStyleSheet(f"""
            QProgressBar {{
                background-color: {var_theme.colors['secondary_bg']};
                border: 1px solid {var_theme.colors['border_primary']};
                border-radius: 4px;
                padding: 0px;
                margin: 0px;
            }}
            QProgressBar::chunk {{
                background-color: {var_theme.colors['button_primary']};
                border-radius: 3px;
            }}
        """)
        self.progress.setGeometry(50, 120, window_width - 100, 10)
        self.info = QLabel("Loading interface components...", self)
        self.info.setAlignment(Qt.AlignCenter)
        self.info.setFont(QFont('Segoe UI', 8))
        self.info.setStyleSheet(f"""
            QLabel {{
                color: {var_theme.colors['text_muted']};
                background-color: transparent;
                padding: 0px;
                margin: 0px;
            }}
        """)
        self.info.setGeometry(0, 150, window_width, 25)
        self.setStyleSheet(f"""
            QWidget {{
                background-color: {var_theme.colors['window_bg']};
            }}
        """)
        self.apply_dark_titlebar()
    def center_on_screen(self):
        """Center the loading screen on the screen"""
        from PyQt5.QtWidgets import QApplication
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)
    def apply_dark_titlebar(self):
        """Apply dark theme to Windows title bar using DWM API"""
        try:
            import ctypes
            from ctypes import wintypes
            hwnd = int(self.winId())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = ctypes.c_int(1)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(value),
                ctypes.sizeof(value)
            )
        except Exception as e:
            pass  
    def update_progress(self, value, status_text=""):
        """Update progress bar and status text"""
        self.progress.setValue(value)
        if status_text:
            self.status_label.setText(status_text)
        QApplication.processEvents()
    def close_loading(self):
        """Close the loading screen smoothly"""
        self.update_progress(100, "Complete!")
        QTimer.singleShot(200, self.close)
if __name__ == "__main__":
    from PyQt5.QtWidgets import QApplication
    import sys
    app = QApplication(sys.argv)
    loading = LoadingScreen()
    loading.show()
    for i in range(0, 101, 20):
        QTimer.singleShot(i * 50, lambda v=i: loading.update_progress(v, f"Loading... {v}%"))
    QTimer.singleShot(6000, app.quit)
    sys.exit(app.exec_())
