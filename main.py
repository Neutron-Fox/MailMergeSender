import sys
import os
import logging
if hasattr(sys, 'frozen'):
    log_dir = os.path.join(os.path.expanduser('~'), 'EmailSender_Logs')
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, 'main.log')
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(funcName)s: %(message)s',
        handlers=[
            logging.FileHandler(log_file, mode='w'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logging.info(f"Starting as frozen executable - Log file: {log_file}")
else:
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s: %(message)s')
    logging.info("Starting in development mode")
from PyQt5.QtWidgets import QApplication, QMessageBox
from PyQt5.QtCore import Qt, QTimer
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from mail_merge_sender import UniversalSender
    from loading_screen import LoadingScreen
    from theme import apply_theme
    logging.info("All modules imported successfully")
except ImportError as e:
    logging.error(f"CRITICAL: Failed to import modules: {e}")
    if hasattr(sys, 'frozen'):
        import traceback
        error_msg = f"Failed to load application modules:\n\n{str(e)}\n\n{traceback.format_exc()}"
        logging.error(error_msg)
        app = QApplication(sys.argv)
        QMessageBox.critical(None, "Import Error", error_msg)
        sys.exit(1)
    raise
class EmailSenderApp:
    def __init__(self):
        self.app = None
        self.main_window = None
        self.loading_screen = None
    def setup_application(self):
        """Setup Qt application with proper settings"""
        os.environ['QT_AUTO_SCREEN_SCALE_FACTOR'] = '1'
        if hasattr(Qt, 'AA_EnableHighDpiScaling'):
            QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
            QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
        self.app = QApplication(sys.argv)
        self.app.setApplicationName("Universal Email Sender")
        self.app.setApplicationVersion("2.0")
        apply_theme(self.app)
    def show_loading_screen(self):
        """Show loading screen during initialization"""
        self.loading_screen = LoadingScreen()
        self.loading_screen.show()
        self.loading_screen.update_progress(10, "Initializing application...")
        QApplication.processEvents()
    def initialize_application(self):
        """Initialize application components with loading progress"""
        try:
            self.loading_screen.update_progress(30, "Creating main window...")
            QApplication.processEvents()
            self.main_window = UniversalSender(loading_screen=self.loading_screen)
            self.loading_screen.update_progress(90, "Finalizing...")
            QApplication.processEvents()
            QTimer.singleShot(300, self.show_main_window)
        except Exception as e:
            print(f"Error creating main window: {e}")
            import traceback
            traceback.print_exc()
            if self.loading_screen:
                self.loading_screen.close()
            sys.exit(1)
    def show_main_window(self):
        """Show main window and close loading screen"""
        if self.loading_screen:
            self.loading_screen.close_loading()
        if self.main_window:
            self.main_window.show()
            logging.info("Main window displayed successfully")
    def run(self):
        """Run the application"""
        try:
            logging.info("Setting up application...")
            self.setup_application()
            logging.info("Showing loading screen...")
            self.show_loading_screen()
            logging.info("Initializing application components...")
            self.initialize_application()
            logging.info("Starting event loop...")
            return self.app.exec_()
        except Exception as e:
            error_msg = f"Critical error starting application: {e}"
            logging.error(error_msg)
            import traceback
            traceback.print_exc()
            logging.error(traceback.format_exc())
            try:
                QMessageBox.critical(None, "Application Error", 
                    f"{error_msg}\n\nSee log file for details:\n{os.path.join(os.path.expanduser('~'), 'EmailSender_Logs', 'main.log')}")
            except:
                pass
            return 1
def main():
    print("Starting Universal Email Sender...")
    try:
        app = EmailSenderApp()
        exit_code = app.run()
        print(f"Application exited with code: {exit_code}")
        return exit_code
    except KeyboardInterrupt:
        print("\\nApplication interrupted by user")
        return 1
    except Exception as e:
        print(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        return 1
if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)