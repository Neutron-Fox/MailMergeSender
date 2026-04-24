"""
================================================================================
UNIVERSAL EMAIL SENDER - Application Entry Point
================================================================================

Purpose:
    This is the main entry point for the Universal Email Sender application.
    It initializes the Qt application framework, sets up logging, configures
    persistence, applies theming, and orchestrates the UI startup sequence.

Key Responsibilities:
    1. Initialize logging and data directories
    2. Setup Qt application with High-DPI scaling
    3. Apply dark theme to all UI elements
    4. Display loading screen during startup
    5. Create and show the main application window
    6. Handle errors and cleanup on exit

Workflow:
    main() → EmailSenderApp.run() → [Show loading screen]
         → [Initialize main window] → [Show window] → [Start event loop]

Architecture:
    - Startup with logging enabled for debugging
    - Multi-step initialization with progress feedback
    - Graceful error handling with user-friendly messages
    - Proper thread cleanup and session saving on exit
================================================================================
"""

import sys
import os
import logging

# Add project root to Python path for absolute imports (works in both dev and frozen exe)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Initialize logging to files and console before any other imports
from assets.logger_setup import LoggerSetup
from assets.constants import LOG_FILE
from assets.persistence import PersistenceManager

LoggerSetup.setup_logging()  # Set up rotating file logger in ~/EmailSender_Logs/
PersistenceManager.ensure_directories()  # Create ~/.MailMergeSender/ data directories

logger = logging.getLogger(__name__)

# Log startup mode (frozen executable vs development)
if hasattr(sys, 'frozen'):
    logger.info("Starting as frozen executable")
else:
    logger.info("Starting in development mode")

# Import PyQt5 framework and main application modules with error handling
from PyQt5.QtWidgets import QApplication, QMessageBox
from PyQt5.QtCore import Qt, QTimer

try:
    from source_code.mail_merge_sender import UniversalSender  # Main UI window
    from source_code.loading_screen import LoadingScreen  # Splash screen
    from source_code.theme import apply_theme  # Dark theme configuration
    logger.info("All modules imported successfully")
except ImportError as e:
    # Critical error: Cannot start without main modules
    logger.error(f"CRITICAL: Failed to import modules: {e}")
    if hasattr(sys, 'frozen'):
        import traceback
        error_msg = f"Failed to load application modules:\n\n{str(e)}\n\n{traceback.format_exc()}"
        logging.error(error_msg)
        app = QApplication(sys.argv)
        QMessageBox.critical(None, "Import Error", error_msg)
        sys.exit(1)
    raise


class EmailSenderApp:
    """
    Application orchestrator for Universal Email Sender.
    
    Responsibilities:
        - Setup Qt application framework
        - Configure High-DPI scaling and themes
        - Display loading screen during initialization
        - Create main window with progress feedback
        - Run Qt event loop
        - Handle startup errors gracefully
    """
    
    def __init__(self):
        self.app = None  # QApplication instance
        self.main_window = None  # UniversalSender main window
        self.loading_screen = None  # LoadingScreen splash
    
    def setup_application(self):
        """
        Setup Qt application framework with High-DPI support and theming.
        
        What this does:
            1. Enable automatic screen scale factor for High-DPI displays
            2. Create QApplication singleton
            3. Set application metadata (name, version)
            4. Apply dark theme stylesheet to entire application
        """
        # Enable High-DPI scaling for monitors with high pixel density
        os.environ['QT_AUTO_SCREEN_SCALE_FACTOR'] = '1'
        if hasattr(Qt, 'AA_EnableHighDpiScaling'):
            QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
            QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
        
        # Create Qt application (singleton - manages event loop and UI)
        self.app = QApplication(sys.argv)
        self.app.setApplicationName("Universal Email Sender")
        self.app.setApplicationVersion("2.0")
        
        # Apply dark theme stylesheet to all UI components
        apply_theme(self.app)
    
    def show_loading_screen(self):
        """
        Display loading screen with startup progress indicator.
        
        Shows splash screen with progress updates while application initializes.
        """
        self.loading_screen = LoadingScreen()
        self.loading_screen.show()
        self.loading_screen.update_progress(10, "Initializing application...")
        QApplication.processEvents()  # Let Qt render the loading screen
    
    def initialize_application(self):
        """
        Create main application window with progress tracking.
        
        What this does:
            1. Update loading screen (30%)
            2. Create main window (UniversalSender)
            3. Pre-load all UI tabs
            4. Schedule main window display after delay
        
        Note: Uses QTimer for delayed display to allow tab loading to complete
        """
        try:
            self.loading_screen.update_progress(30, "Creating main window...")
            QApplication.processEvents()
            
            # Create main window - this pre-loads all 5 tabs
            self.main_window = UniversalSender(loading_screen=self.loading_screen)
            
            self.loading_screen.update_progress(90, "Finalizing...")
            QApplication.processEvents()
            
            # Schedule main window display after 300ms to ensure UI is ready
            QTimer.singleShot(300, self.show_main_window)
        except Exception as e:
            print(f"Error creating main window: {e}")
            import traceback
            traceback.print_exc()
            if self.loading_screen:
                self.loading_screen.close()
            sys.exit(1)
    
    def show_main_window(self):
        """
        Display main window and close loading screen.
        """
        if self.loading_screen:
            self.loading_screen.close_loading()  # Fade out splash screen
        if self.main_window:
            self.main_window.show()  # Show main application window
            logging.info("Main window displayed successfully")
    
    def run(self):
        """
        Run the application event loop.
        
        Returns:
            int: Exit code (0 for success, 1 for error)
        
        Orchestrates:
            1. Qt application setup
            2. Loading screen display
            3. Application initialization
            4. Qt event loop execution
        """
        try:
            logging.info("Setting up application...")
            self.setup_application()
            
            logging.info("Showing loading screen...")
            self.show_loading_screen()
            
            logging.info("Initializing application components...")
            self.initialize_application()
            
            logging.info("Starting event loop...")
            # Start Qt event loop - blocks until application closes
            return self.app.exec_()
        except Exception as e:
            error_msg = f"Critical error starting application: {e}"
            logging.error(error_msg)
            import traceback
            traceback.print_exc()
            logging.error(traceback.format_exc())
            try:
                QMessageBox.critical(None, "Application Error", 
                    f"{error_msg}\n\nSee log file for details:\n{LOG_FILE}")
            except:
                pass
            return 1


def main():
    """
    Main entry point for application execution.
    
    What this does:
        1. Create application instance
        2. Run application event loop
        3. Cleanup threads on exit
        4. Save session state
        5. Handle keyboard interrupt and unexpected errors
    
    Returns:
        int: Exit code for sys.exit()
    """
    print("Starting Universal Email Sender...")
    try:
        from source_code.threading_manager import ThreadingManager
        
        # Create and run application
        app = EmailSenderApp()
        exit_code = app.run()
        print(f"Application exited with code: {exit_code}")
        
        # Cleanup: Stop all background threads
        ThreadingManager.cleanup()
        
        return exit_code
    
    except KeyboardInterrupt:
        # User pressed Ctrl+C
        print("\nApplication interrupted by user")
        logger.info("Application interrupted by user")
        from source_code.threading_manager import ThreadingManager
        ThreadingManager.cleanup()
        return 1
    
    except Exception as e:
        # Unexpected error during startup
        print(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        logger.error(f"Unexpected error: {e}")
        logger.error(traceback.format_exc())
        from source_code.threading_manager import ThreadingManager
        ThreadingManager.cleanup()
        return 1


if __name__ == "__main__":
    # Execute main entry point and exit with returned code
    exit_code = main()
    sys.exit(exit_code)