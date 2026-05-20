"""
Threading manager for Universal Email Sender application.
Handles long-running operations on separate threads without freezing UI.
"""
from PyQt5.QtCore import QThread, pyqtSignal
from concurrent.futures import ThreadPoolExecutor
from typing import Callable, Any

from assets.logger_setup import get_logger

logger = get_logger(__name__)


class WorkerThread(QThread):
    """QThread worker for running long-running operations"""
    
    # Signals
    started = pyqtSignal()
    progress = pyqtSignal(str, int, int)  # message, current, total
    finished = pyqtSignal(bool, str, Any)  # success, message, result
    error = pyqtSignal(str)  # error message
    
    def __init__(self, operation: Callable, *args, **kwargs):
        """Initialize worker thread"""
        super().__init__()
        self.operation = operation
        self.args = args
        self.kwargs = kwargs
        self._is_running = True
        
        logger.debug(f"WorkerThread created for operation: {operation.__name__}")
    
    def run(self):
        """Execute operation in thread"""
        try:
            self.started.emit()
            logger.info(f"Starting operation: {self.operation.__name__}")
            
            result = self.operation(*self.args, **self.kwargs)
            
            if isinstance(result, tuple) and len(result) == 2:
                success, message = result
                self.finished.emit(success, message, None)
            elif isinstance(result, dict) and 'success' in result:
                success = result.get('success', False)
                message = result.get('message', '')
                data = {k: v for k, v in result.items() 
                       if k not in ['success', 'message']}
                self.finished.emit(success, message, data if data else None)
            else:
                self.finished.emit(True, "Operation completed", result)
            
            logger.info(f"[OK] Operation completed: {self.operation.__name__}")
        
        except Exception as e:
            error_msg = f"Error in {self.operation.__name__}: {str(e)}"
            logger.error(error_msg)
            self.error.emit(error_msg)
            self.finished.emit(False, error_msg, None)
    
    def stop(self):
        """Stop thread execution"""
        self._is_running = False
        self.quit()
        self.wait()
        logger.debug("WorkerThread stopped")


class ThreadingManager:
    """Manages background thread operations"""
    
    _executor = ThreadPoolExecutor(max_workers=3)
    _active_threads = []
    
    @staticmethod
    def run_async(operation: Callable, *args, **kwargs) -> WorkerThread:
        """Run operation asynchronously in QThread"""
        try:
            thread = WorkerThread(operation, *args, **kwargs)
            ThreadingManager._active_threads.append(thread)
            
            # Auto cleanup finished threads
            def cleanup():
                if thread in ThreadingManager._active_threads:
                    ThreadingManager._active_threads.remove(thread)
            
            thread.finished.connect(cleanup)
            thread.error.connect(cleanup)
            
            thread.start()
            logger.debug(f"Started async operation: {operation.__name__}")
            
            return thread
        
        except Exception as e:
            logger.error(f"Error starting async operation: {e}")
            raise
    
    @staticmethod
    def run_with_progress(operation: Callable, progress_callback: Callable = None, 
                         *args, **kwargs) -> WorkerThread:
        """Run operation with progress callback"""
        thread = ThreadingManager.run_async(operation, *args, **kwargs)
        
        if progress_callback:
            thread.progress.connect(progress_callback)
        
        return thread
    
    @staticmethod
    def cleanup():
        """Cleanup all active threads"""
        for thread in ThreadingManager._active_threads[:]:
            try:
                thread.stop()
            except Exception as e:
                logger.warning(f"Error stopping thread: {e}")
        
        ThreadingManager._active_threads.clear()
        logger.info("All threads cleaned up")
    
    @staticmethod
    def get_active_thread_count() -> int:
        """Get number of active threads"""
        return len([t for t in ThreadingManager._active_threads if t.isRunning()])
