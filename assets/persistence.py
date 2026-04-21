"""
Data persistence module for Universal Email Sender application.
Handles saving and loading sessions, templates, and settings.
"""
import json
import os
import pickle
from pathlib import Path
from datetime import datetime

from .constants import (
    DATA_DIR,
    SESSION_FILE,
    TEMPLATES_DIR,
    IMPORTS_DIR
)
from .exceptions import (
    SessionSaveError,
    SessionLoadError,
    TemplatePersistenceError,
    PersistenceError
)
from .logger_setup import get_logger

logger = get_logger(__name__)


class PersistenceManager:
    """Manages data persistence for sessions, templates, and imports"""
    
    @staticmethod
    def ensure_directories():
        """Ensure all data directories exist"""
        try:
            for directory in [DATA_DIR, TEMPLATES_DIR, IMPORTS_DIR]:
                os.makedirs(directory, exist_ok=True)
                logger.debug(f"Ensured directory exists: {directory}")
        except Exception as e:
            logger.error(f"Failed to create data directories: {e}")
            raise PersistenceError(f"Cannot create data directories: {e}")
    
    # ========================================================================
    # SESSION PERSISTENCE
    # ========================================================================
    
    @staticmethod
    def save_session(session_data: dict) -> bool:
        """
        Save session data to JSON file.
        
        Args:
            session_data: Dictionary containing session state
            
        Returns:
            True if saved successfully
            
        Raises:
            SessionSaveError: If save fails
        """
        try:
            PersistenceManager.ensure_directories()
            
            # Add metadata
            session_data['_saved_at'] = datetime.now().isoformat()
            
            # Convert non-serializable objects
            cleaned_data = PersistenceManager._clean_for_json(session_data)
            
            with open(SESSION_FILE, 'w', encoding='utf-8') as f:
                json.dump(cleaned_data, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Session saved to: {SESSION_FILE}")
            return True
        
        except Exception as e:
            logger.error(f"Failed to save session: {e}")
            raise SessionSaveError(str(e))
    
    @staticmethod
    def load_session() -> dict:
        """
        Load session data from JSON file.
        
        Returns:
            Dictionary containing session state
            
        Raises:
            SessionLoadError: If load fails
        """
        try:
            if not os.path.exists(SESSION_FILE):
                logger.warning(f"Session file not found: {SESSION_FILE}")
                return {}
            
            with open(SESSION_FILE, 'r', encoding='utf-8') as f:
                session_data = json.load(f)
            
            logger.info(f"Session loaded from: {SESSION_FILE}")
            logger.info(f"Session saved at: {session_data.get('_saved_at', 'unknown')}")
            
            return session_data
        
        except json.JSONDecodeError as e:
            logger.error(f"Session file corrupted: {e}")
            raise SessionLoadError(f"Session file is corrupted: {e}")
        except Exception as e:
            logger.error(f"Failed to load session: {e}")
            raise SessionLoadError(str(e))
    
    @staticmethod
    def clear_session() -> bool:
        """
        Clear saved session data.
        
        Returns:
            True if cleared
        """
        try:
            if os.path.exists(SESSION_FILE):
                os.remove(SESSION_FILE)
                logger.info("Session cleared")
            return True
        except Exception as e:
            logger.error(f"Failed to clear session: {e}")
            return False
    
    # ========================================================================
    # TEMPLATE PERSISTENCE
    # ========================================================================
    
    @staticmethod
    def save_template(name: str, template_text: str, subject: str = "") -> bool:
        """
        Save email template to file.
        
        Args:
            name: Template name
            template_text: Template content
            subject: Email subject
            
        Returns:
            True if saved
            
        Raises:
            TemplatePersistenceError: If save fails
        """
        try:
            PersistenceManager.ensure_directories()
            
            template_path = os.path.join(TEMPLATES_DIR, f"{name}.json")
            
            template_data = {
                'name': name,
                'subject': subject,
                'template': template_text,
                'created_at': datetime.now().isoformat()
            }
            
            with open(template_path, 'w', encoding='utf-8') as f:
                json.dump(template_data, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Template saved: {name}")
            return True
        
        except Exception as e:
            logger.error(f"Failed to save template {name}: {e}")
            raise TemplatePersistenceError(f"Failed to save template: {e}")
    
    @staticmethod
    def load_template(name: str) -> dict:
        """
        Load email template from file.
        
        Args:
            name: Template name
            
        Returns:
            Dictionary with 'subject' and 'template' keys
            
        Raises:
            TemplatePersistenceError: If load fails
        """
        try:
            template_path = os.path.join(TEMPLATES_DIR, f"{name}.json")
            
            if not os.path.exists(template_path):
                raise TemplatePersistenceError(f"Template not found: {name}")
            
            with open(template_path, 'r', encoding='utf-8') as f:
                template_data = json.load(f)
            
            logger.info(f"Template loaded: {name}")
            return template_data
        
        except json.JSONDecodeError as e:
            logger.error(f"Template file corrupted: {name}")
            raise TemplatePersistenceError(f"Template file corrupted: {e}")
        except Exception as e:
            logger.error(f"Failed to load template {name}: {e}")
            raise TemplatePersistenceError(str(e))
    
    @staticmethod
    def list_templates() -> list:
        """
        List all saved templates.
        
        Returns:
            List of template names
        """
        try:
            PersistenceManager.ensure_directories()
            
            templates = []
            for file in os.listdir(TEMPLATES_DIR):
                if file.endswith('.json'):
                    templates.append(file.replace('.json', ''))
            
            logger.debug(f"Found {len(templates)} templates")
            return sorted(templates)
        
        except Exception as e:
            logger.error(f"Failed to list templates: {e}")
            return []
    
    @staticmethod
    def delete_template(name: str) -> bool:
        """
        Delete template file.
        
        Args:
            name: Template name
            
        Returns:
            True if deleted
        """
        try:
            template_path = os.path.join(TEMPLATES_DIR, f"{name}.json")
            
            if os.path.exists(template_path):
                os.remove(template_path)
                logger.info(f"Template deleted: {name}")
                return True
            
            return False
        
        except Exception as e:
            logger.error(f"Failed to delete template {name}: {e}")
            return False
    
    # ========================================================================
    # IMPORT HISTORY PERSISTENCE
    # ========================================================================
    
    @staticmethod
    def save_import(file_path: str, headers: list, row_count: int) -> bool:
        """
        Save import history.
        
        Args:
            file_path: Source file path
            headers: Column headers
            row_count: Number of rows imported
            
        Returns:
            True if saved
        """
        try:
            PersistenceManager.ensure_directories()
            
            import_record = {
                'file_path': file_path,
                'file_name': os.path.basename(file_path),
                'headers': headers,
                'row_count': row_count,
                'imported_at': datetime.now().isoformat()
            }
            
            # Save as pickle for later use
            import_name = Path(file_path).stem
            import_path = os.path.join(IMPORTS_DIR, f"{import_name}.pkl")
            
            with open(import_path, 'wb') as f:
                pickle.dump(import_record, f)
            
            logger.info(f"Import history saved: {import_name}")
            return True
        
        except Exception as e:
            logger.error(f"Failed to save import history: {e}")
            return False
    
    @staticmethod
    def list_import_history() -> list:
        """
        List import history.
        
        Returns:
            List of import records
        """
        try:
            PersistenceManager.ensure_directories()
            
            history = []
            for file in os.listdir(IMPORTS_DIR):
                if file.endswith('.pkl'):
                    import_path = os.path.join(IMPORTS_DIR, file)
                    try:
                        with open(import_path, 'rb') as f:
                            record = pickle.load(f)
                            history.append(record)
                    except Exception as e:
                        logger.warning(f"Could not load import record {file}: {e}")
            
            return sorted(history, key=lambda x: x['imported_at'], reverse=True)
        
        except Exception as e:
            logger.error(f"Failed to list import history: {e}")
            return []
    
    # ========================================================================
    # UTILITY METHODS
    # ========================================================================
    
    @staticmethod
    def _clean_for_json(obj):
        """
        Clean object for JSON serialization.
        
        Args:
            obj: Object to clean
            
        Returns:
            JSON-serializable object
        """
        if isinstance(obj, dict):
            return {k: PersistenceManager._clean_for_json(v) 
                    for k, v in obj.items()}
        elif isinstance(obj, (list, tuple)):
            return [PersistenceManager._clean_for_json(i) for i in obj]
        elif isinstance(obj, set):
            return list(obj)
        elif isinstance(obj, (str, int, float, bool, type(None))):
            return obj
        else:
            # For other types, convert to string
            return str(obj)
    
    @staticmethod
    def get_storage_info() -> dict:
        """
        Get storage usage information.
        
        Returns:
            Dictionary with storage stats
        """
        try:
            def get_dir_size(path):
                total = 0
                for entry in os.scandir(path):
                    if entry.is_file():
                        total += entry.stat().st_size
                    elif entry.is_dir():
                        total += get_dir_size(entry.path)
                return total
            
            PersistenceManager.ensure_directories()
            
            return {
                'session_size': os.path.getsize(SESSION_FILE) if os.path.exists(SESSION_FILE) else 0,
                'templates_size': get_dir_size(TEMPLATES_DIR) if os.path.exists(TEMPLATES_DIR) else 0,
                'imports_size': get_dir_size(IMPORTS_DIR) if os.path.exists(IMPORTS_DIR) else 0,
                'total_size': sum([
                    os.path.getsize(SESSION_FILE) if os.path.exists(SESSION_FILE) else 0,
                    get_dir_size(TEMPLATES_DIR) if os.path.exists(TEMPLATES_DIR) else 0,
                    get_dir_size(IMPORTS_DIR) if os.path.exists(IMPORTS_DIR) else 0,
                ])
            }
        
        except Exception as e:
            logger.error(f"Failed to get storage info: {e}")
            return {}
