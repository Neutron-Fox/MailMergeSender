"""
Input validation module for Universal Email Sender application.
Provides validation for emails, files, paths, and other user inputs.
"""
import os
import re
from pathlib import Path

from .constants import (
    EMAIL_VALIDATION_PATTERN,
    MAX_ATTACHMENT_SIZE,
    SUPPORTED_FILE_TYPES,
    TEXT_FILE_DELIMITERS
)
from .exceptions import (
    InvalidEmailError,
    InvalidPathError,
    FileSizeError,
    FileNotFoundError,
    UnsupportedFileTypeError,
    InvalidDataError
)


class EmailValidator:
    """Validates email addresses"""
    
    @staticmethod
    def validate(email: str) -> bool:
        """
        Validate email address using RFC 5322 pattern.
        
        Args:
            email: Email address to validate
            
        Returns:
            True if valid, raises InvalidEmailError if invalid
            
        Raises:
            InvalidEmailError: If email format is invalid
        """
        if not email or not isinstance(email, str):
            raise InvalidEmailError(email)
        
        email = email.strip().lower()
        
        # Basic checks
        if len(email) > 254:
            raise InvalidEmailError(email)
        
        if email.count('@') != 1:
            raise InvalidEmailError(email)
        
        # RFC 5322 pattern
        if not re.match(EMAIL_VALIDATION_PATTERN, email):
            raise InvalidEmailError(email)
        
        local_part, domain = email.rsplit('@', 1)
        
        if len(local_part) > 64:
            raise InvalidEmailError(email)
        
        if not domain or '.' not in domain:
            raise InvalidEmailError(email)
        
        return True
    
    @staticmethod
    def validate_batch(emails: list) -> tuple[list, list]:
        """
        Validate multiple emails and return valid/invalid lists.
        
        Args:
            emails: List of email addresses
            
        Returns:
            Tuple of (valid_emails, invalid_emails)
        """
        valid = []
        invalid = []
        
        for email in emails:
            try:
                EmailValidator.validate(email)
                valid.append(email)
            except InvalidEmailError:
                invalid.append(email)
        
        return valid, invalid


class FileValidator:
    """Validates file operations and properties"""
    
    @staticmethod
    def validate_exists(file_path: str) -> bool:
        """
        Validate that file exists.
        
        Args:
            file_path: Path to file
            
        Returns:
            True if file exists
            
        Raises:
            FileNotFoundError: If file doesn't exist
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(file_path)
        return True
    
    @staticmethod
    def validate_is_file(file_path: str) -> bool:
        """
        Validate that path is a file (not directory).
        
        Args:
            file_path: Path to check
            
        Returns:
            True if path is a file
            
        Raises:
            InvalidPathError: If path is not a file
        """
        FileValidator.validate_exists(file_path)
        
        if not os.path.isfile(file_path):
            raise InvalidPathError(file_path, "Path is not a file")
        return True
    
    @staticmethod
    def validate_readable(file_path: str) -> bool:
        """
        Validate that file is readable.
        
        Args:
            file_path: Path to file
            
        Returns:
            True if file is readable
            
        Raises:
            InvalidPathError: If file is not readable
        """
        FileValidator.validate_is_file(file_path)
        
        if not os.access(file_path, os.R_OK):
            raise InvalidPathError(file_path, "File is not readable")
        return True
    
    @staticmethod
    def validate_size(file_path: str, max_size: int = None) -> int:
        """
        Validate file size.
        
        Args:
            file_path: Path to file
            max_size: Maximum size in bytes (None = no limit)
            
        Returns:
            File size in bytes
            
        Raises:
            FileSizeError: If file exceeds max_size
        """
        FileValidator.validate_readable(file_path)
        
        size = os.path.getsize(file_path)
        
        if max_size and size > max_size:
            size_mb = size / (1024 * 1024)
            max_mb = max_size / (1024 * 1024)
            raise FileSizeError(file_path, size_mb, max_mb)
        
        return size
    
    @staticmethod
    def validate_file_type(file_path: str) -> str:
        """
        Validate file type is supported.
        
        Args:
            file_path: Path to file
            
        Returns:
            File extension (lowercase)
            
        Raises:
            UnsupportedFileTypeError: If file type not supported
        """
        _, ext = os.path.splitext(file_path.lower())
        
        supported_ext = []
        for patterns in SUPPORTED_FILE_TYPES.values():
            for pattern in patterns:
                supported_ext.append(pattern.replace('*', ''))
        
        if ext not in supported_ext:
            raise UnsupportedFileTypeError(ext)
        
        return ext
    
    @staticmethod
    def validate_attachment(file_path: str) -> bool:
        """
        Validate attachment file for email sending.
        
        Args:
            file_path: Path to attachment
            
        Returns:
            True if valid
            
        Raises:
            FileNotFoundError: If file doesn't exist
            FileSizeError: If file exceeds max size
        """
        FileValidator.validate_readable(file_path)
        FileValidator.validate_size(file_path, MAX_ATTACHMENT_SIZE)
        return True


class PathValidator:
    """Validates file paths for security"""
    
    @staticmethod
    def validate_safe_path(file_path: str, base_dir: str = None) -> bool:
        """
        Validate path doesn't contain traversal attempts.
        
        Args:
            file_path: Path to validate
            base_dir: Base directory (if None, uses current directory)
            
        Returns:
            True if path is safe
            
        Raises:
            InvalidPathError: If path contains traversal
        """
        if not file_path:
            raise InvalidPathError(file_path, "Empty path")
        
        if base_dir is None:
            base_dir = os.getcwd()
        
        # Resolve path to absolute
        try:
            abs_path = os.path.abspath(file_path)
            abs_base = os.path.abspath(base_dir)
        except Exception as e:
            raise InvalidPathError(file_path, str(e))
        
        # Check if path is within base_dir
        if not abs_path.startswith(abs_base):
            raise InvalidPathError(file_path, "Path traversal detected")
        
        return True
    
    @staticmethod
    def validate_writable_path(dir_path: str) -> bool:
        """
        Validate directory is writable.
        
        Args:
            dir_path: Directory path
            
        Returns:
            True if directory is writable
            
        Raises:
            InvalidPathError: If not writable
        """
        if not os.path.exists(dir_path):
            try:
                os.makedirs(dir_path)
            except Exception as e:
                raise InvalidPathError(dir_path, f"Cannot create directory: {e}")
        
        if not os.access(dir_path, os.W_OK):
            raise InvalidPathError(dir_path, "Directory is not writable")
        
        return True


class DataValidator:
    """Validates data structure and content"""
    
    @staticmethod
    def validate_recipient_data(recipient: dict, required_fields: list = None) -> bool:
        """
        Validate recipient data structure.
        
        Args:
            recipient: Recipient data dictionary
            required_fields: List of required fields
            
        Returns:
            True if valid
            
        Raises:
            InvalidDataError: If data is invalid
        """
        if not isinstance(recipient, dict):
            raise InvalidDataError("Recipient must be dictionary")
        
        if required_fields:
            missing = [f for f in required_fields if f not in recipient]
            if missing:
                raise InvalidDataError(f"Missing required fields: {missing}")
        
        return True
    
    @staticmethod
    def validate_template(template: str) -> bool:
        """
        Validate email template.
        
        Args:
            template: Email template text
            
        Returns:
            True if valid
            
        Raises:
            InvalidDataError: If template invalid
        """
        if not template or not isinstance(template, str):
            raise InvalidDataError("Template must be non-empty string")
        
        if len(template.strip()) == 0:
            raise InvalidDataError("Template is empty")
        
        if len(template) > 1000000:  # 1 MB limit
            raise InvalidDataError("Template too large (max 1 MB)")
        
        return True
    
    @staticmethod
    def validate_headers(headers: list) -> bool:
        """
        Validate data headers.
        
        Args:
            headers: List of column headers
            
        Returns:
            True if valid
            
        Raises:
            InvalidDataError: If headers invalid
        """
        if not headers or not isinstance(headers, list):
            raise InvalidDataError("Headers must be non-empty list")
        
        if len(headers) == 0:
            raise InvalidDataError("No columns found in data")
        
        # Check for duplicates
        if len(headers) != len(set(h.upper() for h in headers)):
            raise InvalidDataError("Duplicate column names found")
        
        return True
    
    @staticmethod
    def validate_recipients_list(recipients: list) -> bool:
        """
        Validate recipients list.
        
        Args:
            recipients: List of recipient dictionaries
            
        Returns:
            True if valid
            
        Raises:
            InvalidDataError: If list invalid
        """
        if not recipients or not isinstance(recipients, list):
            raise InvalidDataError("Recipients must be non-empty list")
        
        if len(recipients) == 0:
            raise InvalidDataError("No recipients to send to")
        
        return True


class PlaceholderValidator:
    """Validates placeholder format and mapping"""
    
    VALID_FORMATS = ['{', '<', '[[', '<<', '{{', '[']
    
    @staticmethod
    def validate_placeholder(placeholder: str) -> bool:
        """
        Validate placeholder format.
        
        Args:
            placeholder: Placeholder string
            
        Returns:
            True if valid
            
        Raises:
            InvalidDataError: If format invalid
        """
        if not placeholder or not isinstance(placeholder, str):
            raise InvalidDataError("Placeholder must be non-empty string")
        
        # Check if starts with valid format
        valid = False
        for fmt in PlaceholderValidator.VALID_FORMATS:
            if placeholder.startswith(fmt):
                valid = True
                break
        
        if not valid:
            raise InvalidDataError(f"Invalid placeholder format: {placeholder}")
        
        return True
    
    @staticmethod
    def validate_mapping(mapping: dict, headers: list, placeholders: list) -> bool:
        """
        Validate placeholder to column mapping.
        
        Args:
            mapping: Mapping dictionary {placeholder: column}
            headers: Available column headers
            placeholders: Detected placeholders
            
        Returns:
            True if valid
            
        Raises:
            InvalidDataError: If mapping invalid
        """
        for placeholder, column in mapping.items():
            if column not in headers:
                raise InvalidDataError(f"Column '{column}' not found in data")
        
        return True
