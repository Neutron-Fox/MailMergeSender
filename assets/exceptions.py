"""
Custom exception hierarchy for Universal Email Sender application.
Provides specific error handling for different failure scenarios.
"""


class MailMergeSenderException(Exception):
    """Base exception for all Mail Merge Sender errors"""
    pass


# ============================================================================
# FILE IMPORT EXCEPTIONS
# ============================================================================

class FileImportError(MailMergeSenderException):
    """Base exception for file import operations"""
    pass


class FileNotFoundError(FileImportError):
    """Raised when imported file does not exist"""
    def __init__(self, file_path):
        self.file_path = file_path
        super().__init__(f"File not found: {file_path}")


class UnsupportedFileTypeError(FileImportError):
    """Raised when file type is not supported"""
    def __init__(self, file_type):
        self.file_type = file_type
        super().__init__(f"Unsupported file type: {file_type}")


class FileReadError(FileImportError):
    """Raised when file cannot be read"""
    def __init__(self, file_path, reason):
        self.file_path = file_path
        self.reason = reason
        super().__init__(f"Error reading file {file_path}: {reason}")


class InvalidFileFormatError(FileImportError):
    """Raised when file format is corrupted or invalid"""
    def __init__(self, file_path, reason):
        self.file_path = file_path
        self.reason = reason
        super().__init__(f"Invalid file format {file_path}: {reason}")


class NoDataFoundError(FileImportError):
    """Raised when file is empty or contains no data"""
    def __init__(self, file_path):
        self.file_path = file_path
        super().__init__(f"No data found in file: {file_path}")


# ============================================================================
# VALIDATION EXCEPTIONS
# ============================================================================

class ValidationError(MailMergeSenderException):
    """Base exception for validation errors"""
    pass


class InvalidEmailError(ValidationError):
    """Raised when email address is invalid"""
    def __init__(self, email):
        self.email = email
        super().__init__(f"Invalid email address: {email}")


class InvalidPathError(ValidationError):
    """Raised when file path is invalid"""
    def __init__(self, path, reason="Path traversal detected"):
        self.path = path
        super().__init__(f"Invalid path: {path} - {reason}")


class FileSizeError(ValidationError):
    """Raised when file exceeds size limit"""
    def __init__(self, file_path, size_mb, max_size_mb):
        self.file_path = file_path
        self.size_mb = size_mb
        self.max_size_mb = max_size_mb
        super().__init__(
            f"File too large: {size_mb:.1f} MB (max: {max_size_mb} MB)"
        )


class InvalidDataError(ValidationError):
    """Raised when data fails validation"""
    def __init__(self, reason):
        self.reason = reason
        super().__init__(f"Invalid data: {reason}")


# ============================================================================
# EMAIL & OUTLOOK EXCEPTIONS
# ============================================================================

class EmailError(MailMergeSenderException):
    """Base exception for email operations"""
    pass


class OutlookError(EmailError):
    """Base exception for Outlook-related errors"""
    pass


class OutlookNotRunningError(OutlookError):
    """Raised when Outlook is not running"""
    def __init__(self):
        super().__init__(
            "Outlook is not running. Please start Outlook and try again."
        )


class OutlookConnectionError(OutlookError):
    """Raised when cannot connect to Outlook"""
    def __init__(self, reason):
        self.reason = reason
        super().__init__(f"Cannot connect to Outlook: {reason}")


class OutlookAccountError(OutlookError):
    """Raised when Outlook account operations fail"""
    def __init__(self, reason):
        self.reason = reason
        super().__init__(f"Outlook account error: {reason}")


class NoOutlookAccountsError(OutlookError):
    """Raised when no Outlook accounts are configured"""
    def __init__(self):
        super().__init__(
            "No email accounts found in Outlook. "
            "Please configure at least one email account in Outlook."
        )


class EmailSendError(EmailError):
    """Raised when email sending fails"""
    def __init__(self, reason, recipient=None):
        self.reason = reason
        self.recipient = recipient
        if recipient:
            super().__init__(f"Failed to send email to {recipient}: {reason}")
        else:
            super().__init__(f"Email send error: {reason}")


class InvalidRecipientError(EmailError):
    """Raised when recipient email is invalid"""
    def __init__(self, recipient):
        self.recipient = recipient
        super().__init__(f"Invalid recipient email: {recipient}")


# ============================================================================
# PLACEHOLDER & TEMPLATE EXCEPTIONS
# ============================================================================

class TemplateError(MailMergeSenderException):
    """Base exception for template operations"""
    pass


class PlaceholderError(TemplateError):
    """Base exception for placeholder operations"""
    pass


class InvalidPlaceholderError(PlaceholderError):
    """Raised when placeholder format is invalid"""
    def __init__(self, placeholder):
        self.placeholder = placeholder
        super().__init__(f"Invalid placeholder format: {placeholder}")


class MissingMappingError(PlaceholderError):
    """Raised when placeholder mapping is missing"""
    def __init__(self, placeholder, available_columns):
        self.placeholder = placeholder
        self.available_columns = available_columns
        super().__init__(
            f"No mapping found for placeholder {placeholder}. "
            f"Available columns: {', '.join(available_columns)}"
        )


class InvalidTemplateError(TemplateError):
    """Raised when template is invalid"""
    def __init__(self, reason):
        self.reason = reason
        super().__init__(f"Invalid template: {reason}")


# ============================================================================
# DATA & PERSISTENCE EXCEPTIONS
# ============================================================================

class DataError(MailMergeSenderException):
    """Base exception for data operations"""
    pass


class PersistenceError(DataError):
    """Base exception for data persistence operations"""
    pass


class SessionSaveError(PersistenceError):
    """Raised when session cannot be saved"""
    def __init__(self, reason):
        self.reason = reason
        super().__init__(f"Failed to save session: {reason}")


class SessionLoadError(PersistenceError):
    """Raised when session cannot be loaded"""
    def __init__(self, reason):
        self.reason = reason
        super().__init__(f"Failed to load session: {reason}")


class TemplatePersistenceError(PersistenceError):
    """Raised when template cannot be saved/loaded"""
    def __init__(self, reason):
        self.reason = reason
        super().__init__(f"Template persistence error: {reason}")


# ============================================================================
# THREADING & ASYNC EXCEPTIONS
# ============================================================================

class ThreadingError(MailMergeSenderException):
    """Base exception for threading operations"""
    pass


class OperationTimeoutError(ThreadingError):
    """Raised when operation exceeds timeout"""
    def __init__(self, operation, timeout_seconds):
        self.operation = operation
        self.timeout_seconds = timeout_seconds
        super().__init__(
            f"Operation '{operation}' exceeded timeout ({timeout_seconds}s)"
        )


class OperationCancelledError(ThreadingError):
    """Raised when operation is cancelled by user"""
    def __init__(self, operation):
        self.operation = operation
        super().__init__(f"Operation cancelled: {operation}")


# ============================================================================
# CONFIGURATION EXCEPTIONS
# ============================================================================

class ConfigurationError(MailMergeSenderException):
    """Base exception for configuration errors"""
    pass


class MissingDependencyError(ConfigurationError):
    """Raised when required dependency is missing"""
    def __init__(self, package_name):
        self.package_name = package_name
        super().__init__(
            f"Missing required package: {package_name}. "
            f"Install with: pip install {package_name}"
        )
