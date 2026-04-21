"""
Data manager service for Universal Email Sender application.
Handles data state, processing, and placeholder operations.
"""
import re
from typing import Dict, Any, List, Set

from .constants import PLACEHOLDER_PATTERNS
from .exceptions import (
    PlaceholderError,
    InvalidPlaceholderError,
    MissingMappingError,
    InvalidDataError
)
from .validators import PlaceholderValidator, DataValidator
from .logger_setup import get_logger

logger = get_logger(__name__)


class PlaceholderManager:
    """Manages placeholder detection and extraction"""
    
    # Compiled regex patterns for efficiency
    _compiled_patterns = [re.compile(pattern, re.IGNORECASE) 
                         for pattern in PLACEHOLDER_PATTERNS]
    
    @staticmethod
    def extract_placeholders(text: str) -> List[str]:
        """
        Extract all placeholders from text.
        
        Args:
            text: Text containing placeholders
            
        Returns:
            Sorted list of unique placeholders in {PLACEHOLDER} format
        """
        try:
            if not text or not isinstance(text, str):
                logger.debug("No text to extract placeholders from")
                return []
            
            placeholders = set()
            
            for pattern in PlaceholderManager._compiled_patterns:
                matches = pattern.findall(text)
                for match in matches:
                    match = match.strip().upper()
                    if match and any(c.isalpha() for c in match):
                        standardized = '{' + match + '}'
                        placeholders.add(standardized)
            
            result = sorted(list(placeholders))
            logger.debug(f"Extracted {len(result)} placeholders")
            return result
        
        except Exception as e:
            logger.error(f"Error extracting placeholders: {e}")
            return []
    
    @staticmethod
    def suggest_mappings(placeholders: List[str], headers: List[str]) -> Dict[str, str]:
        """
        Suggest column mappings for placeholders using fuzzy matching.
        
        Args:
            placeholders: List of placeholders
            headers: List of available column headers
            
        Returns:
            Dictionary mapping placeholders to best-matching columns
        """
        try:
            if not placeholders or not headers:
                logger.debug("No placeholders or headers to map")
                return {}
            
            suggestions = {}
            
            for placeholder in placeholders:
                PlaceholderValidator.validate_placeholder(placeholder)
                placeholder_clean = placeholder.strip('{}').upper()
                
                best_match = None
                best_score = 0
                
                for header in headers:
                    header_upper = header.upper()
                    
                    # Exact match
                    if placeholder_clean == header_upper:
                        best_match = header
                        best_score = 100
                        break
                    
                    # Contains match
                    if placeholder_clean in header_upper or header_upper in placeholder_clean:
                        if 80 > best_score:
                            best_match = header
                            best_score = 80
                    
                    # Fuzzy match for common patterns
                    if PlaceholderManager._fuzzy_match(placeholder_clean, header_upper):
                        if 60 > best_score:
                            best_match = header
                            best_score = 60
                
                if best_match and best_score >= 60:
                    suggestions[placeholder] = best_match
                    logger.debug(f"Suggested mapping: {placeholder} → {best_match} (score: {best_score})")
            
            logger.info(f"Created {len(suggestions)}/{len(placeholders)} mappings")
            return suggestions
        
        except Exception as e:
            logger.error(f"Error suggesting mappings: {e}")
            return {}
    
    @staticmethod
    def _fuzzy_match(placeholder: str, header: str) -> bool:
        """
        Fuzzy match placeholder to header using common patterns.
        
        Args:
            placeholder: Placeholder name (uppercase)
            header: Header name (uppercase)
            
        Returns:
            True if fuzzy match found
        """
        # Common placeholder mappings
        common_maps = {
            'NAME': ['NAME', 'FULLNAME', 'FULL_NAME', 'SURNAME'],
            'EMAIL': ['EMAIL', 'MAIL', 'E-MAIL', 'E_MAIL'],
            'FIRST': ['FIRST', 'FIRSTNAME', 'FIRST_NAME'],
            'LAST': ['LAST', 'LASTNAME', 'LAST_NAME'],
            'PHONE': ['PHONE', 'TELEPHONE', 'TEL', 'MOBILE'],
            'ADDRESS': ['ADDRESS', 'ADDR', 'LOCATION'],
            'CITY': ['CITY', 'TOWN'],
            'COMPANY': ['COMPANY', 'BUSINESS', 'ORG'],
        }
        
        for key, values in common_maps.items():
            if key in placeholder:
                for value in values:
                    if value in header:
                        return True
        
        return False


class DataManager:
    """Manages application data state"""
    
    def __init__(self):
        """Initialize data manager"""
        self.imported_data = []
        self.processed_data = []
        self.filtered_data = []
        self.selected_rows: Set[int] = set()
        
        self.headers = []
        self.placeholders = []
        self.column_mapping = {}
        
        self.subject = ""
        self.template = ""
        
        self.attachments = []
        self.email_accounts = []
        self.selected_account = None
        
        self.formatting_rules = []
        self.bullet_config = {}
        
        logger.info("DataManager initialized")
    
    def set_imported_data(self, headers: List[str], data: List[List[Any]]) -> bool:
        """
        Set imported data.
        
        Args:
            headers: Column headers
            data: Row data
            
        Returns:
            True if set successfully
        """
        try:
            DataValidator.validate_headers(headers)
            DataValidator.validate_recipients_list(data)
            
            self.headers = headers
            self.imported_data = data
            self.filtered_data = data
            self.selected_rows.clear()
            self.processed_data.clear()
            
            logger.info(f"Imported data set: {len(headers)} columns, {len(data)} rows")
            return True
        
        except Exception as e:
            logger.error(f"Error setting imported data: {e}")
            raise InvalidDataError(str(e))
    
    def select_rows(self, row_indices: List[int]) -> bool:
        """
        Set selected rows for email sending.
        
        Args:
            row_indices: List of row indices (0-based)
            
        Returns:
            True if valid
        """
        try:
            if not row_indices:
                self.selected_rows.clear()
                return True
            
            # Validate indices
            for idx in row_indices:
                if not isinstance(idx, int) or idx < 0 or idx >= len(self.filtered_data):
                    raise ValueError(f"Invalid row index: {idx}")
            
            self.selected_rows = set(row_indices)
            logger.info(f"Selected {len(self.selected_rows)} rows for sending")
            return True
        
        except Exception as e:
            logger.error(f"Error selecting rows: {e}")
            raise InvalidDataError(str(e))
    
    def get_selected_recipients(self) -> List[Dict[str, Any]]:
        """
        Get selected recipient data with headers as keys.
        
        Returns:
            List of dictionaries with header keys
        """
        recipients = []
        for row_idx in sorted(self.selected_rows):
            if row_idx < len(self.filtered_data):
                row_data = self.filtered_data[row_idx]
                recipient = {self.headers[i]: row_data[i] 
                           for i in range(len(self.headers))}
                recipients.append(recipient)
        
        logger.debug(f"Returning {len(recipients)} recipients")
        return recipients
    
    def set_template(self, subject: str, template: str) -> bool:
        """
        Set email template.
        
        Args:
            subject: Email subject
            template: Email body template
            
        Returns:
            True if valid
        """
        try:
            DataValidator.validate_template(template)
            
            self.subject = subject.strip()
            self.template = template.strip()
            
            # Extract placeholders
            all_placeholders = set()
            all_placeholders.update(PlaceholderManager.extract_placeholders(subject))
            all_placeholders.update(PlaceholderManager.extract_placeholders(template))
            
            self.placeholders = sorted(list(all_placeholders))
            
            logger.info(f"Template set: {len(self.placeholders)} placeholders found")
            return True
        
        except Exception as e:
            logger.error(f"Error setting template: {e}")
            raise InvalidDataError(str(e))
    
    def set_column_mapping(self, mapping: Dict[str, str]) -> bool:
        """
        Set placeholder to column mapping.
        
        Args:
            mapping: Dictionary {placeholder: column}
            
        Returns:
            True if valid
        """
        try:
            DataValidator.validate_mapping(mapping, self.headers, self.placeholders)
            
            self.column_mapping = mapping.copy()
            logger.info(f"Column mapping set: {len(mapping)} mappings")
            return True
        
        except Exception as e:
            logger.error(f"Error setting mapping: {e}")
            raise InvalidDataError(str(e))
    
    def process_template_for_recipient(self, recipient_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Process template for specific recipient by replacing placeholders.
        
        Args:
            recipient_data: Recipient data dictionary
            
        Returns:
            Recipient dict with _processed_subject and _processed_template keys
        """
        try:
            processed = recipient_data.copy()
            
            # Replace placeholders in subject
            subject = self.subject
            for placeholder, column in self.column_mapping.items():
                if column in recipient_data:
                    value = str(recipient_data[column])
                    subject = subject.replace(placeholder, value)
            
            # Replace placeholders in template
            template = self.template
            for placeholder, column in self.column_mapping.items():
                if column in recipient_data:
                    value = str(recipient_data[column])
                    template = template.replace(placeholder, value)
            
            processed['_processed_subject'] = subject
            processed['_processed_template'] = template
            
            return processed
        
        except Exception as e:
            logger.error(f"Error processing template: {e}")
            raise InvalidDataError(str(e))
    
    def add_attachment(self, file_path: str) -> bool:
        """
        Add attachment file.
        
        Args:
            file_path: Path to attachment file
            
        Returns:
            True if added
        """
        if file_path not in self.attachments:
            self.attachments.append(file_path)
            logger.info(f"Attachment added: {file_path}")
        return True
    
    def remove_attachment(self, file_path: str) -> bool:
        """
        Remove attachment file.
        
        Args:
            file_path: Path to attachment
            
        Returns:
            True if removed
        """
        if file_path in self.attachments:
            self.attachments.remove(file_path)
            logger.info(f"Attachment removed: {file_path}")
        return True
    
    def add_formatting_rule(self, find_text: str, replace_text: str = "", 
                          special_value: str = None) -> bool:
        """
        Add find/replace formatting rule.
        
        Args:
            find_text: Text to find
            replace_text: Text to replace with
            special_value: Special value (\\n, space, tab)
            
        Returns:
            True if added
        """
        rule = {
            'find': find_text,
            'replace': replace_text or "",
            'special': special_value
        }
        self.formatting_rules.append(rule)
        logger.debug(f"Formatting rule added: find='{find_text}'")
        return True
    
    def apply_formatting(self, text: str) -> str:
        """
        Apply all formatting rules to text.
        
        Args:
            text: Text to format
            
        Returns:
            Formatted text
        """
        result = text
        
        for rule in self.formatting_rules:
            find_text = rule['find']
            replace_text = rule['special'] if rule['special'] else rule['replace']
            
            if find_text:
                result = result.replace(find_text, replace_text)
        
        return result
    
    def get_state_summary(self) -> Dict[str, Any]:
        """
        Get current application state summary.
        
        Returns:
            Dictionary with state information
        """
        return {
            'data_rows': len(self.imported_data),
            'selected_rows': len(self.selected_rows),
            'headers': len(self.headers),
            'placeholders': len(self.placeholders),
            'mappings': len(self.column_mapping),
            'attachments': len(self.attachments),
            'formatting_rules': len(self.formatting_rules),
            'template_length': len(self.template),
            'has_email_account': self.selected_account is not None
        }
