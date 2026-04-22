"""
Email service for Universal Email Sender application.
Handles Microsoft Outlook integration with COM automation and proper resource management.
"""
import os
import sys
import subprocess
import time
from typing import Dict, Any, List, Optional
from .constants import (
    OUTLOOK_PROCESS_NAME,
    OUTLOOK_PROCESS_TIMEOUT,
    OUTLOOK_INSTALL_PATHS,
    OUTLOOK_STARTUP_WAIT_FROZEN,
    OUTLOOK_STARTUP_WAIT_DEV,
    OUTLOOK_STARTUP_MIN_WAIT,
    EMAIL_SEND_DELAY
)
from .exceptions import (
    OutlookError,
    OutlookConnectionError,
    OutlookAccountError,
    NoOutlookAccountsError,
    EmailSendError,
    InvalidRecipientError,
    MissingDependencyError
)
from .validators import EmailValidator
from .logger_setup import get_logger

logger = get_logger(__name__)


class EmailService:
    """Handles email sending via Microsoft Outlook"""
    
    _outlook_instance = None
    
    @staticmethod
    def check_outlook_running() -> bool:
        """
        Check if Outlook process is running.
        
        Returns:
            True if Outlook is running
        """
        try:
            result = subprocess.run(
                ['tasklist', '/FI', f'IMAGENAME eq {OUTLOOK_PROCESS_NAME}'],
                capture_output=True,
                text=True,
                timeout=OUTLOOK_PROCESS_TIMEOUT
            )
            
            is_running = OUTLOOK_PROCESS_NAME in result.stdout
            status = "✓" if is_running else "✗"
            logger.info(f"{status} Outlook process running: {is_running}")
            return is_running
        
        except Exception as e:
            logger.warning(f"Could not check Outlook status: {e}")
            return False
    
    @staticmethod
    def start_outlook() -> bool:
        """
        Start Microsoft Outlook in background (minimized).
        
        Returns:
            True if Outlook started successfully
            
        Raises:
            OutlookError: If Outlook cannot be started
        """
        try:
            if EmailService.check_outlook_running():
                logger.info("Outlook already running")
                return True
            
            logger.info("Starting Microsoft Outlook in background...")
            
            # Try configured paths first
            outlook_path = None
            for path in OUTLOOK_INSTALL_PATHS:
                if os.path.exists(path):
                    outlook_path = path
                    logger.info(f"Found Outlook at: {path}")
                    break
            
            if not outlook_path:
                logger.debug("Trying to start Outlook via system PATH...")
                outlook_path = "outlook.exe"
            
            # Start minimized
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 6  # Minimized
            
            subprocess.Popen([outlook_path], startupinfo=startupinfo)
            logger.info("Outlook startup command sent")
            
            # Wait for Outlook to initialize
            wait_time = OUTLOOK_STARTUP_WAIT_FROZEN if hasattr(sys, 'frozen') else OUTLOOK_STARTUP_WAIT_DEV
            logger.info(f"Waiting {wait_time}s for Outlook to initialize...")
            time.sleep(wait_time)
            
            # Verify it started
            if EmailService.check_outlook_running():
                logger.info("✓ Outlook started successfully")
                return True
            else:
                logger.warning("Outlook process not detected after startup")
                time.sleep(OUTLOOK_STARTUP_MIN_WAIT)
                return EmailService.check_outlook_running()
        
        except Exception as e:
            logger.error(f"Error starting Outlook: {e}")
            raise OutlookError(f"Failed to start Outlook: {e}")
    
    @staticmethod
    def get_email_accounts() -> List[Dict[str, Any]]:
        """
        Get list of email accounts configured in Outlook.
        
        Returns:
            List of account dictionaries with 'email' and 'account_object' keys
            
        Raises:
            OutlookError: If account retrieval fails
            MissingDependencyError: If pywin32 not installed
        """
        accounts = []
        
        try:
            logger.info("Loading Outlook email accounts...")
            
            # Check for pywin32
            try:
                import win32com.client
            except ImportError as e:
                logger.error("pywin32 not installed")
                raise MissingDependencyError("pywin32")
            
            # Get or create Outlook instance
            if EmailService._outlook_instance is not None:
                try:
                    # Test if cached instance is still valid
                    _ = EmailService._outlook_instance.Version
                    logger.info("✓ Reusing cached Outlook instance")
                    outlook = EmailService._outlook_instance
                except Exception:
                    logger.warning("Cached instance invalid, creating new connection")
                    EmailService._outlook_instance = None
            
            if EmailService._outlook_instance is None:
                try:
                    logger.info("Creating new Outlook connection...")
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    version = outlook.Version
                    logger.info(f"✓ Connected to Outlook (version {version})")
                    EmailService._outlook_instance = outlook
                except Exception as e:
                    logger.error(f"Failed to connect to Outlook: {e}")
                    raise OutlookConnectionError(str(e))
            
            # Get MAPI namespace
            try:
                namespace = outlook.GetNamespace("MAPI")
                logger.info("Connected to MAPI namespace")
            except Exception as e:
                logger.error(f"Failed to get MAPI namespace: {e}")
                raise OutlookConnectionError(f"MAPI error: {e}")
            
            # Get accounts
            try:
                outlook_accounts = outlook.Session.Accounts
                account_count = outlook_accounts.Count
                logger.info(f"Found {account_count} Outlook account(s)")
                
                if account_count == 0:
                    raise NoOutlookAccountsError()
            
            except NoOutlookAccountsError:
                raise
            except Exception as e:
                logger.error(f"Failed to access accounts: {e}")
                raise OutlookAccountError(str(e))
            
            # Extract account details
            for i in range(1, account_count + 1):
                try:
                    account = outlook_accounts.Item(i)
                    account_name = account.DisplayName
                    
                    try:
                        email_address = account.SmtpAddress
                        
                        if email_address and '@' in email_address:
                            # Validate email
                            try:
                                EmailValidator.validate(email_address)
                                accounts.append({
                                    'email': email_address,
                                    'account_object': account,
                                    'display_name': account_name
                                })
                                logger.info(f"✓ Account {i}: {email_address} ({account_name})")
                            except Exception as e:
                                logger.warning(f"✗ Account {i}: Invalid email - {e}")
                        else:
                            logger.warning(f"✗ Account {i}: No valid SMTP address")
                    
                    except Exception as e:
                        logger.warning(f"✗ Account {i} ({account_name}): {e}")
                
                except Exception as e:
                    logger.warning(f"Error processing account {i}: {e}")
            
            if not accounts:
                raise NoOutlookAccountsError()
            
            logger.info(f"✓ Successfully loaded {len(accounts)} account(s)")
            return accounts
        
        except (NoOutlookAccountsError, MissingDependencyError, OutlookConnectionError, OutlookAccountError):
            raise
        except Exception as e:
            logger.error(f"Critical error loading accounts: {e}")
            raise OutlookAccountError(str(e))
    
    @staticmethod
    def send_emails(recipients: List[Dict[str, Any]], 
                   subject: str,
                   template: str,
                   account: Dict[str, Any],
                   attachments: Optional[List[str]] = None,
                   callback=None) -> Dict[str, Any]:
        """
        Send emails to recipients using Outlook account.
        
        Args:
            recipients: List of recipient dictionaries with email and other fields
            subject: Email subject line
            template: Email body template
            account: Account dictionary from get_email_accounts()
            attachments: Optional list of file paths to attach
            callback: Optional callback(current, total) for progress tracking
            
        Returns:
            Dictionary with 'success', 'message', 'sent', 'failed', 'failed_details'
            
        Raises:
            EmailSendError: If sending fails
        """
        try:
            logger.info(f"Starting email send: {len(recipients)} recipients")
            
            import win32com.client
            
            sender_email = account.get('email')
            if not sender_email:
                raise InvalidRecipientError("No account email")
            
            # Get Outlook instance
            try:
                outlook = EmailService._outlook_instance
                if outlook is None:
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    EmailService._outlook_instance = outlook
                    logger.info("Created new Outlook instance for sending")
            except Exception as e:
                logger.error(f"Failed to get Outlook instance: {e}")
                raise EmailSendError(f"Outlook connection failed: {e}")
            
            # Find account object
            account_object = account.get('account_object')
            if not account_object:
                raise EmailSendError("Account object missing")
            
            sent_count = 0
            failed_count = 0
            failed_details = []
            
            for i, recipient_data in enumerate(recipients, 1):
                try:
                    if callback:
                        callback(i, len(recipients))
                    
                    # Extract email
                    recipient_email = None
                    for field in ['EMAIL', 'Email', 'email', 'E-mail', 'Mail', 'MAIL']:
                        if field in recipient_data:
                            candidate = str(recipient_data[field]).strip()
                            if candidate and '@' in candidate:
                                recipient_email = candidate
                                break
                    
                    if not recipient_email:
                        raise InvalidRecipientError(str(recipient_data))
                    
                    # Validate recipient
                    EmailValidator.validate(recipient_email)
                    
                    # Create mail item
                    mail_item = outlook.CreateItem(0)
                    mail_item.SendUsingAccount = account_object
                    mail_item.To = recipient_email
                    
                    # Set subject (use processed if available)
                    if '_processed_subject' in recipient_data:
                        mail_item.Subject = recipient_data['_processed_subject']
                    else:
                        mail_item.Subject = subject
                    
                    # Set body (use processed if available)
                    if '_processed_template' in recipient_data:
                        body_text = recipient_data['_processed_template']
                    else:
                        body_text = template
                    
                    mail_item.HTMLBody = body_text.replace('\n', '<br>')
                    
                    # Add attachments
                    if attachments:
                        for att_path in attachments:
                            if os.path.exists(att_path):
                                try:
                                    mail_item.Attachments.Add(att_path)
                                except Exception as att_e:
                                    logger.warning(f"Could not add attachment {att_path}: {att_e}")
                    
                    # Send
                    mail_item.SendUsingAccount = account_object
                    mail_item.Send()
                    
                    logger.debug(f"✓ Email {i}: Sent to {recipient_email}")
                    sent_count += 1
                    
                    time.sleep(EMAIL_SEND_DELAY)
                
                except InvalidRecipientError as e:
                    failed_count += 1
                    error_msg = f"Recipient {i}: Invalid email - {str(e)}"
                    failed_details.append(error_msg)
                    logger.warning(error_msg)
                
                except Exception as e:
                    failed_count += 1
                    error_msg = f"Recipient {i}: {str(e)}"
                    failed_details.append(error_msg)
                    logger.warning(error_msg)
            
            success = failed_count == 0
            message = f"Sent {sent_count} emails"
            if failed_count > 0:
                message += f", {failed_count} failed"
            
            logger.info(f"✓ Email send complete: {message}")
            
            return {
                'success': success,
                'message': message,
                'sent': sent_count,
                'failed': failed_count,
                'failed_details': failed_details
            }
        
        except (InvalidRecipientError, EmailSendError):
            raise
        except Exception as e:
            logger.error(f"Critical error sending emails: {e}")
            return {
                'success': False,
                'message': f'Critical error: {str(e)}',
                'sent': 0,
                'failed': len(recipients),
                'failed_details': [str(e)]
            }
