"""
Email service for Universal Email Sender application.
Handles Microsoft Outlook integration with COM automation and proper resource management.
"""
import os
import sys
import subprocess
import time
import html as html_lib
import re
from urllib.parse import unquote
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
    def _prepare_html_body(body_text: str) -> str:
        """Preserve plain-text spacing in mixed text+HTML bodies while safeguarding signature HTML.
        
        This converts newlines to <br> for plain text parts but ignores text inside HTML.
        It respects the <!--SIG_START--> marker to pass signature blocks 100% untouched.
        """
        if body_text is None:
            return ""

        body_str = str(body_text)
        if not body_str:
            return ""

        # If it already seems to have complete HTML structure and no signature markers, pass through
        if "<html" in body_str.lower() and "<body" in body_str.lower() and "<!--SIG_START-->" not in body_str:
            return body_str

        # Split by the signature markers added in mail_merge_sender.py
        sig_parts = re.split(r'(<!--SIG_START-->.*?<!--SIG_END-->)', body_str, flags=re.DOTALL)
        
        final_html = []
        
        for sig_part in sig_parts:
            if sig_part.startswith('<!--SIG_START-->') and sig_part.endswith('<!--SIG_END-->'):
                # This is a signature block. Pass it EXACTLY as is, stripping only the markers.
                clean_sig = sig_part[16:-14]
                final_html.append(clean_sig)
            else:
                if not sig_part:
                    continue
                
                # This is non-signature text (user template, subject placeholders, etc).
                # We apply the standard whitespace-to-HTML conversion, preserving any other HTML tags.
                tag_parts = re.split(r'(<[^>]+>)', sig_part)
                for part in tag_parts:
                    if not part:
                        continue
                    if part.startswith('<') and part.endswith('>'):
                        final_html.append(part)
                    else:
                        # Convert plain text spacing into explicit HTML spacing
                        converted = (
                            part
                            .replace('&', '&amp;')
                            .replace('<', '&lt;')
                            .replace('>', '&gt;')
                            .replace('\t', '&nbsp;&nbsp;&nbsp;&nbsp;')
                            .replace('\r\n', '\n')
                            .replace('\r', '\n')
                            .replace('\n', '<br>')
                        )
                        final_html.append(converted)

        return ''.join(final_html)

    @staticmethod
    def _prepare_signature_inline_images(raw_html: str, signature_dir: str, signature_name: str) -> Dict[str, Any]:
        """Convert local signature image references to CID references and collect inline image metadata."""
        if not raw_html:
            return {'content': '', 'inline_images': []}

        inline_images = []
        path_to_cid = {}

        def replace_image_source(match):
            prefix = match.group(1)
            source = match.group(2).strip()
            suffix = match.group(3)

            if source.lower().startswith(("http://", "https://", "cid:", "data:")):
                return match.group(0)

            normalized_source = unquote(html_lib.unescape(source))
            normalized_source = normalized_source.replace("/", os.sep).replace("\\", os.sep)
            candidate_path = os.path.normpath(os.path.join(signature_dir, normalized_source))

            if not os.path.exists(candidate_path):
                candidate_path = os.path.normpath(os.path.join(signature_dir, os.path.basename(normalized_source)))

            if not os.path.exists(candidate_path):
                logger.warning(f"Signature image not found: {source} (base: {signature_dir})")
                return match.group(0)

            normalized_path_key = os.path.normcase(candidate_path)
            if normalized_path_key not in path_to_cid:
                safe_name = re.sub(r'[^A-Za-z0-9]+', '-', signature_name or 'signature').strip('-').lower() or 'signature'
                cid = f"{safe_name}-{len(path_to_cid) + 1}@inline"
                path_to_cid[normalized_path_key] = cid
                inline_images.append({'path': candidate_path, 'cid': cid})

            cid_value = path_to_cid[normalized_path_key]
            return f'{prefix}cid:{cid_value}{suffix}'

        updated_html = re.sub(
            r'(<img[^>]+src=["\'])([^"\']+)(["\'][^>]*>)',
            replace_image_source,
            raw_html,
            flags=re.IGNORECASE,
        )
        return {'content': updated_html, 'inline_images': inline_images}

    @staticmethod
    def _extract_signature_body_html(raw_html: str, signature_dir: str = "", signature_name: str = "") -> Dict[str, Any]:
        """Extract the HTML body and styles from an Outlook signature file."""
        if not raw_html:
            return {'content': '', 'inline_images': []}

        # Extract styles from head to preserve formatting
        styles = ""
        style_matches = re.finditer(r"<style[^>]*>.*?</style>", raw_html, re.IGNORECASE | re.DOTALL)
        for match in style_matches:
            styles += match.group(0) + "\n"

        # Try to extract just the body to avoid duplicate html/head/body tags if we are injecting into another HTML doc,
        # but keep as much formatting as possible.
        body_match = re.search(r"<body[^>]*>(.*?)</body>", raw_html, re.IGNORECASE | re.DOTALL)
        if body_match:
            raw_html = styles + body_match.group(1)
        else:
            raw_html = styles + raw_html

        inline_images = []
        if signature_dir:
            prepared = EmailService._prepare_signature_inline_images(raw_html, signature_dir, signature_name)
            raw_html = prepared.get('content', raw_html)
            inline_images = prepared.get('inline_images', [])

        return {
            'content': raw_html.strip(),
            'inline_images': inline_images,
        }

    @staticmethod
    def _load_signature_file(file_path: str) -> Dict[str, Any]:
        """Load and normalize a signature file into an HTML fragment or plain text."""
        content = ""
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    content = f.read().strip()
                break
            except Exception:
                content = ""

        if not content:
            with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read().strip()

        signature_name = os.path.splitext(os.path.basename(file_path))[0]
        if file_path.lower().endswith(('.htm', '.html')):
            return EmailService._extract_signature_body_html(content, os.path.dirname(file_path), signature_name)

        return {
            'content': content,
            'inline_images': [],
        }

    @staticmethod
    def _resolve_send_account(outlook: Any, sender_email: str, preferred_account: Any = None) -> Any:
        """Resolve the exact Outlook account object for the selected sender email.

        This avoids stale COM account objects and ensures SendUsingAccount points to
        the account currently selected in the UI dropdown.
        """
        normalized_email = (sender_email or "").strip().lower()
        if not normalized_email:
            raise EmailSendError("Selected sender email is empty")

        # First try preferred account object from UI list if it still belongs to this session.
        if preferred_account is not None:
            try:
                preferred_smtp = str(preferred_account.SmtpAddress).strip().lower()
                if preferred_smtp == normalized_email:
                    return preferred_account
            except Exception:
                logger.warning("Preferred account object is stale; resolving by SMTP address")

        # Resolve account object by SMTP from current Outlook session.
        try:
            session_accounts = outlook.Session.Accounts
            for i in range(1, session_accounts.Count + 1):
                account_obj = session_accounts.Item(i)
                smtp = str(account_obj.SmtpAddress).strip().lower()
                if smtp == normalized_email:
                    return account_obj
        except Exception as e:
            raise EmailSendError(f"Could not resolve selected account '{sender_email}': {e}")

        raise EmailSendError(f"Selected account not found in Outlook session: {sender_email}")

    @staticmethod
    def _create_mail_item_for_account(outlook: Any, account_object: Any) -> Any:
        """Create a mail item scoped to the selected account store when possible.

        Outlook can fallback to the default account when creating a generic item.
        Creating the draft inside the selected account store improves sender fidelity.
        """
        # olFolderDrafts = 16
        try:
            drafts_folder = account_object.DeliveryStore.GetDefaultFolder(16)
            if drafts_folder is not None:
                mail_item = drafts_folder.Items.Add("IPM.Note")
                logger.info("Created mail item in selected account Drafts store")
                return mail_item
        except Exception as e:
            logger.warning(f"Could not create account-scoped draft item: {e}")

        logger.warning("Falling back to generic CreateItem(0)")
        return outlook.CreateItem(0)

    @staticmethod
    def resolve_sender_email(account: Dict[str, Any]) -> str:
        """Resolve and validate the actual sender account address for the provided account selection."""
        try:
            import win32com.client

            sender_email = (account or {}).get('email', '')
            if not sender_email:
                raise EmailSendError("No selected sender email")

            outlook = EmailService._outlook_instance
            if outlook is None:
                outlook = win32com.client.Dispatch("Outlook.Application")
                EmailService._outlook_instance = outlook

            preferred_account_object = account.get('account_object')
            account_object = EmailService._resolve_send_account(outlook, sender_email, preferred_account_object)
            resolved_email = str(account_object.SmtpAddress).strip()
            if not resolved_email:
                raise EmailSendError("Resolved sender account has no SMTP address")
            return resolved_email
        except Exception as e:
            if isinstance(e, EmailSendError):
                raise
            raise EmailSendError(f"Could not resolve sender account: {e}")

    @staticmethod
    def _apply_sender_to_mail_item(mail_item: Any, account_object: Any, sender_email: str) -> None:
        """Apply sender identity using both supported Outlook mechanisms.

        - SendUsingAccount: preferred for normal account sending.
        - SentOnBehalfOfName: explicit From identity when supported/required.
        """
        mail_item.SendUsingAccount = account_object

        # Some Outlook setups require explicit From identity.
        # If this fails, keep SendUsingAccount as the primary mechanism.
        try:
            mail_item.SentOnBehalfOfName = sender_email
            logger.info(f"Set SentOnBehalfOfName to: {sender_email}")
        except Exception as e:
            logger.warning(f"Could not set SentOnBehalfOfName: {e}")

        # Best-effort verification of account binding.
        try:
            bound_account = mail_item.SendUsingAccount
            bound_smtp = str(bound_account.SmtpAddress).strip().lower() if bound_account else ""
            if bound_smtp and bound_smtp != sender_email.strip().lower():
                raise EmailSendError(
                    f"Sender binding mismatch. Selected={sender_email}, Bound={bound_smtp}. "
                    "Outlook may be forcing a different account."
                )
        except EmailSendError:
            raise
        except Exception as e:
            logger.warning(f"Could not verify bound sender account: {e}")
    
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
            status = "[OK]" if is_running else "[NO]"
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
                logger.info("[OK] Outlook started successfully")
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
                    logger.info("[OK] Reusing cached Outlook instance")
                    outlook = EmailService._outlook_instance
                except Exception:
                    logger.warning("Cached instance invalid, creating new connection")
                    EmailService._outlook_instance = None
            
            if EmailService._outlook_instance is None:
                try:
                    logger.info("Creating new Outlook connection...")
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    version = outlook.Version
                    logger.info(f"[OK] Connected to Outlook (version {version})")
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
                                logger.info(f"[OK] Account {i}: {email_address} ({account_name})")
                            except Exception as e:
                                logger.warning(f"[ERROR] Account {i}: Invalid email - {e}")
                        else:
                            logger.warning(f"[ERROR] Account {i}: No valid SMTP address")
                    
                    except Exception as e:
                        logger.warning(f"[ERROR] Account {i} ({account_name}): {e}")
                
                except Exception as e:
                    logger.warning(f"Error processing account {i}: {e}")
            
            if not accounts:
                raise NoOutlookAccountsError()
            
            logger.info(f"[OK] Successfully loaded {len(accounts)} account(s)")
            return accounts
        
        except (NoOutlookAccountsError, MissingDependencyError, OutlookConnectionError, OutlookAccountError):
            raise
        except Exception as e:
            logger.error(f"Critical error loading accounts: {e}")
            raise OutlookAccountError(str(e))

    @staticmethod
    def get_outlook_signatures() -> List[Dict[str, str]]:
        """
        Get list of email signatures available in Outlook.
        
        Signatures are stored in user's signature folder (Windows AppData).
        Supports .txt and .htm/.html signature files.
        
        Returns:
            List of signature dictionaries with 'name' and 'content' keys
            
        Raises:
            OutlookError: If signature retrieval fails
        """
        signatures = []
        
        try:
            logger.info("Loading Outlook signatures...")
            
            # Signature files are stored in AppData\Roaming\Microsoft\Signatures
            signatures_folder = os.path.expanduser(
                r"~\AppData\Roaming\Microsoft\Signatures"
            )
            
            if not os.path.exists(signatures_folder):
                logger.warning(f"Signatures folder not found: {signatures_folder}")
                raise OutlookError(f"Signatures folder not found at {signatures_folder}")
            
            logger.info(f"Scanning signatures folder: {signatures_folder}")
            
            # Scan for signature files (.txt, .htm, .html)
            # Outlook stores the same signature in multiple companion files; keep only one entry per name.
            best_signatures = {}
            priority_map = {
                '.txt': 1,
                '.htm': 2,
                '.html': 2,
            }
            for filename in os.listdir(signatures_folder):
                file_path = os.path.join(signatures_folder, filename)
                
                # Skip directories and hidden files
                if os.path.isdir(file_path) or filename.startswith('.'):
                    continue
                
                # Support .txt, .htm, .html files
                if filename.lower().endswith(('.txt', '.htm', '.html')):
                    try:
                        # Extract signature name (filename without extension)
                        sig_name = os.path.splitext(filename)[0]
                        extension = os.path.splitext(filename)[1].lower()
                        priority = priority_map.get(extension, 0)
                        
                        # Read signature content
                        signature_payload = EmailService._load_signature_file(file_path)
                        content = signature_payload.get('content', '')
                        inline_images = signature_payload.get('inline_images', [])
                        
                        if content:
                            current_entry = best_signatures.get(sig_name)
                            if current_entry is None or priority >= current_entry['priority']:
                                best_signatures[sig_name] = {
                                    'name': sig_name,
                                    'content': content,
                                    'inline_images': inline_images,
                                    'filename': filename,
                                    'priority': priority,
                                }
                                logger.info(f"[OK] Signature candidate: {sig_name} ({filename}, {len(content)} chars)")
                        else:
                            logger.debug(f"Skipped empty signature: {sig_name}")
                    
                    except Exception as e:
                        logger.warning(f"Error reading signature '{filename}': {e}")
                        continue
            
            signatures = [
                {
                    'name': entry['name'],
                    'content': entry['content'],
                    'inline_images': entry.get('inline_images', []),
                    'filename': entry['filename']
                }
                for entry in sorted(best_signatures.values(), key=lambda item: item['name'].lower())
            ]

            if not signatures:
                logger.warning("No signatures found in Outlook signatures folder")
                raise OutlookError("No signatures found")
            
            logger.info(f"[OK] Successfully loaded {len(signatures)} signature(s)")
            return signatures
        
        except OutlookError:
            raise
        except Exception as e:
            logger.error(f"Critical error loading signatures: {e}")
            raise OutlookError(f"Failed to load signatures: {e}")
    
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
            
            # Resolve a fresh account object from current Outlook session using the selected email.
            preferred_account_object = account.get('account_object')
            account_object = EmailService._resolve_send_account(outlook, sender_email, preferred_account_object)
            
            logger.info(f"Using account: {sender_email}")
            logger.info(f"Account object type: {type(account_object)}")
            
            sent_count = 0
            failed_count = 0
            failed_details = []
            
            for i, recipient_data in enumerate(recipients, 1):
                try:
                    if callback:
                        callback(i, len(recipients))
                    
                    # Extract recipient emails. Prefer explicit fields prepared by UI logic.
                    raw_candidates = []
                    if '_recipient_emails' in recipient_data:
                        raw_value = recipient_data.get('_recipient_emails', [])
                        if isinstance(raw_value, (list, tuple, set)):
                            raw_candidates.extend(str(v).strip() for v in raw_value if str(v).strip())
                        else:
                            raw_text = str(raw_value).strip()
                            if raw_text:
                                raw_candidates.append(raw_text)
                    elif '_recipient_email' in recipient_data:
                        raw_text = str(recipient_data.get('_recipient_email', '')).strip()
                        if raw_text:
                            raw_candidates.append(raw_text)
                    else:
                        for field in ['EMAIL', 'Email', 'email', 'E-mail', 'Mail', 'MAIL']:
                            if field in recipient_data:
                                candidate = str(recipient_data[field]).strip()
                                if candidate and '@' in candidate:
                                    raw_candidates.append(candidate)
                                    break

                    # Split any combined values and validate each email.
                    recipient_emails = []
                    for candidate in raw_candidates:
                        parts = [candidate]
                        for separator in [';', ',', '\n', '\r']:
                            new_parts = []
                            for part in parts:
                                new_parts.extend(part.split(separator))
                            parts = new_parts
                        for part in parts:
                            email = part.strip()
                            if email and '@' in email and email not in recipient_emails:
                                recipient_emails.append(email)

                    if not recipient_emails:
                        raise InvalidRecipientError(str(recipient_data))

                    # Validate all recipients
                    for recipient_email in recipient_emails:
                        EmailValidator.validate(recipient_email)
                    
                    # Create mail item in selected account context when possible.
                    mail_item = EmailService._create_mail_item_for_account(outlook, account_object)
                    EmailService._apply_sender_to_mail_item(mail_item, account_object, sender_email)
                    logger.info(f"Mail item created with SendUsingAccount set to: {sender_email}")
                    mail_item.To = '; '.join(recipient_emails)
                    logger.info(f"Recipients: {mail_item.To}")
                    
                    # Set subject (use processed if available)
                    if '_processed_subject' in recipient_data:
                        mail_item.Subject = recipient_data['_processed_subject']
                        logger.debug(f"Using processed subject: {mail_item.Subject[:50]}...")
                    else:
                        mail_item.Subject = subject
                        logger.debug(f"Using default subject: {subject[:50]}...")
                    
                    # Set body (use processed if available)
                    if '_processed_template' in recipient_data:
                        body_text = recipient_data['_processed_template']
                        logger.debug(f"Using processed template ({len(body_text)} chars)")
                    else:
                        body_text = template
                        logger.debug(f"Using default template ({len(body_text)} chars)")

                    # olFormatHTML = 2
                    mail_item.BodyFormat = 2
                    mail_item.HTMLBody = EmailService._prepare_html_body(body_text)

                    # Add inline signature images (cid attachments)
                    inline_images = recipient_data.get('_inline_images', [])
                    for image_data in inline_images:
                        image_path = str(image_data.get('path', '')).strip()
                        image_cid = str(image_data.get('cid', '')).strip()
                        if not image_path or not image_cid:
                            continue
                        if not os.path.exists(image_path):
                            logger.warning(f"Inline signature image path not found: {image_path}")
                            continue
                        try:
                            attachment = mail_item.Attachments.Add(image_path)
                            prop = attachment.PropertyAccessor
                            # PR_ATTACH_CONTENT_ID (PT_STRING8/PT_UNICODE)
                            prop.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", image_cid)
                            # PR_ATTACHMENT_HIDDEN
                            prop.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", True)
                        except Exception as image_err:
                            logger.warning(f"Could not add inline signature image '{image_path}': {image_err}")
                    
                    # Add attachments
                    if attachments:
                        for att_path in attachments:
                            if os.path.exists(att_path):
                                try:
                                    mail_item.Attachments.Add(att_path)
                                except Exception as att_e:
                                    logger.warning(f"Could not add attachment {att_path}: {att_e}")
                    
                    # Send - ensure sender identity is still bound right before send
                    EmailService._apply_sender_to_mail_item(mail_item, account_object, sender_email)
                    logger.info(f"\n>>> SENDING EMAIL <<<")
                    logger.info(f"From: {sender_email}")
                    logger.info(f"To: {mail_item.To}")
                    logger.info(f"Subject: {mail_item.Subject[:60]}..." if len(mail_item.Subject) > 60 else f"Subject: {mail_item.Subject}")
                    logger.info(f"Body preview: {mail_item.HTMLBody[:100]}..." if len(mail_item.HTMLBody) > 100 else f"Body: {mail_item.HTMLBody}")
                    logger.info(f"Account used: {account_object}")
                    mail_item.Send()
                    logger.info(f">>> EMAIL SENT <<<\n")
                    
                    logger.debug(f"[OK] Email {i}: Sent to {'; '.join(recipient_emails)}")
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
            
            logger.info(f"[OK] Email send complete: {message}")
            
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
