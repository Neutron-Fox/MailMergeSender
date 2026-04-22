#!/usr/bin/env python3
"""
Email Configuration Wizard - Standalone CLI Tool
Final product for interactive email extraction configuration
Run: python email_config.py
"""
import re
from typing import List
from dataclasses import dataclass


@dataclass
class EmailConfig:
    """Configuration for email extraction from data"""
    email_columns: List[int]
    multiple_emails_per_row: bool
    separator_chars: List[str]
    email_template: str
    custom_validators: List[str] = None


class EmailConfigurationWizard:
    """Interactive CLI wizard for configuring email extraction"""
    
    def __init__(self):
        self.config = None
        self.sample_data = []
        
    def display_intro(self):
        """Display welcome screen"""
        print("\n" + "="*70)
        print("EMAIL EXTRACTION CONFIGURATION WIZARD")
        print("="*70)
        print("\nThis wizard will help you configure how the program")
        print("extracts and validates email addresses from your data.\n")
    
    def load_sample_data(self) -> List[List[str]]:
        """Load sample data for demonstration"""
        sample = [
            ["John", "john@example.com", "Manager"],
            ["Jane", "jane.doe@company.org", "Developer"],
            ["Bob", "bob@test.co.uk; alt_bob@test.co.uk", "Support"],
            ["Alice", "alice@domain.com,alice.backup@domain.com", "Admin"],
            ["Charlie", "charlie@site.net", "Designer"],
        ]
        return sample
    
    def display_sample_data(self, data: List[List[str]]):
        """Display sample data with column indices"""
        print("\n" + "-"*70)
        print("SAMPLE DATA PREVIEW:")
        print("-"*70)
        
        for i in range(len(data[0])):
            print(f"Column {i}: ", end="")
        print()
        
        for row_idx, row in enumerate(data):
            print(f"Row {row_idx}: ", end="")
            for col_idx, cell in enumerate(row):
                print(f"[{cell:25}] ", end="")
            print()
        print()
    
    def step_1_select_columns(self) -> List[int]:
        """Step 1: Select which columns contain email addresses"""
        print("\n" + "="*70)
        print("STEP 1: SELECT EMAIL COLUMNS")
        print("="*70)
        print("\nLook at the sample data above.")
        print("Which column(s) contain email addresses?")
        print("\nOptions:")
        print("  - Column 0: Names")
        print("  - Column 1: Emails")
        print("  - Column 2: Roles")
        
        print("\nEnter column number(s) separated by commas (e.g., '1' or '1,2')")
        user_input = input("Your choice: ").strip()
        
        try:
            columns = [int(x.strip()) for x in user_input.split(",")]
            print(f"\n✓ Selected columns: {columns}")
            return columns
        except ValueError:
            print("Invalid input. Defaulting to column 1.")
            return [1]
    
    def step_2_multiple_emails_per_row(self) -> bool:
        """Step 2: Ask if multiple emails can be in one row"""
        print("\n" + "="*70)
        print("STEP 2: MULTIPLE EMAILS PER ROW")
        print("="*70)
        print("\nCan a single row contain MULTIPLE email addresses?")
        print("\nExample:")
        print("  - YES: 'john@example.com; jane@example.com' (in one cell)")
        print("  - NO:  Only one email per row")
        
        choice = input("\nMultiple emails per row? (yes/no): ").strip().lower()
        
        result = choice in ['yes', 'y', '1', 'true']
        print(f"\n✓ Multiple emails per row: {result}")
        return result
    
    def step_3_select_separators(self, multiple_emails: bool) -> List[str]:
        """Step 3: Select separator characters"""
        print("\n" + "="*70)
        print("STEP 3: EMAIL SEPARATORS")
        print("="*70)
        
        if not multiple_emails:
            print("\nSince multiple emails per row is disabled,")
            print("separators won't be used.")
            print("\n✓ No separators needed")
            return []
        
        print("\nWhat character(s) separate multiple emails in a single cell?")
        print("\nCommon separators:")
        print("  - Semicolon (;)")
        print("  - Comma (,)")
        print("  - Pipe (|)")
        print("  - Space ( )")
        print("\nExample in data:")
        print("  'john@example.com; jane@example.com' -> separator is ';'")
        print("  'john@example.com,jane@example.com' -> separator is ','")
        
        print("\nEnter separator(s) (e.g., ';' or ';,|' for multiple)")
        user_input = input("Your separators: ").strip()
        
        if not user_input:
            separators = [',', ';']
            print(f"Using default separators: {separators}")
        else:
            separators = list(user_input)
            print(f"\n✓ Separators: {separators}")
        
        return separators
    
    def step_4_email_template(self) -> str:
        """Step 4: Define/confirm email template"""
        print("\n" + "="*70)
        print("STEP 4: EMAIL VALIDATION TEMPLATE")
        print("="*70)
        print("\nThe program needs to know what constitutes a valid email.")
        print("\nStandard email format: username@domain.extension")
        print("Examples:")
        print("  - john@example.com")
        print("  - jane.doe@company.co.uk")
        print("  - support+tag@domain.org")
        
        print("\nThe regex pattern for validating emails:")
        print("  ^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}$")
        
        print("\nUse this standard pattern? (recommended)")
        choice = input("Use standard email pattern? (yes/no): ").strip().lower()
        
        if choice in ['yes', 'y', '1', 'true', '']:
            template = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
            print("\n✓ Using standard email regex pattern")
        else:
            print("\nEnter custom regex pattern (or press Enter to use standard):")
            custom = input("Custom pattern: ").strip()
            template = custom if custom else r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
        
        return template
    
    def step_5_review_and_test(self, config: EmailConfig) -> bool:
        """Step 5: Review configuration and test"""
        print("\n" + "="*70)
        print("STEP 5: REVIEW CONFIGURATION")
        print("="*70)
        print("\nYour Email Extraction Configuration:")
        print(f"  Email Columns:        {config.email_columns}")
        print(f"  Multiple per Row:     {config.multiple_emails_per_row}")
        print(f"  Separators:           {config.separator_chars if config.separator_chars else 'None'}")
        print(f"  Validation Template:  {config.email_template[:50]}...")
        
        print("\n" + "-"*70)
        print("EXTRACTION TEST:")
        print("-"*70)
        
        extracted_emails = self.extract_emails_demo(config)
        print(f"\nFound {len(extracted_emails)} email addresses:")
        for email in extracted_emails:
            print(f"  ✓ {email}")
        
        print("\nDoes this look correct?")
        confirm = input("Proceed with this configuration? (yes/no): ").strip().lower()
        
        return confirm in ['yes', 'y', '1', 'true', '']
    
    def extract_emails_demo(self, config: EmailConfig) -> List[str]:
        """Demo email extraction using configuration"""
        emails = []
        pattern = re.compile(config.email_template)
        
        for row in self.sample_data:
            for col_idx in config.email_columns:
                if col_idx < len(row):
                    cell_content = row[col_idx]
                    
                    if config.multiple_emails_per_row and config.separator_chars:
                        # Split by separators
                        parts = [cell_content]
                        for separator in config.separator_chars:
                            new_parts = []
                            for part in parts:
                                new_parts.extend(part.split(separator))
                            parts = new_parts
                        
                        # Validate and clean each part
                        for part in parts:
                            email = part.strip()
                            if pattern.match(email):
                                emails.append(email)
                    else:
                        # Single email per row
                        email = cell_content.strip()
                        if pattern.match(email):
                            emails.append(email)
        
        return emails
    
    def run_wizard(self):
        """Run the complete wizard"""
        self.display_intro()
        
        # Load sample data
        self.sample_data = self.load_sample_data()
        self.display_sample_data(self.sample_data)
        
        # Run through steps
        email_columns = self.step_1_select_columns()
        multiple_emails = self.step_2_multiple_emails_per_row()
        separators = self.step_3_select_separators(multiple_emails)
        email_template = self.step_4_email_template()
        
        # Create configuration
        self.config = EmailConfig(
            email_columns=email_columns,
            multiple_emails_per_row=multiple_emails,
            separator_chars=separators,
            email_template=email_template,
            custom_validators=[]
        )
        
        # Review and confirm
        if self.step_5_review_and_test(self.config):
            self.display_final_summary()
            return self.config
        else:
            print("\nConfiguration cancelled. Please start over.")
            return None
    
    def display_final_summary(self):
        """Display final configuration summary"""
        print("\n" + "="*70)
        print("CONFIGURATION COMPLETE ✓")
        print("="*70)
        print("\nYour email extraction settings have been saved:")
        print(f"\n  📧 Email Columns:           {self.config.email_columns}")
        print(f"  🔀 Multiple Emails/Row:     {self.config.multiple_emails_per_row}")
        
        if self.config.separator_chars:
            print(f"  🔧 Separators:              {', '.join(repr(s) for s in self.config.separator_chars)}")
        
        print(f"  ✓ Validation Pattern:       Active")
        print("\n" + "-"*70)
        print("The program will now use these settings to extract and")
        print("validate email addresses from your imported data files.")
        print("-"*70 + "\n")


def main():
    """Main entry point"""
    print("\n\n")
    print("╔" + "="*68 + "╗")
    print("║" + " "*10 + "EMAIL EXTRACTION CONFIGURATION WIZARD" + " "*23 + "║")
    print("║" + " "*15 + "MailMergeSender v2.0" + " "*33 + "║")
    print("╚" + "="*68 + "╝")
    
    wizard = EmailConfigurationWizard()
    config = wizard.run_wizard()
    
    if config:
        print("\n📝 Configuration Object Created:")
        print(f"   {config}")
        print("\nThis configuration can be used for")
        print("extracting email addresses from imported data files.")
    
    print("\nWizard completed!\n")


if __name__ == "__main__":
    main()
