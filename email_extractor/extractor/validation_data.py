"""
Validation Data Loader
Loads Excel files containing validation rules for lead processing.
"""

import os
import pandas as pd
from typing import Set, Dict

class ValidationDataLoader:
    """Load and manage validation data from Excel files."""
    
    def __init__(self, data_folder: str = "validation_data"):
        """Initialize loader with data folder path.
        
        Args:
            data_folder: Path to folder containing validation Excel files
        """
        self.data_folder = data_folder
        self.academic_domains: Set[str] = set()
        self.excluded_domains: Set[str] = set()
        self.direct_accounts: Set[str] = set()
        self.blacklisted_countries: Set[str] = set()
        self.freemail_domains: Set[str] = set()
        
        self.load_all()
    
    def load_all(self):
        """Load all validation files."""
        try:
            self._load_file("academic_domains.xlsx", self.academic_domains)
            self._load_file("excluded_domains.xlsx", self.excluded_domains)
            self._load_file("direct_accounts.xlsx", self.direct_accounts)
            self._load_file("blacklisted_countries.xlsx", self.blacklisted_countries)
            self._load_file("freemail_domains.xlsx", self.freemail_domains)
            
            print(f"✓ Loaded validation data:")
            print(f"  - Academic domains: {len(self.academic_domains)}")
            print(f"  - Excluded domains: {len(self.excluded_domains)}")
            print(f"  - Direct accounts: {len(self.direct_accounts)}")
            print(f"  - Blacklisted countries: {len(self.blacklisted_countries)}")
            print(f"  - Freemail domains: {len(self.freemail_domains)}")
        except Exception as e:
            print(f"⚠ Warning: Could not load all validation data: {e}")
            print(f"  Continuing with available data...")
    
    def _load_file(self, filename: str, target_set: Set[str]):
        """Load a single Excel file into a set.
        
        Args:
            filename: Name of Excel file
            target_set: Set to populate with values
        """
        filepath = os.path.join(self.data_folder, filename)
        
        if not os.path.exists(filepath):
            print(f"  ⊘ {filename} not found, skipping...")
            return
        
        try:
            df = pd.read_excel(filepath)
            
            # Look for 'Option Values' column
            if 'Option Values' in df.columns:
                values = df['Option Values'].dropna().astype(str).str.strip().str.lower()
                target_set.update(values)
            else:
                print(f"  ⊘ {filename}: 'Option Values' column not found")
        except Exception as e:
            print(f"  ⊘ Error loading {filename}: {e}")
    
    def is_academic_domain(self, domain: str) -> bool:
        """Check if domain is in academic list."""
        return domain.lower().strip() in self.academic_domains
    
    def is_excluded_domain(self, domain: str) -> bool:
        """Check if domain is in excluded list."""
        return domain.lower().strip() in self.excluded_domains
    
    def is_direct_account(self, company: str) -> bool:
        """Check if company is a direct account."""
        return company.lower().strip() in self.direct_accounts
    
    def is_blacklisted_country(self, country: str) -> bool:
        """Check if country is blacklisted."""
        return country.lower().strip() in self.blacklisted_countries
    
    def is_freemail_domain(self, domain: str) -> bool:
        """Check if domain is a freemail provider."""
        return domain.lower().strip() in self.freemail_domains
    
    def validate_lead(self, company: str, country: str, email: str) -> Dict[str, str]:
        """Comprehensive lead validation.
        
        Args:
            company: Company name
            country: Country name
            email: Email address
            
        Returns:
            dict with keys:
                - is_valid: bool
                - reason: str (if invalid)
                - validation_type: str
        """
        # Extract domain from email
        domain = ""
        if email and "@" in email:
            domain = email.split("@")[-1].lower().strip()
        
        # Check blacklisted country
        if self.is_blacklisted_country(country):
            return {
                "is_valid": False,
                "reason": f"Blacklisted country: {country}",
                "validation_type": "Country"
            }
        
        # Check direct account
        if self.is_direct_account(company):
            return {
                "is_valid": False,
                "reason": f"Direct account: {company}",
                "validation_type": "Direct Account"
            }
        
        # Check excluded domain
        if self.is_excluded_domain(domain):
            return {
                "is_valid": False,
                "reason": f"Excluded domain: {domain}",
                "validation_type": "Excluded Domain"
            }
        
        # Check academic domain
        if self.is_academic_domain(domain):
            return {
                "is_valid": False,
                "reason": f"Academic domain: {domain}",
                "validation_type": "Academic"
            }
        
        # Check freemail
        if self.is_freemail_domain(domain):
            return {
                "is_valid": True,  # Valid but flagged
                "reason": f"Freemail provider: {domain}",
                "validation_type": "Freemail"
            }
        
        # All checks passed
        return {
            "is_valid": True,
            "reason": "",
            "validation_type": "Valid"
        }